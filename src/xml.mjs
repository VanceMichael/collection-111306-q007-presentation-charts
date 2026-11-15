import {PackageError} from './zip.mjs';

const MAX_DEPTH = 256;
const MAX_XML_BYTES = 64 * 1024 * 1024;

export class XmlError extends Error {
  constructor(message, {part} = {}) {
    super(part ? `${part}: ${message}` : message);
    this.name = 'XmlError';
    this.part = part;
  }
}

const entityPattern = /&(#x[0-9a-fA-F]+|#[0-9]+|amp|lt|gt|quot|apos);/g;
const namedEntities = {amp: '&', lt: '<', gt: '>', quot: '"', apos: "'"};

export function decodeEntities(text) {
  return text.replace(entityPattern, (match, body) => {
    if (body[0] === '#') {
      const code = body[1] === 'x' ? Number.parseInt(body.slice(2), 16) : Number.parseInt(body.slice(1), 10);
      if (!Number.isInteger(code) || code < 0 || code > 0x10ffff || (code >= 0xd800 && code <= 0xdfff)) {
        throw new XmlError(`非法字符引用 ${match}`);
      }
      return String.fromCodePoint(code);
    }
    const value = namedEntities[body];
    if (value === undefined) throw new XmlError(`未知实体 ${match}`);
    return value;
  });
}

function parseAttributes(text, part) {
  const attributes = {};
  const pattern = /([^\s=/>]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
  let match;
  let consumed = 0;
  while ((match = pattern.exec(text))) {
    const gap = text.slice(consumed, match.index).trim();
    if (gap) throw new XmlError('属性语法错误', {part});
    consumed = pattern.lastIndex;
    const name = match[1];
    if (Object.hasOwn(attributes, name)) throw new XmlError(`重复属性 ${name}`, {part});
    attributes[name] = decodeEntities(match[3] ?? match[4]);
  }
  if (text.slice(consumed).trim()) throw new XmlError('属性语法错误', {part});
  return attributes;
}

/**
 * A deliberately small XML reader for OOXML parts. It understands elements,
 * attributes, character data and the five XML entities, and refuses DOCTYPE
 * declarations outright. Scripts, macros or external DTDs therefore have no
 * execution or resolution path through this module.
 */
export function parseXml(source, {part} = {}) {
  const text = Buffer.isBuffer(source) ? source.toString('utf8') : String(source);
  if (Buffer.byteLength(text, 'utf8') > MAX_XML_BYTES) throw new XmlError('XML 过大', {part});
  if (/<!DOCTYPE/i.test(text)) throw new XmlError('拒绝含 DOCTYPE 的 XML', {part});
  const root = {name: '#document', attributes: {}, children: [], text: ''};
  const stack = [root];
  const length = text.length;
  let index = 0;
  while (index < length) {
    const open = text.indexOf('<', index);
    if (open < 0) {
      appendText(stack, text.slice(index));
      break;
    }
    if (open > index) appendText(stack, text.slice(index, open));
    if (text.startsWith('<!--', open)) {
      const close = text.indexOf('-->', open + 4);
      if (close < 0) throw new XmlError('注释未闭合', {part});
      index = close + 3;
      continue;
    }
    if (text.startsWith('<![CDATA[', open)) {
      const close = text.indexOf(']]>', open + 9);
      if (close < 0) throw new XmlError('CDATA 未闭合', {part});
      stack[stack.length - 1].text += text.slice(open + 9, close);
      index = close + 3;
      continue;
    }
    if (text.startsWith('<?', open)) {
      const close = text.indexOf('?>', open + 2);
      if (close < 0) throw new XmlError('处理指令未闭合', {part});
      index = close + 2;
      continue;
    }
    if (text.startsWith('<!', open)) throw new XmlError('不支持的 XML 声明', {part});
    const close = findTagEnd(text, open + 1);
    if (close < 0) throw new XmlError('标签未闭合', {part});
    let body = text.slice(open + 1, close);
    index = close + 1;
    if (body.startsWith('/')) {
      const name = body.slice(1).trim();
      if (stack.length < 2) throw new XmlError(`多余的结束标签 ${name}`, {part});
      const current = stack.pop();
      if (current.name !== name) throw new XmlError(`结束标签 ${name} 与 ${current.name} 不匹配`, {part});
      continue;
    }
    let selfClosing = false;
    if (body.endsWith('/')) {
      selfClosing = true;
      body = body.slice(0, -1);
    }
    const separator = body.search(/[\s]/);
    const name = (separator < 0 ? body : body.slice(0, separator)).trim();
    if (!name || /[<>"'=&]/.test(name)) throw new XmlError(`非法元素名 ${name}`, {part});
    const attributes = separator < 0 ? {} : parseAttributes(body.slice(separator + 1), part);
    const element = {name, attributes, children: [], text: ''};
    stack[stack.length - 1].children.push(element);
    if (!selfClosing) {
      stack.push(element);
      if (stack.length > MAX_DEPTH) throw new XmlError('XML 嵌套过深', {part});
    }
  }
  if (stack.length !== 1) throw new XmlError(`元素未闭合 ${stack[stack.length - 1].name}`, {part});
  if (root.children.length !== 1) throw new XmlError('XML 必须有且只有一个根元素', {part});
  return root.children[0];
}

// Whitespace is structural between elements, but inside value-bearing elements
// such as <t>, <v> or <f> it belongs to the value (xml:space="preserve").
const TEXT_BEARING = new Set(['t', 'v', 'f', 'formatCode']);

function appendText(stack, raw) {
  if (!raw) return;
  const current = stack[stack.length - 1];
  const name = localName(current.name);
  if (!raw.trim() && !TEXT_BEARING.has(name)) return;
  current.text += raw.trim() ? decodeEntities(raw) : raw;
}

function findTagEnd(text, start) {
  let quote = null;
  for (let index = start; index < text.length; index++) {
    const character = text[index];
    if (quote) {
      if (character === quote) quote = null;
    } else if (character === '"' || character === "'") quote = character;
    else if (character === '<') return -1;
    else if (character === '>') return index;
  }
  return -1;
}

export function localName(name) {
  const colon = name.indexOf(':');
  return colon < 0 ? name : name.slice(colon + 1);
}

export function elements(node, name) {
  const wanted = localName(name);
  return node.children.filter(child => localName(child.name) === wanted);
}

export function firstElement(node, name) {
  return node.children.find(child => localName(child.name) === localName(name)) ?? null;
}

export function descendants(node, name, output = []) {
  const wanted = localName(name);
  for (const child of node.children) {
    if (localName(child.name) === wanted) output.push(child);
    descendants(child, wanted, output);
  }
  return output;
}

export function textOf(node) {
  if (!node) return null;
  let value = node.text;
  for (const child of node.children) value += textOf(child);
  return value;
}

export function attr(node, name) {
  if (!node) return null;
  if (Object.hasOwn(node.attributes, name)) return node.attributes[name];
  const wanted = localName(name);
  for (const [key, value] of Object.entries(node.attributes)) if (localName(key) === wanted) return value;
  return null;
}
