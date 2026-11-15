import {posix} from 'node:path';
import {PackageError, ZipArchive} from './zip.mjs';
import {parseXml, attr, elements, localName} from './xml.mjs';

const RELATIONSHIP_OFFICE_DOCUMENT = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';

export function isExternalTargetMode(value) {
  return typeof value === 'string' && value.toLowerCase() === 'external';
}

export function relationshipsPartFor(partName) {
  if (partName === '' || partName === null || partName === undefined) return '_rels/.rels';
  const slash = partName.lastIndexOf('/');
  const directory = slash < 0 ? '' : partName.slice(0, slash + 1);
  const base = slash < 0 ? partName : partName.slice(slash + 1);
  return `${directory}_rels/${base}.rels`;
}

export function resolveRelationshipTarget(sourcePart, target) {
  if (typeof target !== 'string' || target.trim() === '') throw new PackageError('关系目标为空');
  if (target.startsWith('#')) throw new PackageError('关系目标是内部锚点，不是资料');
  if (/^[a-zA-Z][a-zA-Z0-9+.-]*:/.test(target)) throw new PackageError(`关系目标不是包内路径: ${target}`);
  const normalized = target.startsWith('/')
    ? posix.normalize(target.slice(1))
    : posix.normalize(posix.join(sourcePart ? posix.dirname(sourcePart) : '.', target));
  if (normalized === '..' || normalized.startsWith('../') || posix.isAbsolute(normalized)) {
    throw new PackageError(`关系目标越出包边界: ${target}`);
  }
  return normalized.replace(/^\.\//, '');
}

export class OpcPackage {
  constructor(input, {label = '演示文稿'} = {}) {
    this.label = label;
    this.archive = new ZipArchive(input);
    this._xml = new Map();
    this._rels = new Map();
    this._contentTypes = null;
  }

  hasPart(name) { return this.archive.has(name); }

  rawPart(name) {
    try {
      return {ok: true, bytes: this.archive.read(name)};
    } catch (error) {
      if (!this.archive.has(name)) return {ok: false, status: 'missing', message: '包内缺少该资料', part: name};
      return {ok: false, status: 'corrupt', message: error.message, part: name};
    }
  }

  xmlPart(name) {
    if (this._xml.has(name)) {
      const cached = this._xml.get(name);
      if (cached instanceof Error) throw cached;
      return cached;
    }
    let value;
    try {
      const raw = this.rawPart(name);
      if (!raw.ok) throw new PackageError(raw.message, {part: name});
      value = parseXml(raw.bytes, {part: name});
    } catch (error) {
      this._xml.set(name, error);
      throw error;
    }
    this._xml.set(name, value);
    return value;
  }

  /**
   * Relationship list for a part. Returns an empty list when no .rels part
   * exists (a part may legitimately have none) and throws when the existing
   * relationship part cannot be read or parsed.
   */
  relationships(partName) {
    const relsPart = relationshipsPartFor(partName);
    if (this._rels.has(relsPart)) {
      const cached = this._rels.get(relsPart);
      if (cached instanceof Error) throw cached;
      return cached;
    }
    let value;
    try {
      if (!this.archive.has(relsPart)) value = [];
      else {
        const root = this.xmlPart(relsPart);
        if (localName(root.name) !== 'Relationships') throw new PackageError('关系资料根元素不是 Relationships', {part: relsPart});
        value = elements(root, 'Relationship').map(node => {
          const targetMode = attr(node, 'TargetMode');
          const target = attr(node, 'Target');
          const type = attr(node, 'Type');
          const id = attr(node, 'Id');
          const external = isExternalTargetMode(targetMode);
          const relationship = {id, type, target, targetMode: targetMode ?? 'Internal', external, resolved: null, resolutionError: null};
          if (!id || !type || target === null) throw new PackageError('关系缺少 Id、Type 或 Target', {part: relsPart});
          if (!external) {
            try { relationship.resolved = resolveRelationshipTarget(partName, target); }
            catch (error) { relationship.resolutionError = error.message; }
          }
          return relationship;
        });
        const seen = new Set();
        for (const relationship of value) {
          if (seen.has(relationship.id)) throw new PackageError(`关系 Id 重复: ${relationship.id}`, {part: relsPart});
          seen.add(relationship.id);
        }
      }
    } catch (error) {
      this._rels.set(relsPart, error);
      throw error;
    }
    this._rels.set(relsPart, value);
    return value;
  }

  relationship(partName, id) {
    return this.relationships(partName).find(value => value.id === id) ?? null;
  }

  contentTypes() {
    if (this._contentTypes) return this._contentTypes;
    const defaults = new Map();
    const overrides = new Map();
    const root = this.xmlPart('[Content_Types].xml');
    for (const node of elements(root, 'Default')) {
      const extension = attr(node, 'Extension');
      const contentType = attr(node, 'ContentType');
      if (extension && contentType) defaults.set(extension.toLowerCase(), contentType);
    }
    for (const node of elements(root, 'Override')) {
      const partName = attr(node, 'PartName');
      const contentType = attr(node, 'ContentType');
      if (partName && contentType) overrides.set(partName.replace(/^\//, ''), contentType);
    }
    this._contentTypes = {
      defaults,
      overrides,
      contentTypeOf(name) {
        if (overrides.has(name)) return overrides.get(name);
        const extension = name.includes('.') ? name.slice(name.lastIndexOf('.') + 1).toLowerCase() : '';
        return defaults.get(extension) ?? null;
      },
    };
    return this._contentTypes;
  }

  mainDocumentPart() {
    const relationship = this.relationships('').find(value => value.type === RELATIONSHIP_OFFICE_DOCUMENT) ?? null;
    if (!relationship || !relationship.resolved) return null;
    return relationship.resolved;
  }
}

export function relationshipTypeSuffix(type) {
  if (typeof type !== 'string') return '';
  return type.slice(type.lastIndexOf('/') + 1);
}
