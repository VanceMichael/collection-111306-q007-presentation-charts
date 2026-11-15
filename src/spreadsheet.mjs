import {PackageError} from './zip.mjs';
import {attr, elements, firstElement, localName, textOf} from './xml.mjs';

const REL_SHEET = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet';
const REL_SHARED_STRINGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings';

const columnPattern = /^([A-Z]+)([0-9]+)$/i;

export function splitCellReference(reference) {
  const match = columnPattern.exec(String(reference ?? '').trim());
  if (!match) return null;
  return {column: match[1].toUpperCase(), row: Number(match[2])};
}

export function columnIndex(letters) {
  let value = 0;
  for (const character of letters.toUpperCase()) value = value * 26 + (character.charCodeAt(0) - 64);
  return value;
}

export function columnLetters(index) {
  let value = index;
  let output = '';
  while (value > 0) {
    const remainder = (value - 1) % 26;
    output = String.fromCharCode(65 + remainder) + output;
    value = Math.floor((value - 1) / 26);
  }
  return output;
}

function parseRange(reference) {
  const body = String(reference ?? '').replace(/\$/g, '').trim();
  const parts = body.split(':');
  if (parts.length > 2) return null;
  const start = splitCellReference(parts[0]);
  const end = parts.length === 2 ? splitCellReference(parts[1]) : start;
  if (!start || !end) return null;
  return {
    start,
    end,
    columns: [Math.min(columnIndex(start.column), columnIndex(end.column)), Math.max(columnIndex(start.column), columnIndex(end.column))],
    rows: [Math.min(start.row, end.row), Math.max(start.row, end.row)],
  };
}

/**
 * Inline (or cached) cell values as Excel would evaluate them. Strings are
 * stored without a type, numbers carry t="n". Shared strings are resolved via
 * the workbook relationship, never guessed.
 */
export class Workbook {
  constructor(pkg, partName) {
    this.pkg = pkg;
    this.partName = partName;
    this.part = partName;
    this.status = 'read';
    this._sheets = null;
    this._sharedStrings = undefined;
    this._grids = new Map();
    this.diagnostics = [];
  }

  static open(pkg, partName) {
    const workbook = new Workbook(pkg, partName);
    workbook.workbookXml = pkg.xmlPart(partName);
    return workbook;
  }

  get sheets() {
    if (this._sheets) return this._sheets;
    const root = this.workbookXml;
    if (localName(root.name) !== 'workbook') throw new PackageError('工作簿根元素不是 workbook', {part: this.partName});
    const relationships = this.pkg.relationships(this.partName);
    const sheets = [];
    const workbookPr = firstElement(root, 'workbookPr');
    const date1904 = workbookPr ? attr(workbookPr, 'date1904') === '1' || attr(workbookPr, 'date1904') === 'true' : false;
    for (const node of elements(firstElement(root, 'sheets') ?? root, 'sheet')) {
      const name = attr(node, 'name');
      const id = attr(node, 'r:id') ?? attr(node, 'id');
      const relationship = id ? relationships.find(value => value.id === id) ?? null : null;
      sheets.push({
        name,
        sheetId: attr(node, 'sheetId'),
        state: attr(node, 'state') ?? 'visible',
        relationshipId: id,
        part: relationship?.resolved ?? null,
        problem: relationship ? (relationship.resolutionError ?? (relationship.external ? '工作表关系指向外部资料' : null)) : (id ? '缺少工作表关系' : '工作表缺少关系 Id'),
      });
    }
    this._sheets = {sheets, date1904};
    return this._sheets;
  }

  sheet(name) {
    return this.sheets.sheets.find(value => value.name === name) ?? null;
  }

  sheetByName(name) {
    return this.sheet(name);
  }

  resolveSheet(name) {
    const exact = this.sheet(name);
    if (exact) return {sheet: exact, warning: null};
    const folded = this.sheets.sheets.filter(sheet => String(sheet.name ?? '').toLowerCase() === String(name ?? '').toLowerCase());
    if (folded.length === 1) {
      return {sheet: folded[0], warning: {code: 'sheetNameCaseMismatch', message: `引用工作表 ${name}，工作簿中实际为 ${folded[0].name}`, actualName: folded[0].name}};
    }
    return {sheet: null, warning: null};
  }

  grid(name, {refresh = false} = {}) {
    const sheet = typeof name === 'string' ? this.sheet(name) : name;
    if (!sheet) throw new PackageError(`工作簿没有工作表: ${name}`);
    if (!refresh && this._grids.has(sheet.name)) return this._grids.get(sheet.name);
    const grid = this.readSheet(sheet);
    this._grids.set(sheet.name, grid);
    return grid;
  }

  sheetByIndex(index) {
    return this.sheets.sheets[index] ?? null;
  }

  sharedStrings() {
    if (this._sharedStrings !== undefined) return this._sharedStrings;
    const relationships = this.pkg.relationships(this.partName).filter(value => value.type === REL_SHARED_STRINGS);
    if (relationships.length === 0) { this._sharedStrings = []; return this._sharedStrings; }
    if (relationships.length > 1) {
      this.diagnostics.push({code: 'sharedStringsAmbiguous', message: '工作簿存在多个共享字符串关系，无法确定引用'});
      this._sharedStrings = [];
      return this._sharedStrings;
    }
    const relationship = relationships[0];
    if (!relationship.resolved) {
      this.diagnostics.push({code: 'sharedStringsUnavailable', message: relationship.resolutionError ?? '共享字符串关系不可解析'});
      this._sharedStrings = [];
      return this._sharedStrings;
    }
    try {
      const root = this.pkg.xmlPart(relationship.resolved);
      this._sharedStrings = elements(root, 'si').map(node => {
        const text = textOf(node);
        return text ?? '';
      });
    } catch (error) {
      this.diagnostics.push({code: 'sharedStringsCorrupt', message: error.message, part: relationship.resolved});
      this._sharedStrings = [];
    }
    return this._sharedStrings;
  }

  /**
   * Reads a worksheet into a sparse grid. Hidden columns are recorded rather
   * than silently ignored: formulas still evaluate them, but the collector can
   * report that a category came from a hidden column.
   */
  readSheet(sheet) {
    if (!sheet?.part) throw new PackageError(sheet?.problem ?? '工作表路径缺失');
    const root = this.pkg.xmlPart(sheet.part);
    if (localName(root.name) !== 'worksheet') throw new PackageError('工作表根元素不是 worksheet', {part: sheet.part});
    const hiddenColumns = new Set();
    const hiddenRows = new Set();
    const colsNode = firstElement(root, 'cols');
    if (colsNode) {
      for (const col of elements(colsNode, 'col')) {
        const hidden = attr(col, 'hidden') === '1' || attr(col, 'hidden') === 'true';
        if (!hidden) continue;
        const min = Number(attr(col, 'min'));
        const max = Number(attr(col, 'max'));
        if (!Number.isInteger(min) || !Number.isInteger(max) || max < min) continue;
        for (let index = min; index <= max; index++) hiddenColumns.add(index);
      }
    }
    const cells = new Map();
    const shared = this.sharedStrings();
    const sheetData = firstElement(root, 'sheetData');
    if (!sheetData) return {cells, hiddenColumns, hiddenRows, rowCount: 0};
    let maxRow = 0;
    for (const row of elements(sheetData, 'row')) {
      const rowNumber = Number(attr(row, 'r'));
      const rowHidden = attr(row, 'hidden') === '1' || attr(row, 'hidden') === 'true';
      if (rowHidden && Number.isInteger(rowNumber)) hiddenRows.add(rowNumber);
      for (const cell of elements(row, 'c')) {
        const reference = attr(cell, 'r');
        const split = splitCellReference(reference);
        if (!split) continue;
        if (!Number.isInteger(rowNumber) || split.row !== rowNumber) {
          this.diagnostics.push({code: 'cellRowMismatch', message: `单元格 ${reference} 与行号 ${rowNumber} 不一致`, part: sheet.part});
        }
        maxRow = Math.max(maxRow, split.row);
        const type = attr(cell, 't') ?? 'n';
        const formula = firstElement(cell, 'f');
        const valueNode = firstElement(cell, 'v');
        let raw = valueNode ? textOf(valueNode) : null;
        let value = null;
        let kind = 'number';
        if (type === 'inlineStr') {
          const inline = firstElement(cell, 'is');
          value = inline ? textOf(inline) : null;
          kind = 'string';
        } else if (type === 's') {
          kind = 'sharedString';
          const index = raw === null || raw === '' ? null : Number(raw);
          if (raw === null || raw === '') {
            this.diagnostics.push({code: 'emptySharedStringIndex', message: `单元格 ${reference} 缺少共享字符串索引`, part: sheet.part});
          } else if (!Number.isInteger(index) || index < 0) {
            this.diagnostics.push({code: 'badSharedStringIndex', message: `单元格 ${reference} 的共享字符串索引非法: ${raw}`, part: sheet.part});
          } else if (index >= shared.length) {
            this.diagnostics.push({code: 'missingSharedString', message: `单元格 ${reference} 引用了不存在的共享字符串 ${index}`, part: sheet.part});
          } else value = shared[index];
        } else if (type === 'str') {
          kind = 'formulaString';
          value = raw;
        } else if (type === 'b') {
          kind = 'boolean';
          value = raw === '1' || raw === 'true';
        } else if (type === 'e') {
          kind = 'error';
          value = raw;
        } else if (type === 'd') {
          kind = 'date';
          value = raw;
        } else {
          kind = 'number';
          value = raw === null || raw === '' ? null : Number(raw);
          if (value !== null && !Number.isFinite(value)) {
            this.diagnostics.push({code: 'badNumber', message: `单元格 ${reference} 不是有限数值: ${raw}`, part: sheet.part});
            value = null;
          }
        }
        if (raw === null && formula) {
          this.diagnostics.push({code: 'formulaWithoutCachedValue', message: `单元格 ${reference} 只有公式没有缓存值`, part: sheet.part});
          kind = 'emptyFormula';
        }
        cells.set(`${split.column}${split.row}`, {
          reference: `${split.column}${split.row}`,
          column: split.column,
          columnIndex: columnIndex(split.column),
          row: split.row,
          hiddenColumn: hiddenColumns.has(columnIndex(split.column)),
          hiddenRow: rowHidden,
          type: kind,
          value,
          raw,
          formula: formula ? textOf(formula) : null,
          sharedStringIndex: kind === 'sharedString' && Number.isInteger(Number(raw)) ? Number(raw) : null,
        });
      }
    }
    return {cells, hiddenColumns, hiddenRows, rowCount: maxRow};
  }
}

export function parseFormulaReference(formula) {
  const text = String(formula ?? '').trim();
  if (!text) return null;
  const bang = text.lastIndexOf('!');
  if (bang <= 0 || bang === text.length - 1) return {kind: 'unsupported', formula: text};
  const prefix = text.slice(0, bang);
  const rangeText = text.slice(bang + 1);
  let sheetName = prefix;
  let external = false;
  if (prefix.startsWith("'") && prefix.endsWith("'")) sheetName = prefix.slice(1, -1).replace(/''/g, "'");
  else if (/^\[[^\]]*\]/.test(prefix)) {
    external = true;
    sheetName = prefix.slice(prefix.indexOf(']') + 1);
  }
  if (sheetName.includes('[')) external = true;
  const range = parseRange(rangeText);
  if (!range) return {kind: 'unsupported', formula: text, sheetName};
  return {kind: 'range', formula: text, sheetName, external, range};
}
