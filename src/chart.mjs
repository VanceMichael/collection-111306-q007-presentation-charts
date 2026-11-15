import {elements, firstElement, localName, textOf, attr} from './xml.mjs';
import {columnIndex, columnLetters, parseFormulaReference} from './spreadsheet.mjs';

export function sameValue(left, right) {
  if (left === null || left === undefined || right === null || right === undefined) return false;
  if (typeof left === 'number' && typeof right === 'number') {
    if (!Number.isFinite(left) || !Number.isFinite(right)) return false;
    return Math.abs(left - right) <= Number.EPSILON * Math.max(1, Math.abs(left), Math.abs(right)) * 8;
  }
  if ((typeof left === 'number') !== (typeof right === 'number')) {
    const numericLeft = typeof left === 'number' ? left : Number(String(left).trim());
    const numericRight = typeof right === 'number' ? right : Number(String(right).trim());
    if (String(left).trim() !== '' && String(right).trim() !== '' && Number.isFinite(numericLeft) && Number.isFinite(numericRight)) {
      return numericLeft === numericRight;
    }
  }
  return String(left) === String(right);
}

function direct(node, name) {
  return firstElement(node, name);
}

function formatRange(range) {
  const start = `${range.start.column}${range.start.row}`;
  const end = `${range.end.column}${range.end.row}`;
  return start === end ? start : `${start}:${end}`;
}

function referenceOrientation(range) {
  const rows = range.rows[1] - range.rows[0] + 1;
  const columns = range.columns[1] - range.columns[0] + 1;
  if (rows === 1 && columns === 1) return {kind: 'single', size: 1, rows, columns};
  if (rows === 1) return {kind: 'horizontal', size: columns, rows, columns};
  if (columns === 1) return {kind: 'vertical', size: rows, rows, columns};
  return {kind: 'matrix', size: rows * columns, rows, columns};
}

function mappedCell(range, orientation, index) {
  if (index < 0 || index >= orientation.size) return null;
  if (orientation.kind === 'single') return {column: range.start.column, row: range.start.row};
  if (orientation.kind === 'horizontal') return {column: columnLetters(columnIndex(range.start.column) + index), row: range.start.row};
  if (orientation.kind === 'vertical') return {column: range.start.column, row: range.start.row + index};
  return null;
}

function issue(severity, code, message, source = null, extra = {}) {
  return {severity, code, message, ...(source ? {source} : {}), ...extra};
}

function readCache(cacheNode, pointer, issues, {numeric = false, literal = false} = {}) {
  if (!cacheNode) return null;
  const countNode = direct(cacheNode, 'ptCount');
  const declaredRaw = countNode ? attr(countNode, 'val') : attr(cacheNode, 'ptCount');
  let declaredPtCount = null;
  if (declaredRaw !== null && declaredRaw !== undefined && declaredRaw !== '') {
    const parsed = Number(declaredRaw);
    if (Number.isInteger(parsed) && parsed >= 0) declaredPtCount = parsed;
    else issues.push(issue('warning', 'badPointCount', `缓存 ptCount 非法: ${declaredRaw}`, {pointer}));
  }
  const formatCodeNode = direct(cacheNode, 'formatCode');
  const points = [];
  const seen = new Set();
  elements(cacheNode, 'pt').forEach((point, position) => {
    const rawIndex = attr(point, 'idx');
    let index = rawIndex === null || rawIndex === '' ? position : Number(rawIndex);
    if (!Number.isInteger(index) || index < 0) {
      issues.push(issue('warning', 'badPointIndex', `缓存点索引非法: ${rawIndex}，按出现顺序记录`, {pointer}));
      index = position;
    }
    if (seen.has(index)) issues.push(issue('blocking', 'duplicatePointIndex', `缓存点索引重复: ${index}`, {pointer}));
    seen.add(index);
    const valueNode = direct(point, 'v');
    const raw = valueNode ? textOf(valueNode) : null;
    let value = raw;
    if (numeric) {
      value = raw === null || String(raw).trim() === '' ? null : Number(raw);
      if (value !== null && !Number.isFinite(value)) {
        issues.push(issue('warning', 'nonNumericCacheValue', `数值缓存点 ${index} 不是有限数字: ${raw}`, {pointer: `${pointer}/pt[@idx=${index}]`}));
        value = null;
      }
    }
    points.push({idx: index, value, raw, source: {part: null, pointer: `${pointer}/pt[@idx=${index}]`}});
  });
  points.sort((left, right) => left.idx - right.idx);
  const missingIndexes = declaredPtCount === null ? [] : Array.from({length: declaredPtCount}, (_, index) => index).filter(index => !seen.has(index));
  return {
    declaredPtCount,
    formatCode: formatCodeNode ? textOf(formatCodeNode) : null,
    points,
    missingIndexes,
    literal,
    pointer,
  };
}

/**
 * Reads one chart data source exactly as it appears in chart XML. Absent cache
 * points stay absent and literal data stays literal; nothing is synthesized.
 */
export function readDataReference(node, {role, pointer}) {
  const empty = {
    role, sourceKind: 'none', formula: null, resolution: 'none', reference: null,
    cache: null, issues: [], isReference: false, literal: false,
  };
  if (!node) return empty;
  const kind = ['strRef', 'numRef', 'multiLvlStrRef', 'strLit', 'numLit'].find(name => direct(node, name));
  if (!kind) {
    return {...empty, sourceKind: 'unknown', resolution: 'unsupported', issues: [issue('blocking', 'unknownDataSource', `${role} 的数据来源类型无法识别`, {part: null, pointer})]};
  }
  const sourceNode = direct(node, kind);
  const isReference = kind.endsWith('Ref');
  const formulaNode = isReference ? direct(sourceNode, 'f') : null;
  const formula = formulaNode ? textOf(formulaNode).trim() : null;
  const issues = [];
  const numeric = kind === 'numRef' || kind === 'numLit';
  let cache = null;
  if (kind === 'multiLvlStrRef') {
    const container = direct(sourceNode, 'multiLvlStrCache');
    if (container) {
      const levels = elements(container, 'lvl').map((level, levelIndex) => {
        const levelCache = direct(level, 'strCache');
        const pointerForLevel = `${pointer}/multiLvlStrRef/multiLvlStrCache/lvl[${levelIndex}]/strCache`;
        return {level: levelIndex, ...readCache(levelCache, pointerForLevel, issues)};
      });
      const primary = levels.find(value => value.declaredPtCount !== null || value.points.length > 0) ?? levels[0] ?? null;
      if (primary) cache = {...primary, levels, pointer: `${pointer}/multiLvlStrRef/multiLvlStrCache`};
    }
  } else if (isReference) {
    const cacheName = kind === 'strRef' ? 'strCache' : 'numCache';
    cache = readCache(direct(sourceNode, cacheName), `${pointer}/${kind}/${cacheName}`, issues, {numeric});
  } else {
    cache = readCache(sourceNode, `${pointer}/${kind}`, issues, {numeric, literal: true});
  }
  if (isReference && !cache) issues.push(issue('blocking', 'missingCache', `${role} 引用没有显示缓存`, {part: null, pointer}));

  let resolution = isReference ? 'unsupported' : 'literal';
  let reference = null;
  let parsed = null;
  if (isReference) {
    parsed = parseFormulaReference(formula);
    if (parsed?.kind === 'range') {
      resolution = parsed.external ? 'external' : 'resolved';
      reference = {
        formula,
        sheetName: parsed.sheetName,
        range: formatRange(parsed.range),
        rangeDetail: parsed.range,
        orientation: referenceOrientation(parsed.range),
      };
    }
  }
  return {role, sourceKind: kind, formula, resolution, reference, cache, issues, isReference, literal: !isReference};
}

function maxCachedIndex(ref) {
  return (ref.cache?.points ?? []).reduce((maximum, point) => Math.max(maximum, point.idx + 1), 0);
}

function correlationShell(ref) {
  const shell = {
    role: ref.role,
    sourceKind: ref.sourceKind,
    formula: ref.formula,
    resolution: ref.resolution,
    reference: ref.reference ? {
      sheet: ref.reference.sheetName,
      range: ref.reference.range,
      orientation: ref.reference.orientation,
    } : null,
    display: ref.cache ? {
      declaredPtCount: ref.cache.declaredPtCount,
      formatCode: ref.cache.formatCode,
      points: ref.cache.points,
      missingIndexes: ref.cache.missingIndexes,
      pointer: ref.cache.pointer,
    } : null,
    workbook: {
      status: 'notRequested', part: null, sheet: null, sheetState: null,
      cells: [], hiddenColumns: [], hiddenRows: [], issues: [],
    },
    expectedCount: 0,
    lengths: {
      declaredPtCount: ref.cache?.declaredPtCount ?? null,
      cachedPoints: ref.cache?.points?.length ?? 0,
      referenceSize: ref.reference?.orientation?.size ?? null,
      workbookCells: 0,
    },
    points: [],
    gaps: [],
    conflicts: [],
    issues: [...(ref.issues ?? [])],
    warnings: [],
    blocking: [],
    counts: {match: 0, conflict: 0, displayOnly: 0, workbookOnly: 0, missing: 0, literal: 0, unmapped: 0},
  };
  for (const entry of ref.issues ?? []) {
    if (entry.severity === 'blocking') shell.blocking.push(entry);
    else shell.warnings.push(entry);
  }
  return shell;
}

function addIssue(correlation, value) {
  correlation.issues.push(value);
  if (value.severity === 'blocking') correlation.blocking.push(value);
  else correlation.warnings.push(value);
}

/**
 * Joins one display cache with the embedded workbook. Mapping is explicit:
 * single row/column ranges map by offset, while a multi-row/multi-column range
 * is reported as ambiguous and receives no cell mapping. Values are never
 * backfilled and gaps stay gaps.
 */
export function correlate(ref, workbook, {chartPart, label, workbookStatus = null}) {
  const result = correlationShell(ref);
  for (const point of result.display?.points ?? []) point.source.part = chartPart;
  const referenceSize = ref.reference?.orientation?.size ?? null;
  const sourceBase = {part: chartPart, pointer: ref.cache?.pointer ?? null};

  if (!ref.isReference) {
    result.workbook.status = ref.sourceKind === 'none' ? 'absent' : 'literal';
    if (ref.sourceKind !== 'none') addIssue(result, issue('warning', 'literalData', `${label} 使用字面量数据，没有可核对的单元格`, sourceBase));
  } else if (ref.resolution === 'external') {
    result.workbook.status = 'external';
    addIssue(result, issue('blocking', 'externalReference', `${label} 引用外部工作簿，未访问: ${ref.formula ?? '(无公式)'}`, sourceBase));
  } else if (ref.resolution === 'unsupported' || !ref.reference) {
    result.workbook.status = 'unresolved';
    addIssue(result, issue('blocking', 'unresolvedReference', `${label} 的引用无法解析为普通工作表区域: ${ref.formula ?? '(无公式)'}`, sourceBase));
  } else if (!workbook || workbook.status !== 'read') {
    result.workbook.status = workbook?.status ?? workbookStatus ?? 'unavailable';
    addIssue(result, issue('blocking', 'workbookUnavailable', `${label} 的嵌入工作簿不可用: ${workbook?.message ?? workbookStatus ?? '不存在'}`, {part: workbook?.part ?? null}));
  } else {
    result.workbook.part = workbook.part;
    result.workbook.sheet = ref.reference.sheetName;
    const resolvedSheet = workbook.resolveSheet(ref.reference.sheetName);
    if (resolvedSheet.warning) {
      addIssue(result, issue('warning', resolvedSheet.warning.code, resolvedSheet.warning.message, {part: workbook.part, sheet: resolvedSheet.warning.actualName ?? ref.reference.sheetName}));
    }
    if (!resolvedSheet.sheet) {
      result.workbook.status = 'sheetMissing';
      addIssue(result, issue('blocking', 'worksheetMissing', `${label} 引用的工作表不存在: ${ref.reference.sheetName}`, {part: workbook.part}));
    } else if (ref.reference.orientation.kind === 'matrix') {
      result.workbook.status = 'ambiguousRange';
      result.workbook.sheetState = resolvedSheet.sheet.state;
      addIssue(result, issue('blocking', 'ambiguousRange', `${label} 的引用 ${ref.reference.range} 覆盖多行多列，无法把缓存点逐点对应到单元格`, {part: workbook.part, sheet: resolvedSheet.sheet.name, range: ref.reference.range}));
    } else {
      let grid;
      try { grid = workbook.grid(resolvedSheet.sheet.name); }
      catch (error) {
        result.workbook.status = 'sheetUnreadable';
        addIssue(result, issue('blocking', 'worksheetUnreadable', `${label} 的工作表无法读取: ${error.message}`, {part: resolvedSheet.sheet.part ?? workbook.part}));
      }
      if (grid) {
        result.workbook.status = 'read';
        result.workbook.sheetState = resolvedSheet.sheet.state;
        if (resolvedSheet.sheet.state !== 'visible') {
          addIssue(result, issue('warning', 'worksheetHidden', `${label} 引用的工作表 ${resolvedSheet.sheet.name} 状态为 ${resolvedSheet.sheet.state}`, {part: workbook.part, sheet: resolvedSheet.sheet.name}));
        }
        const orientation = ref.reference.orientation;
        for (let index = 0; index < orientation.size; index++) {
          const position = mappedCell(ref.reference.rangeDetail, orientation, index);
          if (!position) continue;
          const cellName = `${position.column}${position.row}`;
          const cell = grid.cells.get(cellName) ?? null;
          const present = Boolean(cell && cell.value !== null && cell.value !== undefined);
          const hiddenColumn = grid.hiddenColumns.has(columnIndex(position.column));
          const hiddenRow = grid.hiddenRows.has(position.row);
          result.workbook.cells.push({
            idx: index, cell: cellName, present,
            value: cell ? cell.value : null,
            raw: cell ? cell.raw : null,
            type: cell ? cell.type : 'empty',
            formula: cell ? cell.formula : null,
            hiddenColumn, hiddenRow,
            sheetState: resolvedSheet.sheet.state,
            source: {part: workbook.part, sheet: resolvedSheet.sheet.name, cell: cellName},
          });
          if (hiddenColumn) addIssue(result, issue('warning', 'hiddenColumnSource', `${label} 的 ${cellName} 位于隐藏列 ${position.column}`, {part: workbook.part, sheet: resolvedSheet.sheet.name, cell: cellName}));
          if (hiddenRow) addIssue(result, issue('warning', 'hiddenRowSource', `${label} 的 ${cellName} 位于隐藏行 ${position.row}`, {part: workbook.part, sheet: resolvedSheet.sheet.name, cell: cellName}));
          if (cell?.formula && !present) addIssue(result, issue('blocking', 'formulaWithoutCachedValue', `${label} 的 ${cellName} 只有公式没有缓存值`, {part: workbook.part, sheet: resolvedSheet.sheet.name, cell: cellName}));
        }
        result.workbook.hiddenColumns = [...grid.hiddenColumns].sort((left, right) => left - right);
        result.workbook.hiddenRows = [...grid.hiddenRows].sort((left, right) => left - right);
        result.lengths.workbookCells = result.workbook.cells.filter(value => value.present).length;
        if (orientation.size !== (result.lengths.declaredPtCount ?? orientation.size)) {
          addIssue(result, issue('warning', 'lengthMismatch', `${label} 的引用范围有 ${orientation.size} 个位置，显示缓存声明 ${result.lengths.declaredPtCount ?? '未知'} 个`, {part: workbook.part, sheet: resolvedSheet.sheet.name, range: ref.reference.range}));
        }
      }
    }
  }

  result.expectedCount = Math.max(
    referenceSize ?? 0,
    result.lengths.declaredPtCount ?? 0,
    result.lengths.cachedPoints,
    maxCachedIndex(ref),
  );
  const displayByIndex = new Map((ref.cache?.points ?? []).map(point => [point.idx, point]));
  const cellByIndex = new Map(result.workbook.cells.map(cell => [cell.idx, cell]));
  if (result.workbook.status !== 'read' && result.workbook.status !== 'literal' && result.workbook.status !== 'absent') {
    // Nothing was mapped; points below are display-only evidence, never a joined value.
  }
  const mapped = result.workbook.status === 'read';
  for (let index = 0; index < result.expectedCount; index++) {
    const cachePoint = displayByIndex.get(index) ?? null;
    const cell = cellByIndex.get(index) ?? null;
    const cachePresent = Boolean(cachePoint && cachePoint.value !== null && cachePoint.value !== undefined);
    const cellPresent = Boolean(cell?.present);
    const withinReference = referenceSize === null || index < referenceSize;
    let status;
    if (ref.literal) status = cachePresent ? 'literal' : 'missing';
    else if (!mapped) status = 'unmapped';
    else if (cachePresent && cellPresent) status = sameValue(cachePoint.value, cell.value) ? 'match' : 'conflict';
    else if (cachePresent) status = 'displayOnly';
    else if (cellPresent) status = 'workbookOnly';
    else status = 'missing';
    if (!withinReference && (cachePresent || cellPresent)) {
      addIssue(result, issue('blocking', 'pointOutsideReference', `${label} 的第 ${index} 个点超出引用区域 ${ref.reference?.range ?? ''}`, {part: chartPart, pointer: cachePoint?.source?.pointer ?? null}, {index}));
    }
    const point = {
      index,
      withinReference,
      display: cachePoint ? {
        present: cachePresent, value: cachePoint.value, raw: cachePoint.raw,
        source: {part: chartPart, pointer: cachePoint.source.pointer},
      } : {present: false, value: null, raw: null, source: null},
      workbook: cell ? {
        present: cellPresent, value: cell.value, cell: cell.cell,
        sheet: result.workbook.sheet, part: result.workbook.part,
        type: cell.type, formula: cell.formula,
        hiddenColumn: cell.hiddenColumn, hiddenRow: cell.hiddenRow,
        source: {...cell.source},
      } : {present: false, value: null, cell: null, sheet: result.workbook.sheet, part: result.workbook.part, source: null},
      status,
    };
    result.points.push(point);
    if (status === 'conflict') result.conflicts.push({index, display: cachePoint.value, workbook: cell.value, displaySource: point.display.source, workbookSource: point.workbook.source});
    if (status === 'missing') result.gaps.push({index, missing: ['display', 'workbook']});
    else if (status === 'displayOnly') result.gaps.push({index, missing: ['workbook']});
    else if (status === 'workbookOnly') result.gaps.push({index, missing: ['display']});
    else if (status === 'conflict') result.gaps.push({index, missing: [], conflict: true});
    else if (status === 'unmapped' && ref.isReference) result.gaps.push({index, missing: [], unmapped: true});
  }
  if (result.lengths.declaredPtCount !== null && result.lengths.declaredPtCount !== result.lengths.cachedPoints) {
    addIssue(result, issue('warning', 'declaredCountMismatch', `${label} 的存储点声明 ${result.lengths.declaredPtCount} 个，实际只有 ${result.lengths.cachedPoints} 个缓存点`, sourceBase));
  }
  if (referenceSize !== null && result.lengths.declaredPtCount !== null && referenceSize !== result.lengths.declaredPtCount) {
    addIssue(result, issue('warning', 'referenceLengthMismatch', `${label} 的引用范围有 ${referenceSize} 个位置，缓存声明 ${result.lengths.declaredPtCount} 个`, sourceBase));
  }
  for (const key of Object.keys(result.counts)) result.counts[key] = result.points.filter(point => point.status === key).length;
  return result;
}

function readSeriesName(ser, pointer) {
  const tx = direct(ser, 'tx');
  if (!tx) return {kind: 'none', value: null, comparison: 'absent', cell: null, reference: null, formula: null, cache: null, workbook: null, issues: [], warnings: [], blocking: []};
  if (direct(tx, 'v')) return {kind: 'literal', value: textOf(direct(tx, 'v')), comparison: 'literal', cell: null, reference: null, formula: null, cache: null, workbook: null, issues: [], warnings: [], blocking: []};
  if (!direct(tx, 'strRef') && !direct(tx, 'strLit')) {
    return {kind: 'unknown', value: null, comparison: 'unresolved', cell: null, reference: null, formula: null, cache: null, workbook: null, issues: [issue('blocking', 'unknownSeriesName', '系列名称结构无法识别', {part: null, pointer: `${pointer}/tx`})], warnings: [], blocking: []};
  }
  const ref = readDataReference(tx, {role: 'name', pointer: `${pointer}/tx`});
  if (ref.cache && ref.cache.points.length > 1) {
    ref.issues.push(issue('warning', 'multiPointSeriesName', `系列名称来源包含 ${ref.cache.points.length} 个缓存点，只按第一个点核对`, {part: null, pointer: `${pointer}/tx`}));
  }
  return {kind: ref.isReference ? 'reference' : 'literal', ref, value: ref.cache?.points?.[0]?.value ?? null, issues: ref.issues};
}

function finishSeriesName(raw, workbook, chartPart, seriesIssues, workbookStatus = null) {
  if (raw.kind === 'reference') {
    const correlation = correlate(raw.ref, workbook, {chartPart, label: '系列名称', workbookStatus});
    const point = correlation.points[0] ?? null;
    for (const entry of correlation.issues) seriesIssues.push({...entry, seriesName: true});
    return {
      kind: 'reference',
      value: point?.display?.value ?? null,
      reference: correlation.reference,
      formula: correlation.formula,
      resolution: correlation.resolution,
      comparison: point?.status ?? 'missing',
      cell: point?.workbook?.cell ?? null,
      cellValue: point?.workbook?.value ?? null,
      hiddenColumn: Boolean(point?.workbook?.hiddenColumn),
      hiddenRow: Boolean(point?.workbook?.hiddenRow),
      cache: correlation.display,
      workbook: correlation.workbook,
      issues: correlation.issues,
      warnings: correlation.warnings,
      blocking: correlation.blocking,
    };
  }
  if (raw.kind === 'literal') {
    return {kind: 'literal', value: raw.value ?? null, comparison: 'literal', cell: null, cellValue: null, reference: null, formula: null, hiddenColumn: false, hiddenRow: false, cache: null, workbook: null, issues: [], warnings: [], blocking: []};
  }
  for (const entry of raw.issues ?? []) seriesIssues.push({...entry, seriesName: true});
  return {kind: raw.kind, value: raw.value ?? null, comparison: raw.comparison ?? 'absent', cell: null, cellValue: null, reference: null, formula: null, hiddenColumn: false, hiddenRow: false, cache: null, workbook: null, issues: raw.issues ?? [], warnings: [], blocking: (raw.issues ?? []).filter(value => value.severity === 'blocking')};
}

export function readSeries(ser, {groupKind, seriesIndex, pointer, workbook, workbookStatus = null, chartPart}) {
  const seriesIssues = [];
  const name = finishSeriesName(readSeriesName(ser, pointer), workbook, chartPart, seriesIssues, workbookStatus);
  const categoriesNode = direct(ser, 'cat') ?? direct(ser, 'xVal');
  const valuesNode = direct(ser, 'val') ?? direct(ser, 'yVal');
  if (!valuesNode) seriesIssues.push(issue('blocking', 'missingValues', '系列没有数值来源', {part: chartPart, pointer}));
  const categoryRef = readDataReference(categoriesNode, {role: 'category', pointer: `${pointer}/cat`});
  const valueRef = readDataReference(valuesNode, {role: 'value', pointer: `${pointer}/val`});
  const category = categoriesNode ? correlate(categoryRef, workbook, {chartPart, label: '类别', workbookStatus}) : null;
  const value = correlate(valueRef, workbook, {chartPart, label: '数值', workbookStatus});
  if (!category && categoriesNode) seriesIssues.push(issue('blocking', 'missingCategories', '系列没有可用的类别来源', {part: chartPart, pointer: `${pointer}/cat`}));
  if (!categoriesNode) seriesIssues.push(issue('warning', 'missingCategories', '系列没有类别来源', {part: chartPart, pointer}));
  const categoryByIndex = new Map((category?.points ?? []).map(point => [point.index, point]));
  const points = value.points.map(point => {
    const categoryPoint = categoryByIndex.get(point.index) ?? null;
    let categoryStatus;
    if (!category) categoryStatus = 'absent';
    else if (!categoryPoint) categoryStatus = 'missing';
    else categoryStatus = categoryPoint.status;
    return {
      index: point.index,
      withinReference: point.withinReference,
      category: categoryPoint ? {
        display: categoryPoint.display,
        workbook: categoryPoint.workbook,
        status: categoryPoint.status,
      } : {display: {present: false, value: null, source: null}, workbook: {present: false, value: null, cell: null, source: null}, status: category ? 'missing' : 'absent'},
      value: {display: point.display, workbook: point.workbook, status: point.status},
      status: point.status,
      categoryStatus,
    };
  });
  const reasons = [];
  const warnings = [];
  const blockingCodes = new Set();
  const pushReason = (value) => {
    const key = `${value.code}|${value.message}`;
    if (reasons.some(existing => `${existing.code}|${existing.message}` === key)) return;
    reasons.push(value);
    blockingCodes.add(value.code);
  };
  const pushWarning = (value) => {
    const key = `${value.code}|${value.message}`;
    if (warnings.some(existing => `${existing.code}|${existing.message}` === key)) return;
    warnings.push(value);
  };
  for (const entry of seriesIssues) {
    if (entry.severity === 'blocking') pushReason({code: entry.code, message: entry.message, source: entry.source ?? null});
    else pushWarning({code: entry.code, message: entry.message, source: entry.source ?? null});
  }
  for (const entry of [...value.blocking, ...(category?.blocking ?? []), ...(name.blocking ?? [])]) pushReason({code: entry.code, message: entry.message, source: entry.source ?? null});
  for (const entry of [...value.warnings, ...(category?.warnings ?? []), ...(name.warnings ?? [])]) pushWarning({code: entry.code, message: entry.message, source: entry.source ?? null});
  if (value.workbook.status === 'literal') pushReason({code: 'literalValueUncorroborated', message: '数值来自图表内的字面量，没有工作簿可供核对', count: value.points.length});
  if (value.counts.unmapped > 0) pushReason({code: 'valueUnmapped', message: `${value.counts.unmapped} 个数值位置无法对应到工作簿单元格`, count: value.counts.unmapped});
  if (value.counts.conflict > 0) pushReason({code: 'valueConflict', message: `${value.counts.conflict} 个数值显示缓存与工作簿单元格不一致`, count: value.counts.conflict});
  if (value.counts.displayOnly > 0) pushReason({code: 'valueMissingInWorkbook', message: `${value.counts.displayOnly} 个数值只在显示缓存中`, count: value.counts.displayOnly});
  if (value.counts.workbookOnly > 0) pushReason({code: 'valueMissingInDisplay', message: `${value.counts.workbookOnly} 个数值只在工作簿中`, count: value.counts.workbookOnly});
  if (value.counts.missing > 0) pushReason({code: 'valueGap', message: `${value.counts.missing} 个数值位置在显示缓存与工作簿中都没有内容`, count: value.counts.missing});
  if (!value.points.length) pushReason({code: 'noValuePoints', message: '数值来源没有可记录的数据点'});
  if (category) {
    if (category.counts.unmapped > 0) pushReason({code: 'categoryUnmapped', message: `${category.counts.unmapped} 个类别位置无法对应到工作簿单元格`, count: category.counts.unmapped});
    if (category.counts.conflict > 0) pushReason({code: 'categoryConflict', message: `${category.counts.conflict} 个类别显示缓存与工作簿单元格不一致`, count: category.counts.conflict});
    if (category.counts.displayOnly > 0) pushReason({code: 'categoryMissingInWorkbook', message: `${category.counts.displayOnly} 个类别只在显示缓存中`, count: category.counts.displayOnly});
    if (category.counts.workbookOnly > 0) pushReason({code: 'categoryMissingInDisplay', message: `${category.counts.workbookOnly} 个类别只在工作簿中`, count: category.counts.workbookOnly});
    const categoryGaps = points.filter(point => point.categoryStatus === 'missing' || point.categoryStatus === 'displayOnly' || point.categoryStatus === 'workbookOnly').length;
    if (categoryGaps > 0) pushReason({code: 'categoryGap', message: `${categoryGaps} 个位置缺少可核对的类别`, count: categoryGaps});
    if (category.expectedCount !== value.expectedCount) {
      pushWarning({code: 'categoryLengthMismatch', message: `类别有 ${category.expectedCount} 个位置，数值有 ${value.expectedCount} 个位置`, category: category.expectedCount, value: value.expectedCount});
    }
  }
  if (name.comparison === 'conflict') pushReason({code: 'nameConflict', message: '系列名称的显示缓存与工作簿单元格不一致', source: name});
  if (name.comparison === 'displayOnly' || name.comparison === 'workbookOnly' || name.comparison === 'missing') pushWarning({code: 'nameNotCorroborated', message: `系列名称无法与工作簿核对（${name.comparison}）`, source: name});
  const confirmed = reasons.length === 0 &&
    value.points.length > 0 &&
    value.points.every(point => point.status === 'match' || point.status === 'literal') &&
    (!category || points.every(point => point.categoryStatus === 'match' || point.categoryStatus === 'literal' || (!category.points.length && point.categoryStatus === 'absent')));
  const series = {
    index: seriesIndex,
    order: Number(attr(direct(ser, 'order'), 'val') ?? seriesIndex),
    chartKind: groupKind,
    name,
    category: publicCorrelation(category),
    value: publicCorrelation(value),
    points,
    confirmed,
    reasons,
    warnings,
    issues: [...value.issues, ...(category?.issues ?? []), ...(name.issues ?? [])],
  };
  return series;
}

function publicCorrelation(correlation) {
  if (!correlation) {
    return {
      role: 'category', sourceKind: 'none', formula: null, resolution: 'none', reference: null,
      display: null,
      workbook: {status: 'absent', part: null, sheet: null, sheetState: null, cells: [], hiddenColumns: [], hiddenRows: [], issues: []},
      expectedCount: 0,
      lengths: {declaredPtCount: null, cachedPoints: 0, referenceSize: null, workbookCells: 0},
      points: [], gaps: [], conflicts: [], issues: [], warnings: [], blocking: [],
      counts: {match: 0, conflict: 0, displayOnly: 0, workbookOnly: 0, missing: 0, literal: 0, unmapped: 0},
    };
  }
  return {
    role: correlation.role,
    sourceKind: correlation.sourceKind,
    formula: correlation.formula,
    resolution: correlation.resolution,
    reference: correlation.reference,
    display: correlation.display,
    workbook: correlation.workbook,
    expectedCount: correlation.expectedCount,
    lengths: correlation.lengths,
    points: correlation.points,
    gaps: correlation.gaps,
    conflicts: correlation.conflicts,
    issues: correlation.issues,
    warnings: correlation.warnings,
    blocking: correlation.blocking,
    counts: correlation.counts,
  };
}

const CHART_GROUPS = new Set(['barChart', 'bar3DChart', 'lineChart', 'line3DChart', 'pieChart', 'pie3DChart', 'doughnutChart', 'areaChart', 'area3DChart', 'scatterChart', 'bubbleChart', 'radarChart', 'surfaceChart', 'surface3DChart', 'stockChart', 'ofPieChart']);

/**
 * Reads one chart part. Isolation boundary: a corrupt chart or a bad series is
 * recorded on that entity only and never prevents other charts from reading.
 */
export function readChart(pkg, partName, {workbook = null, workbookInfo = null} = {}) {
  const workbookStatus = workbookInfo?.status ?? null;
  const chart = {
    part: partName,
    status: 'parsed',
    digest: null,
    title: null,
    chartTypes: [],
    series: [],
    externalData: null,
    externalLinks: [],
    workbook: workbookInfo,
    issues: [],
    warnings: [],
    reasons: [],
    confirmed: false,
    references: [],
  };
  let root;
  try { root = pkg.xmlPart(partName); }
  catch (error) {
    chart.status = 'corrupt';
    chart.issues.push(issue('blocking', 'chartUnreadable', error.message, {part: partName}));
    chart.reasons.push({code: 'chartUnreadable', message: error.message, source: {part: partName}});
    return chart;
  }
  if (localName(root.name) !== 'chartSpace') {
    chart.status = 'corrupt';
    chart.issues.push(issue('blocking', 'notAChart', `根元素 ${root.name} 不是 chartSpace`, {part: partName}));
    chart.reasons.push({code: 'notAChart', message: `根元素 ${root.name} 不是 chartSpace`, source: {part: partName}});
    return chart;
  }
  const chartNode = firstElement(root, 'chart');
  if (!chartNode) {
    chart.status = 'corrupt';
    chart.issues.push(issue('blocking', 'missingChartElement', 'chartSpace 内没有 chart 元素', {part: partName}));
    chart.reasons.push({code: 'missingChartElement', message: 'chartSpace 内没有 chart 元素', source: {part: partName}});
    return chart;
  }
  const titleNode = firstElement(chartNode, 'title');
  if (titleNode) {
    const texts = [];
    for (const node of elements(titleNode, 't')) texts.push(textOf(node));
    if (texts.length === 0) {
      for (const node of descendantsOf(titleNode, 't')) texts.push(textOf(node));
    }
    chart.title = texts.join('') || null;
  }
  const plotArea = firstElement(chartNode, 'plotArea');
  if (!plotArea) chart.issues.push(issue('blocking', 'missingPlotArea', '图表没有 plotArea', {part: partName}));
  let seriesIndex = 0;
  for (const group of plotArea ? plotArea.children : []) {
    const groupKind = localName(group.name);
    if (!CHART_GROUPS.has(groupKind)) continue;
    chart.chartTypes.push(groupKind);
    const groupPosition = chart.chartTypes.length;
    for (const ser of elements(group, 'ser')) {
      try {
        const series = readSeries(ser, {
          groupKind,
          seriesIndex,
          pointer: `/chartSpace/chart/plotArea/${groupKind}[${groupPosition}]/ser[${seriesIndex}]`,
          workbook,
          workbookStatus,
          chartPart: partName,
        });
        chart.series.push(series);
      } catch (error) {
        chart.series.push({
          index: seriesIndex, order: seriesIndex, chartKind: groupKind,
          name: null, category: null, value: null, points: [],
          confirmed: false,
          reasons: [{code: 'seriesUnreadable', message: error.message}],
          warnings: [], issues: [issue('blocking', 'seriesUnreadable', error.message, {part: partName})],
        });
      }
      seriesIndex++;
    }
  }
  if (chart.series.length === 0) chart.issues.push(issue('blocking', 'noSeries', '图表没有可读的系列', {part: partName}));

  const externalData = firstElement(root, 'externalData') ?? firstElement(chartNode, 'externalData');
  if (externalData) {
    const dataId = attr(externalData, 'r:id') ?? attr(externalData, 'id');
    let relationship = null;
    try { relationship = dataId ? pkg.relationship(partName, dataId) : null; }
    catch (error) { chart.issues.push(issue('blocking', 'chartRelationshipsUnreadable', error.message, {part: partName})); }
    chart.externalData = {
      relationshipId: dataId ?? null,
      relationshipFound: Boolean(relationship),
      external: Boolean(relationship?.external),
      target: relationship?.resolved ?? relationship?.target ?? null,
      rawTarget: relationship?.target ?? null,
      kind: relationship ? relationship.type.split('/').pop() : null,
      followed: false,
    };
    if (!relationship && !chart.issues.some(value => value.code === 'chartRelationshipsUnreadable')) {
      chart.issues.push(issue('blocking', 'missingWorkbookRelationship', `externalData 引用的关系 ${dataId} 不存在`, {part: partName}));
    } else if (relationship?.external) {
      chart.externalLinks.push({declaredIn: partName, relationshipId: relationship.id, kind: 'externalData', target: relationship.target, targetMode: relationship.targetMode, followed: false});
      chart.issues.push(issue('blocking', 'externalWorkbook', `图表的外部数据源未访问: ${relationship.target}`, {part: partName}));
    } else if (relationship?.resolutionError) {
      chart.issues.push(issue('blocking', 'unresolvedWorkbook', relationship.resolutionError, {part: partName}));
    }
  }
  let relationships = [];
  try { relationships = pkg.relationships(partName); }
  catch (error) { chart.issues.push(issue('blocking', 'chartRelationshipsUnreadable', error.message, {part: partName})); }
  for (const relationship of relationships) {
    if (!relationship.external) continue;
    if (chart.externalLinks.some(link => link.relationshipId === relationship.id)) continue;
    chart.externalLinks.push({declaredIn: partName, relationshipId: relationship.id, kind: relationship.type.split('/').pop() || 'external', target: relationship.target, targetMode: relationship.targetMode, followed: false});
  }
  for (const entry of chart.issues) {
    if (entry.severity === 'blocking') chart.reasons.push({code: entry.code, message: entry.message, source: entry.source ?? null});
    else chart.warnings.push({code: entry.code, message: entry.message, source: entry.source ?? null});
  }
  for (const series of chart.series) {
    for (const entry of series.reasons) chart.reasons.push({...entry, series: series.index, message: `系列 ${series.index + 1}: ${entry.message}`});
    for (const entry of series.warnings) chart.warnings.push({...entry, series: series.index, message: `系列 ${series.index + 1}: ${entry.message}`});
  }
  chart.confirmed = chart.status === 'parsed' && chart.series.length > 0 && chart.reasons.length === 0 && chart.series.every(series => series.confirmed);
  const uniqueReasons = new Map();
  for (const entry of chart.reasons) uniqueReasons.set(`${entry.code}|${entry.message}`, entry);
  chart.reasons = [...uniqueReasons.values()];
  const uniqueWarnings = new Map();
  for (const entry of chart.warnings) uniqueWarnings.set(`${entry.code}|${entry.message}`, entry);
  chart.warnings = [...uniqueWarnings.values()];
  return chart;
}

function descendantsOf(node, name) {
  const wanted = localName(name);
  const output = [];
  (function walk(current) {
    for (const child of current.children) {
      if (localName(child.name) === wanted) output.push(child);
      walk(child);
    }
  })(node);
  return output;
}
