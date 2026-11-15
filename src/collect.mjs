import {createHash} from 'node:crypto';
import {OpcPackage, relationshipTypeSuffix} from './opc.mjs';
import {attr, descendants, elements, firstElement} from './xml.mjs';
import {Workbook} from './spreadsheet.mjs';
import {readChart} from './chart.mjs';

const CHART_RELATIONSHIP_SUFFIX = '/chart';
const CHART_URI = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

function sha256(bytes) {
  return createHash('sha256').update(bytes).digest('hex');
}

export function parsePageSelection(value) {
  if (value === undefined || value === null) return null;
  if (typeof value === 'string' && value.trim() === '') throw new Error('页码选择为空');
  const parts = Array.isArray(value) ? value : String(value).split(',');
  const pages = [];
  for (const part of parts) {
    const text = String(part).trim();
    if (!text) continue;
    const match = /^(\d+)(?:-(\d+))?$/.exec(text);
    if (!match) throw new Error(`页码格式无效: ${text}（应为 1 或 1-3）`);
    const start = Number(match[1]);
    const end = match[2] ? Number(match[2]) : start;
    if (start < 1 || end < start) throw new Error(`页码范围无效: ${text}`);
    for (let page = start; page <= end; page++) if (!pages.includes(page)) pages.push(page);
  }
  if (pages.length === 0) throw new Error('页码选择为空');
  return pages.sort((left, right) => left - right);
}

function slideEntityId(slidePart, digest) {
  return `slide:${slidePart}#${digest.slice(0, 16)}`;
}

function chartEntityId(chartPart, digest) {
  return `chart:${chartPart}#${digest.slice(0, 16)}`;
}

function workbookEntityId(workbookPart, digest) {
  return `workbook:${workbookPart}#${digest.slice(0, 16)}`;
}

export function listSlides(pkg) {
  const presentationPart = pkg.mainDocumentPart();
  const slides = [];
  const issues = [];
  if (!presentationPart) {
    issues.push({severity: 'blocking', code: 'missingPresentation', message: '包内没有 presentation.xml 主文档关系'});
    return {presentationPart: null, slides, issues};
  }
  let root;
  try { root = pkg.xmlPart(presentationPart); }
  catch (error) {
    issues.push({severity: 'blocking', code: 'presentationUnreadable', message: error.message, source: {part: presentationPart}});
    return {presentationPart, slides, issues};
  }
  let relationships = [];
  try { relationships = pkg.relationships(presentationPart); }
  catch (error) {
    issues.push({severity: 'blocking', code: 'presentationRelationshipsUnreadable', message: error.message, source: {part: presentationPart}});
    return {presentationPart, slides, issues};
  }
  const list = firstElement(root, 'sldIdLst');
  const entries = list ? elements(list, 'sldId') : [];
  if (entries.length === 0) issues.push({severity: 'blocking', code: 'noSlides', message: '演示文稿没有幻灯片列表', source: {part: presentationPart}});
  for (const [position, entry] of entries.entries()) {
    const relationshipId = attr(entry, 'r:id') ?? attr(entry, 'id');
    const slideId = attr(entry, 'id');
    const relationship = relationships.find(value => value.id === relationshipId) ?? null;
    if (!relationship) {
      issues.push({severity: 'blocking', code: 'missingSlideRelationship', message: `幻灯片 ${position + 1} 的关系 ${relationshipId} 不存在`, source: {part: presentationPart}});
      slides.push({page: position + 1, slideId, relationshipId, part: null, problem: 'missingSlideRelationship'});
      continue;
    }
    if (relationship.external) {
      issues.push({severity: 'blocking', code: 'externalSlide', message: `幻灯片 ${position + 1} 指向外部资料，未访问: ${relationship.target}`, source: {part: presentationPart}});
      slides.push({page: position + 1, slideId, relationshipId, part: null, problem: 'externalSlide', externalTarget: relationship.target});
      continue;
    }
    if (relationship.resolutionError) {
      issues.push({severity: 'blocking', code: 'unresolvedSlide', message: `幻灯片 ${position + 1} 的路径无法解析: ${relationship.resolutionError}`, source: {part: presentationPart}});
      slides.push({page: position + 1, slideId, relationshipId, part: null, problem: 'unresolvedSlide'});
      continue;
    }
    slides.push({page: position + 1, slideId, relationshipId, part: relationship.resolved, problem: null, relationshipType: relationship.type});
  }
  return {presentationPart, slides, issues};
}

function slideChartFrames(pkg, slidePart) {
  const root = pkg.xmlPart(slidePart);
  const frames = [];
  const graphicFrames = descendants(root, 'graphicFrame');
  for (const [frameIndex, frame] of graphicFrames.entries()) {
    const graphic = firstElement(frame, 'graphic');
    const graphicData = graphic ? firstElement(graphic, 'graphicData') : null;
    const uri = graphicData ? attr(graphicData, 'uri') : null;
    if (uri !== null && uri !== CHART_URI) continue;
    const chartNodes = graphicData ? descendants(graphicData, 'chart') : [];
    for (const [chartIndex, chartNode] of chartNodes.entries()) {
      const properties = firstElement(frame, 'nvGraphicFramePr');
      const nonVisual = properties ? firstElement(properties, 'cNvPr') : null;
      frames.push({
        frameIndex,
        chartIndex,
        uri,
        relationshipId: attr(chartNode, 'r:id') ?? attr(chartNode, 'id'),
        name: nonVisual ? attr(nonVisual, 'name') : null,
        id: nonVisual ? attr(nonVisual, 'id') : null,
      });
    }
  }
  return frames;
}

const ACTIVE_CONTENT_PATTERNS = [
  {kind: 'macro', pattern: /(^|\/)vbaProject\.bin$/i, contentType: 'application/vnd.ms-office.vbaProject'},
  {kind: 'oleObject', pattern: /(^|\/)embeddings\/oleObject[^/]*\.bin$/i},
  {kind: 'embeddedPackage', pattern: /(^|\/)embeddings\/[^/]+\.(exe|dll|js|vbs|ps1|bat|cmd|jar)$/i},
];

/**
 * Lists parts that could execute code or run programs if opened by an office
 * application. This collector never opens them; it only names them.
 */
export function scanActiveContent(pkg) {
  const found = [];
  let contentTypes = null;
  try { contentTypes = pkg.contentTypes(); } catch { contentTypes = null; }
  for (const name of pkg.archive.names()) {
    for (const rule of ACTIVE_CONTENT_PATTERNS) {
      if (!rule.pattern.test(name)) continue;
      let contentType = rule.contentType ?? null;
      if (!contentType && contentTypes) {
        try { contentType = contentTypes.contentTypeOf(name); } catch { contentType = null; }
      }
      found.push({
        part: name,
        kind: rule.kind,
        contentType,
        bytes: pkg.archive.entry(name)?.uncompressedSize ?? null,
        executed: false,
        read: false,
      });
      break;
    }
  }
  return found;
}

function slideHidden(pkg, slidePart) {
  try {
    const root = pkg.xmlPart(slidePart);
    const show = attr(root, 'show');
    return show === '0' || show === 'false';
  } catch {
    return false;
  }
}

/**
 * Opens the workbook behind a chart. A workbook is a shared entity keyed by
 * part name and content digest, so repeated references share one record.
 */
function workbookForChart(pkg, chartPart, cache, result) {
  const state = {workbook: null, entityId: null, info: {status: 'absent', part: null}, issues: []};
  let relationships;
  try { relationships = pkg.relationships(chartPart); }
  catch (error) {
    state.info = {status: 'error', part: null, message: error.message};
    state.issues.push({severity: 'blocking', code: 'chartRelationshipsUnreadable', message: error.message, source: {part: chartPart}});
    return state;
  }
  let relationship = null;
  let externalData = null;
  try {
    const root = pkg.xmlPart(chartPart);
    const chartNode = firstElement(root, 'chart') ?? root;
    externalData = firstElement(root, 'externalData') ?? firstElement(chartNode, 'externalData');
  } catch {
    externalData = null;
  }
  if (externalData) {
    const id = attr(externalData, 'r:id') ?? attr(externalData, 'id');
    relationship = relationships.find(value => value.id === id) ?? null;
    if (!relationship) state.issues.push({severity: 'blocking', code: 'missingWorkbookRelationship', message: `externalData 的关系 ${id} 不存在`, source: {part: chartPart}});
  }
  if (!relationship) {
    const candidates = relationships.filter(value => value.type.endsWith('/package') || value.type.endsWith('/oleObject'));
    if (candidates.length === 1) relationship = candidates[0];
    else if (candidates.length > 1) {
      state.issues.push({severity: 'blocking', code: 'ambiguousWorkbookRelationship', message: '图表有多个可能的工作簿关系，无法确定引用', source: {part: chartPart}});
      state.info = {status: 'ambiguous', part: null};
      return state;
    }
  }
  if (!relationship) {
    state.info = {status: 'absent', part: null};
    return state;
  }
  if (relationship.external) {
    state.info = {status: 'external', part: null, target: relationship.target, relationshipId: relationship.id};
    state.issues.push({severity: 'blocking', code: 'externalWorkbook', message: `嵌入工作簿关系指向外部资料，未访问: ${relationship.target}`, source: {part: chartPart}});
    result.externalLinks.push({declaredIn: chartPart, relationshipId: relationship.id, kind: relationshipTypeSuffix(relationship.type) || 'external', target: relationship.target, targetMode: relationship.targetMode, followed: false});
    return state;
  }
  if (relationship.resolutionError) {
    state.info = {status: 'unresolved', part: null, message: relationship.resolutionError};
    state.issues.push({severity: 'blocking', code: 'unresolvedWorkbook', message: relationship.resolutionError, source: {part: chartPart}});
    return state;
  }
  const part = relationship.resolved;
  let bytes;
  try { bytes = pkg.archive.read(part); }
  catch (error) {
    const status = pkg.hasPart(part) ? 'corrupt' : 'missing';
    state.info = {status, part, message: error.message};
    state.issues.push({severity: 'blocking', code: status === 'missing' ? 'missingWorkbookPart' : 'corruptWorkbookPart', message: error.message, source: {part}});
    return state;
  }
  const digest = sha256(bytes);
  const id = workbookEntityId(part, digest);
  const cached = cache.get(id);
  state.entityId = id;
  if (cached) {
    state.workbook = cached.workbook;
    state.info = {status: cached.entity.status, part, entity: id};
    return state;
  }
  const entity = {
    id, part, digest: `sha256:${digest}`, bytes: bytes.length,
    status: 'read', message: null, sheets: [], hiddenColumns: null,
  };
  let workbook = null;
  try {
    const inner = new OpcPackage(bytes, {label: part});
    const workbookPart = inner.mainDocumentPart() ?? (inner.hasPart('xl/workbook.xml') ? 'xl/workbook.xml' : null);
    if (!workbookPart) throw new Error('嵌入资料没有 xl/workbook.xml');
    workbook = Workbook.open(inner, workbookPart);
    workbook.part = part;
    // Force workbook and sheet metadata now so a broken workbook is recorded once.
    const meta = workbook.sheets;
    entity.sheets = meta.sheets.map(sheet => ({name: sheet.name, state: sheet.state, part: sheet.part, problem: sheet.problem}));
    entity.date1904 = meta.date1904;
  } catch (error) {
    entity.status = 'unreadable';
    entity.message = `嵌入工作簿无法作为 XLSX 读取: ${error.message}`;
    workbook = null;
  }
  state.workbook = workbook;
  state.info = {status: workbook ? 'read' : 'unreadable', part, entity: id, message: entity.message ?? null};
  cache.set(id, {workbook, entity});
  result.entities.workbooks[id] = entity;
  return state;
}

function stampPackage(correlation, entityId) {
  if (!correlation) return;
  if (correlation.workbook) correlation.workbook.package = entityId;
  for (const point of correlation.points ?? []) {
    if (point.workbook?.source) point.workbook.source.package = entityId;
  }
}

/**
 * Collects chart evidence from a presentation package, page by page. Page
 * failures are isolated; charts and workbooks are shared entities identified by
 * content digest. `pages` optionally limits collection to specific one-based
 * slide numbers, which is how supplements are added later.
 */
export function collectPresentation(bytes, {packageName = 'presentation', pages = null, label = null} = {}) {
  const buffer = Buffer.isBuffer(bytes) ? bytes : Buffer.from(bytes);
  const packageDigest = sha256(buffer);
  const pkg = new OpcPackage(buffer, {label: packageName});
  const listing = listSlides(pkg);
  const result = {
    schema: 'presentation-charts/1',
    label: label ?? packageName,
    collectedAt: new Date().toISOString(),
    package: {name: packageName, digest: `sha256:${packageDigest}`, bytes: buffer.length},
    selection: {requested: pages ?? 'all', available: listing.slides.length, resolved: [], missing: []},
    pages: [],
    entities: {slides: {}, charts: {}, workbooks: {}},
    references: [],
    externalLinks: [],
    activeContent: scanActiveContent(pkg),
    issues: [...listing.issues],
    summary: null,
  };
  const available = new Set(listing.slides.map(slide => slide.page));
  if (pages) for (const page of pages) if (!available.has(page)) result.selection.missing.push(page);
  const selected = listing.slides.filter(slide => !pages || pages.includes(slide.page));
  if (selected.length === 0 && !pages) result.issues.push({severity: 'blocking', code: 'emptySelection', message: '没有可采集的页面'});
  const workbookCache = new Map();

  for (const slide of selected) {
    const pageRecord = {
      page: slide.page,
      slideId: slide.slideId,
      relationshipId: slide.relationshipId,
      slidePart: slide.part,
      status: 'collected',
      hidden: false,
      slide: null,
      origin: {package: {name: packageName, digest: `sha256:${packageDigest}`}, collectedAt: result.collectedAt},
      references: [],
      issues: [],
    };
    result.selection.resolved.push({
      page: slide.page,
      slidePart: slide.part,
      status: slide.problem ? 'failed' : 'selected',
      reasons: slide.problem ? [slide.problem] : [],
    });
    if (slide.problem) {
      pageRecord.status = slide.problem;
      pageRecord.issues.push({severity: 'blocking', code: slide.problem, message: `第 ${slide.page} 页无法读取`, source: {part: null}});
      result.pages.push(pageRecord);
      continue;
    }
    let slideBytes;
    try { slideBytes = pkg.archive.read(slide.part); }
    catch (error) {
      pageRecord.status = pkg.hasPart(slide.part) ? 'corrupt' : 'missing';
      pageRecord.issues.push({severity: 'blocking', code: pageRecord.status === 'missing' ? 'missingSlidePart' : 'corruptSlidePart', message: error.message, source: {part: slide.part}});
      result.pages.push(pageRecord);
      continue;
    }
    const slideDigest = sha256(slideBytes);
    const slideId = slideEntityId(slide.part, slideDigest);
    pageRecord.hidden = slideHidden(pkg, slide.part);
    pageRecord.slide = {id: slideId, part: slide.part, digest: `sha256:${slideDigest}`, bytes: slideBytes.length};
    if (!result.entities.slides[slideId]) {
      result.entities.slides[slideId] = {id: slideId, part: slide.part, digest: `sha256:${slideDigest}`, bytes: slideBytes.length, pages: [slide.page], hidden: pageRecord.hidden};
    } else if (!result.entities.slides[slideId].pages.includes(slide.page)) {
      result.entities.slides[slideId].pages.push(slide.page);
    }

    let frames;
    try { frames = slideChartFrames(pkg, slide.part); }
    catch (error) {
      pageRecord.status = 'corrupt';
      pageRecord.issues.push({severity: 'blocking', code: 'slideUnreadable', message: error.message, source: {part: slide.part}});
      result.pages.push(pageRecord);
      continue;
    }
    if (frames.length === 0) pageRecord.issues.push({severity: 'warning', code: 'noChartFrames', message: `第 ${slide.page} 页没有图表图形`, source: {part: slide.part}});
    try {
      for (const relationship of pkg.relationships(slide.part)) {
        if (!relationship.external) continue;
        if (result.externalLinks.some(link => link.declaredIn === slide.part && link.relationshipId === relationship.id)) continue;
        result.externalLinks.push({declaredIn: slide.part, relationshipId: relationship.id, kind: relationshipTypeSuffix(relationship.type) || 'external', target: relationship.target, targetMode: relationship.targetMode, followed: false, page: slide.page});
      }
    } catch (error) {
      pageRecord.issues.push({severity: 'blocking', code: 'slideRelationshipsUnreadable', message: error.message, source: {part: slide.part}});
    }

    for (const frame of frames) {
      const reference = {
        id: `page:${slide.page}/frame:${frame.frameIndex + 1}/chart:${frame.chartIndex + 1}`,
        page: slide.page,
        slide: slideId,
        frame: {name: frame.name, id: frame.id, uri: frame.uri, index: frame.frameIndex + 1},
        relationshipId: frame.relationshipId,
        chartPart: null,
        chart: null,
        status: 'unresolved',
        reasons: [],
      };
      const attach = () => {
        pageRecord.references.push(reference);
        result.references.push(reference);
      };
      if (!frame.relationshipId) {
        reference.reasons.push({code: 'chartWithoutRelationship', message: '图表图形没有指向图表的 r:id', source: {part: slide.part}});
        attach();
        continue;
      }
      let relationship;
      try { relationship = pkg.relationship(slide.part, frame.relationshipId); }
      catch (error) {
        reference.reasons.push({code: 'slideRelationshipsUnreadable', message: error.message, source: {part: slide.part}});
        attach();
        continue;
      }
      if (!relationship) {
        reference.reasons.push({code: 'missingChartRelationship', message: `图表关系 ${frame.relationshipId} 不存在`, source: {part: slide.part}});
        attach();
        continue;
      }
      if (relationship.external) {
        reference.status = 'external';
        reference.reasons.push({code: 'externalChart', message: `图表位于外部资料，未访问: ${relationship.target}`, source: {part: slide.part}});
        result.externalLinks.push({declaredIn: slide.part, relationshipId: relationship.id, kind: relationshipTypeSuffix(relationship.type) || 'external', target: relationship.target, targetMode: relationship.targetMode, followed: false, page: slide.page});
        attach();
        continue;
      }
      if (relationship.resolutionError) {
        reference.reasons.push({code: 'unresolvedChart', message: relationship.resolutionError, source: {part: slide.part}});
        attach();
        continue;
      }
      if (!relationship.type.endsWith(CHART_RELATIONSHIP_SUFFIX)) {
        const contentType = pkg.hasPart(relationship.resolved) ? pkg.contentTypes().contentTypeOf(relationship.resolved) : null;
        if (contentType !== 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml') {
          reference.reasons.push({code: 'unexpectedChartRelationship', message: `关系类型不是图表: ${relationship.type}`, source: {part: slide.part}});
          attach();
          continue;
        }
      }
      const chartPart = relationship.resolved;
      reference.chartPart = chartPart;
      let chartBytes;
      try { chartBytes = pkg.archive.read(chartPart); }
      catch (error) {
        reference.status = pkg.hasPart(chartPart) ? 'corrupt' : 'missing';
        const code = reference.status === 'missing' ? 'missingChartPart' : 'corruptChartPart';
        reference.reasons.push({severity: 'blocking', code, message: error.message, source: {part: chartPart}});
        const brokenId = `chart:${chartPart}#${reference.status}`;
        reference.entity = brokenId;
        reference.chart = brokenId;
        if (!result.entities.charts[brokenId]) {
          result.entities.charts[brokenId] = {
            id: brokenId, part: chartPart, digest: null, bytes: null,
            status: reference.status, title: null, chartTypes: [], series: [],
            externalData: null, externalLinks: [], workbook: null,
            issues: [{severity: 'blocking', code, message: error.message, source: {part: chartPart}}],
            warnings: [], reasons: [{code, message: error.message, source: {part: chartPart}}],
            confirmed: false, references: [],
          };
        }
        attach();
        continue;
      }
      const chartDigest = sha256(chartBytes);
      const entityId = chartEntityId(chartPart, chartDigest);
      reference.entity = entityId;
      reference.chart = entityId;
      const existing = result.entities.charts[entityId];
      if (existing) {
        reference.shared = true;
        reference.status = existing.status === 'corrupt' ? 'corrupt' : existing.confirmed ? 'confirmed' : 'unconfirmed';
        reference.chartStatus = existing.status;
        reference.chartConfirmed = Boolean(existing.confirmed);
        reference.reasons = (existing.reasons ?? []).map(entry => ({...entry}));
        attach();
        continue;
      }
      const workbookState = workbookForChart(pkg, chartPart, workbookCache, result);
      let chart;
      try {
        chart = readChart(pkg, chartPart, {workbook: workbookState.workbook, workbookInfo: workbookState.info});
      } catch (error) {
        chart = {
          part: chartPart, status: 'corrupt', title: null, chartTypes: [], series: [],
          externalData: null, externalLinks: [], workbook: workbookState.info,
          issues: [{severity: 'blocking', code: 'chartUnreadable', message: error.message, source: {part: chartPart}}],
          warnings: [], reasons: [{code: 'chartUnreadable', message: error.message}],
          confirmed: false,
        };
      }
      chart.id = entityId;
      chart.digest = `sha256:${chartDigest}`;
      chart.bytes = chartBytes.length;
      chart.workbook = {...workbookState.info, entity: workbookState.entityId};
      for (const value of workbookState.issues) {
        if (value.severity === 'blocking') chart.reasons.push({code: value.code, message: value.message, source: value.source ?? null});
        else chart.warnings.push({code: value.code, message: value.message, source: value.source ?? null});
      }
      for (const link of chart.externalLinks) result.externalLinks.push({...link, page: slide.page});
      reference.reasons = chart.reasons.map(entry => ({...entry}));
      reference.chartStatus = chart.status;
      reference.chartConfirmed = Boolean(chart.confirmed);
      const workbookEntityIdValue = workbookState.entityId;
      if (workbookEntityIdValue) {
        for (const series of chart.series) {
          stampPackage(series.value, workbookEntityIdValue);
          stampPackage(series.category, workbookEntityIdValue);
          if (series.name?.workbook) stampPackage({workbook: series.name.workbook, points: []}, workbookEntityIdValue);
          for (const point of series.points) {
            for (const side of ['category', 'value']) {
              if (point[side]?.workbook?.source) point[side].workbook.source.package = workbookEntityIdValue;
            }
          }
        }
      }
      if (chart.reasons.length > 0) chart.confirmed = false;
      const uniqueReasons = new Map();
      for (const value of chart.reasons) uniqueReasons.set(`${value.code}|${value.message}`, value);
      chart.reasons = [...uniqueReasons.values()];
      const uniqueWarnings = new Map();
      for (const value of chart.warnings) uniqueWarnings.set(`${value.code}|${value.message}`, value);
      chart.warnings = [...uniqueWarnings.values()];
      chart.confirmed = Boolean(chart.confirmed) && chart.reasons.length === 0;
      result.entities.charts[entityId] = chart;
      reference.status = chart.status === 'corrupt' ? 'corrupt' : chart.confirmed ? 'confirmed' : 'unconfirmed';
      attach();
    }
    if (pageRecord.references.length === 0) pageRecord.status = 'noCharts';
    else if (pageRecord.references.some(reference => ['unresolved', 'external', 'corrupt', 'missing'].includes(reference.status))) pageRecord.status = 'partial';
    else if (pageRecord.references.every(reference => reference.chartConfirmed)) pageRecord.status = 'confirmed';
    else pageRecord.status = 'unconfirmed';
    result.pages.push(pageRecord);
  }

  result.summary = summarize(result);
  return result;
}

function emptyCounts() {
  return {match: 0, conflict: 0, displayOnly: 0, workbookOnly: 0, missing: 0, literal: 0, unmapped: 0};
}

function countPoints(points, side) {
  const counts = emptyCounts();
  for (const point of points) {
    const status = point[side]?.status;
    if (status && Object.hasOwn(counts, status)) counts[status]++;
  }
  return counts;
}

export function summarize(result) {
  const charts = Object.values(result.entities.charts);
  const series = charts.flatMap(chart => chart.series ?? []);
  const points = series.flatMap(value => value.points ?? []);
  const counts = {value: countPoints(points, 'value'), category: countPoints(points, 'category')};
  const unconfirmed = charts.filter(chart => !chart.confirmed).map(chart => ({
    chart: chart.id,
    part: chart.part,
    status: chart.status,
    reasons: chart.reasons,
    warnings: chart.warnings,
  }));
  const blockingReasons = new Set([
    'chartUnreadable', 'notAChart', 'missingChartElement', 'missingPlotArea', 'noSeries',
    'externalWorkbook', 'unresolvedWorkbook', 'missingWorkbookRelationship', 'ambiguousWorkbookRelationship',
    'missingWorkbookPart', 'corruptWorkbookPart', 'chartRelationshipsUnreadable',
    'externalReference', 'unresolvedReference', 'workbookUnavailable', 'worksheetMissing',
    'worksheetUnreadable', 'ambiguousRange', 'formulaWithoutCachedValue', 'pointOutsideReference',
    'valueConflict', 'valueMissingInWorkbook', 'valueMissingInDisplay', 'valueGap', 'noValuePoints',
    'valueUnmapped', 'categoryUnmapped', 'missingChartPart', 'corruptChartPart',
    'categoryConflict', 'categoryMissingInWorkbook', 'categoryMissingInDisplay', 'categoryGap',
    'nameConflict', 'missingValues', 'seriesUnreadable', 'unknownDataSource', 'unknownSeriesName',
    'missingCache', 'duplicatePointIndex', 'missingCategories',
  ]);
  const allReasons = charts.flatMap(chart => [
    ...(chart.reasons ?? []).map(value => ({chart: chart.id, ...value})),
    ...(chart.warnings ?? []).map(value => ({chart: chart.id, warning: true, ...value})),
  ]);
  return {
    pages: result.pages.length,
    selected: result.selection.resolved.filter(value => value.status === 'selected').length,
    referenceCount: result.references.length,
    failedPages: result.pages.filter(page => ['partial', 'missing', 'corrupt', 'missingSlidePart', 'corruptSlidePart', 'externalSlide', 'unresolvedSlide', 'missingSlideRelationship'].includes(page.status)).map(page => ({page: page.page, status: page.status, issues: page.issues})),
    stalePages: result.pages.filter(page => page.origin && page.origin.package.digest !== result.package.digest).map(page => ({page: page.page, collectedFrom: page.origin.package, collectedAt: page.origin.collectedAt})),
    missingPages: result.selection.missing,
    charts: charts.length,
    confirmedCharts: charts.filter(chart => chart.confirmed).length,
    unconfirmedCharts: unconfirmed.length,
    unconfirmed,
    series: series.length,
    points: points.length,
    counts,
    externalLinks: result.externalLinks.length,
    externalTargets: [...new Set(result.externalLinks.map(link => link.target))],
    activeContentParts: result.activeContent.length,
    blockingReasons: allReasons.filter(value => blockingReasons.has(value.code) || (value.severity === 'blocking')),
    warnings: allReasons.filter(value => value.warning || value.severity === 'warning'),
  };
}

/**
 * Merges a later collection into an earlier one. Pages selected again replace
 * their prior page record; other pages stay intact. A different package digest
 * starts a new baseline and is recorded as superseding the old one.
 */
export function chartDetail(collection, chartId) {
  const chart = collection.entities?.charts?.[chartId] ?? null;
  if (!chart) return null;
  return {
    ...chart,
    references: (collection.references ?? []).filter(reference => reference.chart === chartId || reference.entity === chartId),
  };
}

export function referencesFor(collection, {page = null, chart = null} = {}) {
  return (collection.references ?? []).filter(reference =>
    (page === null || reference.page === page) &&
    (chart === null || reference.chart === chart || reference.entity === chart));
}

export function mergeCollection(previous, next) {
  if (!previous) return next;
  const pages = new Map((previous.pages ?? []).map(page => [page.page, page]));
  const selectedPages = new Set(next.pages.map(page => page.page));
  for (const page of next.pages) pages.set(page.page, page);
  const entities = {
    slides: {...previous.entities?.slides, ...next.entities.slides},
    charts: {...previous.entities?.charts, ...next.entities.charts},
    workbooks: {...previous.entities?.workbooks, ...next.entities.workbooks},
  };
  for (const [id, slide] of Object.entries(entities.slides)) {
    const before = previous.entities?.slides?.[id];
    if (!before) continue;
    slide.pages = [...new Set([...(before.pages ?? []), ...(slide.pages ?? [])])].sort((left, right) => left - right);
  }
  const activeContent = new Map();
  for (const part of [...(previous.activeContent ?? []), ...next.activeContent]) activeContent.set(part.part, part);
  const references = [
    ...(previous.references ?? []).filter(reference => !selectedPages.has(reference.page)),
    ...next.references,
  ];
  const externalLinks = new Map();
  for (const link of [...(previous.externalLinks ?? []), ...next.externalLinks]) {
    externalLinks.set(`${link.declaredIn}|${link.relationshipId}|${link.target}`, link);
  }
  const resolved = new Map((previous.selection?.resolved ?? []).map(value => [value.page, value]));
  for (const value of next.selection.resolved) resolved.set(value.page, value);
  const merged = {
    ...next,
    selection: {
      requested: next.selection.requested,
      available: next.selection.available,
      resolved: [...resolved.values()].sort((left, right) => left.page - right.page),
      missing: next.selection.missing,
    },
    pages: [...pages.values()].sort((left, right) => left.page - right.page),
    entities,
    references,
    externalLinks: [...externalLinks.values()],
    activeContent: [...activeContent.values()],
    history: [
      ...(previous.history ?? []),
      ...(previous.collectedAt ? [{collectedAt: previous.collectedAt, package: previous.package, pages: previous.pages.map(page => page.page)}] : []),
    ],
  };
  if (previous.package?.digest !== next.package?.digest) {
    merged.packageChanges = [
      ...(previous.packageChanges ?? []),
      {
        from: previous.package,
        to: next.package,
        at: next.collectedAt,
        retainedPages: [...pages.values()].filter(page => !selectedPages.has(page.page)).map(page => page.page),
      },
    ];
  } else merged.packageChanges = previous.packageChanges;
  merged.summary = summarize(merged);
  return merged;
}
