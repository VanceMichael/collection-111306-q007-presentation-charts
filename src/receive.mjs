import {resolve, basename} from 'node:path';
import {readFile} from 'node:fs/promises';
import {parseArgs} from 'node:util';
import {startSource} from './source.mjs';
import {CaptureStore} from './store.mjs';
import {collectPresentation, mergeCollection, referencesFor, chartDetail, parsePageSelection} from './collect.mjs';

export const RECORD_KIND = 'presentation-charts';

/**
 * Receives a presentation (bytes) into the capture store. Only the selected
 * pages are collected; earlier pages remain in the merged record unless they
 * are selected again. Records store the merged collection plus its per-run
 * summary, so a later reader never has to replay logs to know what happened.
 */
export async function receivePresentation({store, bytes, packageName, pages = null, at = new Date().toISOString()}) {
  const collection = collectPresentation(bytes, {packageName, pages, label: packageName});
  if (pages && collection.pages.length === 0) {
    throw new Error(`请求的页码均不存在: ${collection.selection.missing.join(',')}（文稿共 ${collection.selection.available ?? '?'} 页）`);
  }
  const state = await store.read();
  const previousRecords = state.records.filter(record => record.kind === RECORD_KIND && record.collection?.package?.name === packageName);
  const previousRecord = previousRecords.at(-1) ?? null;
  const previous = previousRecord?.collection ?? null;
  const merged = previous ? mergeCollection(previous, collection) : collection;
  const record = {
    kind: RECORD_KIND,
    receivedAt: at,
    package: merged.package,
    selection: merged.selection,
    summary: merged.summary,
    collection: merged,
    previousRevision: previousRecord ? state.revision : null,
  };
  const receipt = await store.append(record);
  return {record: {...record, revision: receipt.revision}, collection: merged};
}

export function describeCollection(collection, {page = null, chartId = null} = {}) {
  const lines = [];
  const summary = collection.summary;
  lines.push(`文稿: ${collection.package.name} (${collection.package.digest.slice(0, 19)}…) ${collection.package.bytes} 字节`);
  const pageNumbers = collection.pages.map(value => value.page);
  lines.push(`页面: 已保存 ${pageNumbers.length ? pageNumbers.join(',') : '无'} 页；本次请求 ${collection.selection.requested === 'all' ? '全部' : collection.selection.requested.join(',')}；未找到页 ${summary.missingPages.length ? summary.missingPages.join(',') : '无'}`);
  if (summary.stalePages?.length) {
    lines.push(`注意: 第 ${summary.stalePages.map(value => value.page).join(',')} 页来自旧文稿 ${summary.stalePages.map(value => value.collectedFrom.digest.slice(0, 19)).join(',')}，需要重新采集`);
  }
  lines.push(`图表实体: ${summary.charts} 个（已核实 ${summary.confirmedCharts}，未确认 ${summary.unconfirmedCharts}）；每页引用 ${summary.referenceCount} 处；外部链接 ${summary.externalLinks} 个（均未访问）`);
  if (chartId !== null) {
    const detail = chartDetail(collection, chartId);
    if (!detail) {
      lines.push(`找不到图表实体 ${chartId}`);
      return lines.join('\n');
    }
    lines.push(`图表 ${detail.id}: ${detail.part}，状态 ${detail.status}，${detail.confirmed ? '已核实' : '未确认'}`);
    lines.push(`被引用页: ${detail.references.map(value => value.page).join(',') || '无'}`);
  }
  const views = chartId !== null ? (chartDetail(collection, chartId)?.references ?? []) : page === null ? collection.references : referencesFor(collection, {page});
  for (const reference of views) {
    const chart = reference.entity ? collection.entities.charts[reference.entity] : null;
    const title = chart?.title ? `「${chart.title}」` : '';
    const shared = reference.shared ? '，与其他页共用实体' : '';
    lines.push(`- 第 ${reference.page} 页 ${reference.frame.name ?? reference.frame.id ?? '图形'}${title} → ${reference.chartPart ?? '(未解析)'} [${reference.status}${shared}] 引用 ${reference.id}`);
    if (!chart) {
      for (const reason of reference.reasons) lines.push(`    ! ${reason.code}: ${reason.message}`);
      continue;
    }
    for (const reason of chart.reasons) lines.push(`    ! 未确认 ${reason.code}: ${reason.message}`);
    for (const warning of chart.warnings) lines.push(`    ⚠ ${warning.code}: ${warning.message}`);
    for (const series of chart.series) {
      const range = series.value?.reference ? `${series.value.reference.sheet}!${series.value.reference.range}` : (series.value?.formula ?? '无引用');
      lines.push(`    系列 ${series.index + 1} ${series.name?.value ?? '(无名)'}: ${range} 数值 ${JSON.stringify(series.value.counts)} 类别 ${JSON.stringify(series.category.counts)}`);
      for (const point of series.points) {
        lines.push(`      [${point.index}] ${point.value.status}: 缓存 ${formatValue(point.value.display)} @ ${point.value.display.source?.pointer ?? '—'} | 工作簿 ${formatValue(point.value.workbook)} @ ${formatCell(point.value.workbook.source)}`);
      }
      if (!series.confirmed) for (const reason of series.reasons) if (!chart.reasons.some(value => value.message.endsWith(reason.message))) lines.push(`    ! 系列 ${series.index + 1} ${reason.code}: ${reason.message}`);
    }
  }
  if (page === null && chartId === null && collection.history?.length) {
    lines.push(`历史: ${collection.history.map(entry => `${entry.collectedAt.slice(0, 19)} 采集第 ${entry.pages.join(',')} 页`).join('；')}`);
  }
  return lines.join('\n');
}

function formatValue(side) {
  if (!side || !side.present) return '空';
  return JSON.stringify(side.value);
}

function formatCell(source) {
  if (!source) return '—';
  if (source.cell) return `${source.sheet ?? source.part ?? ''}!${source.cell}`;
  return source.part ?? '—';
}

export async function loadPackage({file, sourceUrl} = {}) {
  if (file) {
    const bytes = await readFile(resolve(file));
    return {bytes, packageName: basename(resolve(file)), source: 'file'};
  }
  if (sourceUrl) {
    const response = await fetch(new URL('/api/package', sourceUrl));
    if (!response.ok) throw new Error(`资料源返回 ${response.status}`);
    return {bytes: Buffer.from(await response.arrayBuffer()), packageName: 'capacity.pptx', source: sourceUrl};
  }
  const source = await startSource();
  try {
    const response = await fetch(source.url + '/api/package');
    if (!response.ok) throw new Error(`资料源返回 ${response.status}`);
    return {bytes: Buffer.from(await response.arrayBuffer()), packageName: 'capacity.pptx', source: source.url, close: () => source.close()};
  } catch (error) {
    await source.close();
    throw error;
  }
}

const OPTIONS = {
  pages: {type: 'string'},
  file: {type: 'string'},
  source: {type: 'string'},
  capture: {type: 'string'},
  json: {type: 'boolean', default: false},
  page: {type: 'string'},
  chart: {type: 'string'},
};

export async function main(argv = process.argv.slice(2)) {
  const {values} = parseArgs({args: argv, options: OPTIONS, allowPositionals: false});
  const pages = values.pages ? parsePageSelection(values.pages) : null;
  if (values.page !== undefined && !/^[1-9][0-9]*$/.test(values.page)) throw new Error(`--page 需要正整数页号: ${values.page}`);
  const loaded = await loadPackage({file: values.file, sourceUrl: values.source});
  try {
    const path = resolve(values.capture ?? process.env.CAPTURE_PATH ?? 'var/captures.json');
    const store = new CaptureStore(path);
    const {record, collection} = await receivePresentation({
      store,
      bytes: loaded.bytes,
      packageName: loaded.packageName,
      pages,
    });
    if (values.json) {
      console.log(JSON.stringify({captureFile: path, revision: record.revision, package: collection.package, selection: collection.selection, summary: collection.summary, collection}, null, 2));
    } else {
      console.log(`接收文件: ${path}（第 ${record.revision} 次记录，页码 ${pages ? pages.join(',') : '全部'}）`);
      console.log(describeCollection(collection, {page: values.page ? Number(values.page) : null, chartId: values.chart ?? null}));
    }
    return collection;
  } finally {
    if (typeof loaded.close === 'function') await loaded.close();
  }
}

if (import.meta.url === `file://${process.argv[1]}`) {
  main().catch(error => {
    console.error(`采集失败: ${error.message}`);
    process.exitCode = 1;
  });
}
