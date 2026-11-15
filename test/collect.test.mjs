import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {collectPresentation, parsePageSelection, referencesFor, mergeCollection} from '../src/collect.mjs';
import {presentationPackage, workbookParts, chartXml, slideXml, rels, buildZip} from './helpers/package.mjs';
import {ZipArchive} from '../src/zip.mjs';

const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const fixture = await readFile(new URL('../fixtures/capacity.pptx', import.meta.url));

test('真实封装文稿：两页共用一张图，每页引用都可查，冲突与缺格如实记录', () => {
  const result = collectPresentation(fixture, {packageName: 'capacity.pptx'});
  assert.equal(result.pages.length, 2);
  assert.equal(result.references.length, 2);
  assert.deepEqual(referencesFor(result, {page: 1}).map(value => value.chartPart), ['ppt/charts/chart1.xml']);
  assert.deepEqual(referencesFor(result, {page: 2}).map(value => value.chartPart), ['ppt/charts/chart1.xml']);
  assert.equal(Object.keys(result.entities.charts).length, 1, '同一张图只保留一个图表实体');
  const entityId = Object.keys(result.entities.charts)[0];
  assert.equal(referencesFor(result, {chart: entityId}).length, 2, '共享实体仍能从两页分别查到');
  assert.equal(result.references[1].shared, true, '第二页引用与第一页共用实体');
  assert.equal(result.references[1].status, 'unconfirmed');
  assert.equal(result.references[1].entity, result.references[0].entity);
  const chart = result.entities.charts[entityId];
  assert.equal(chart.confirmed, false);
  const series = chart.series[0];
  assert.equal(series.value.reference.sheet, '产能');
  assert.equal(series.value.reference.range, 'B1:B3');
  assert.equal(series.value.counts.conflict, 1);
  assert.equal(series.value.counts.missing, 1);
  const conflict = series.points[0];
  assert.equal(conflict.status, 'conflict');
  assert.equal(conflict.value.display.value, 105);
  assert.equal(conflict.value.display.source.part, 'ppt/charts/chart1.xml');
  assert.equal(conflict.value.display.source.pointer, '/chartSpace/chart/plotArea/barChart[1]/ser[0]/val/numRef/numCache/pt[@idx=0]');
  assert.equal(conflict.value.workbook.value, 108);
  assert.equal(conflict.value.workbook.source.sheet, '产能');
  assert.equal(conflict.value.workbook.source.cell, 'B1');
  const gap = series.points[1];
  assert.equal(gap.status, 'missing');
  assert.equal(gap.value.display.present, false);
  assert.equal(gap.value.workbook.present, false);
  assert.equal(gap.value.workbook.cell, 'B2');
  assert.equal(series.points.length, 3, '按引用范围保留三个位置，不补零');
  assert.equal(series.name.cell, 'C1');
  assert.equal(series.name.hiddenColumn, true);
  assert.ok(chart.warnings.some(value => value.code === 'hiddenColumnSource'));
  assert.ok(result.summary.unconfirmedCharts === 1);
  assert.ok(result.summary.blockingReasons.some(value => value.code === 'valueConflict'));
});

test('外部链接只记录不访问，工作簿数值与显示缓存分列', () => {
  const result = collectPresentation(fixture, {packageName: 'capacity.pptx'});
  assert.equal(result.externalLinks.length, 1);
  assert.equal(result.externalLinks[0].target, 'https://supplier.invalid/ledger');
  assert.equal(result.externalLinks[0].followed, false);
  const chart = Object.values(result.entities.charts)[0];
  assert.equal(chart.externalData.followed, false);
  assert.equal(chart.externalData.target, 'ppt/embeddings/ledger.xlsx');
  assert.equal(chart.externalData.rawTarget, '../embeddings/ledger.xlsx');
  assert.equal(Object.keys(result.entities.workbooks).length, 1);
});

test('页码选择：只采集选定页，缺失页单独报告', () => {
  assert.deepEqual(parsePageSelection('2'), [2]);
  assert.deepEqual(parsePageSelection('2,1-2'), [1, 2]);
  assert.throws(() => parsePageSelection('0'), /页码范围无效/);
  assert.throws(() => parsePageSelection(''), /页码选择为空/);
  assert.throws(() => parsePageSelection('x'), /页码格式无效/);
  const only = collectPresentation(fixture, {packageName: 'capacity.pptx', pages: [2]});
  assert.deepEqual(only.pages.map(value => value.page), [2]);
  assert.deepEqual(only.selection.resolved.map(value => value.page), [2]);
  assert.equal(only.references.length, 1);
  assert.equal(only.selection.missing.length, 0);
  const missing = collectPresentation(fixture, {packageName: 'capacity.pptx', pages: [5]});
  assert.deepEqual(missing.selection.missing, [5]);
  assert.equal(missing.references.length, 0);
});

test('补页合并：后一页追加，前一页不被清掉，图表实体共享', () => {
  const first = collectPresentation(fixture, {packageName: 'capacity.pptx', pages: [1]});
  const second = collectPresentation(fixture, {packageName: 'capacity.pptx', pages: [2]});
  const merged = mergeCollection(first, second);
  assert.deepEqual(merged.pages.map(value => value.page), [1, 2]);
  assert.equal(merged.references.length, 2);
  assert.equal(Object.keys(merged.entities.charts).length, 1);
  assert.equal(Object.keys(merged.entities.workbooks).length, 1);
  assert.equal(merged.summary.charts, 1);
  assert.equal(merged.summary.referenceCount, 2);
});

function buildBase({slides, charts, workbooks, extraParts} = {}) {
  return presentationPackage({slides, charts, workbooks, extraParts});
}

const LEDGER = () => workbookParts({
  sheets: [{
    name: '产能',
    xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>甲线</t></is></c><c r="B1"><v>10</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>乙线</t></is></c><c r="B2"><v>20</v></c></row></sheetData></worksheet>',
  }],
});

test('单图损坏隔离：损坏的图不影响另一张已核实图', () => {
  const good = chartXml({
    series: [{cat: {formula: '产能!$A$1:$A$2', points: [{idx: 0, value: '甲线'}, {idx: 1, value: '乙线'}]}, val: {formula: '产能!$B$1:$B$2', points: [{idx: 0, value: 10}, {idx: 1, value: 20}]}}],
  });
  const broken = 'this is not XML <<<';
  const workbook = LEDGER();
  const packageBytes = buildBase({
    slides: [{
      charts: [{id: 'rIdGood', name: '好图'}, {id: 'rIdBad', name: '坏图'}],
      rels: [
        {id: 'rIdGood', type: `${REL}/chart`, target: '../charts/good.xml'},
        {id: 'rIdBad', type: `${REL}/chart`, target: '../charts/bad.xml'},
      ],
    }],
    charts: {'good.xml': good, 'bad.xml': broken},
    workbooks: {'ppt/embeddings/ledger.xlsx': workbook},
    extraParts: {
      'ppt/charts/_rels/good.xml.rels': rels([{id: 'rIdBook', type: `${REL}/package`, target: '../embeddings/ledger.xlsx'}]),
      'ppt/charts/_rels/bad.xml.rels': rels([{id: 'rIdBook', type: `${REL}/package`, target: '../embeddings/ledger.xlsx'}]),
    },
  });
  const result = collectPresentation(packageBytes, {packageName: 'mixed.pptx'});
  const charts = Object.values(result.entities.charts);
  assert.equal(charts.length, 2);
  const goodChart = charts.find(chart => chart.part.endsWith('good.xml'));
  const badChart = charts.find(chart => chart.part.endsWith('bad.xml'));
  assert.equal(goodChart.status, 'parsed');
  assert.equal(badChart.status, 'corrupt');
  assert.equal(goodChart.confirmed, true);
  assert.equal(badChart.confirmed, false);
  assert.ok(badChart.reasons.some(value => value.code === 'chartUnreadable'));
  assert.equal(result.references.find(value => value.chartPart.endsWith('bad.xml')).status, 'corrupt');
  assert.equal(result.references.find(value => value.chartPart.endsWith('good.xml')).status, 'confirmed');
});

test('缺格、隐藏列与不同长度如实记录，类别和数值分开说明', () => {
  const workbook = workbookParts({
    sheets: [{
      name: '产能',
      xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><cols><col min="1" max="1" hidden="1"/></cols><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>甲线</t></is></c><c r="B1"><v>4</v></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>乙线</t></is></c></row><row r="3"><c r="B3"><v>6</v></c></row></sheetData></worksheet>',
    }],
  });
  const chart = chartXml({series: [{cat: {formula: '产能!$A$1:$A$2', points: [{idx: 0, value: '甲线'}, {idx: 1, value: '乙线'}]}, val: {formula: '产能!$B$1:$B$3', points: [{idx: 0, value: 4}, {idx: 2, value: 6}]}}]});
  const bytes = buildBase({
    slides: [{charts: [{id: 'rIdC'}], rels: [{id: 'rIdC', type: `${REL}/chart`, target: '../charts/c1.xml'}]}],
    charts: {'c1.xml': chart},
    workbooks: {'ppt/embeddings/ledger.xlsx': workbook},
    extraParts: {'ppt/charts/_rels/c1.xml.rels': rels([{id: 'rIdBook', type: `${REL}/package`, target: '../embeddings/ledger.xlsx'}])},
  });
  const result = collectPresentation(bytes, {packageName: 'gaps.pptx'});
  const parsed = Object.values(result.entities.charts)[0];
  const series = parsed.series[0];
  assert.equal(series.value.counts.missing, 1, 'B2 在缓存与工作簿中都没有内容');
  assert.equal(series.value.counts.match, 2);
  assert.equal(series.value.points.length, 3, '按引用范围保留三个位置，不补零');
  assert.equal(series.confirmed, false);
  assert.equal(series.category.counts.match, 2);
  assert.equal(series.category.counts.missing, 0);
  assert.deepEqual(series.points.map(point => point.categoryStatus), ['match', 'match', 'missing']);
  assert.ok(parsed.reasons.some(value => value.code === 'valueGap'));
  assert.ok(parsed.warnings.some(value => value.code === 'hiddenColumnSource'), '类别引用隐藏列 A 必须报告');
  assert.equal(series.points[1].value.workbook.cell, 'B2');
  assert.equal(series.points[1].value.workbook.present, false);
  assert.equal(series.points[1].value.display.present, false);
  assert.equal(series.category.points[0].workbook.hiddenColumn, true);
});

test('字面量数据不当作单元格引用', () => {
  const chart = chartXml({series: [{cat: {literal: true, points: [{idx: 0, value: '甲'}, {idx: 1, value: '乙'}]}, val: {literal: true, points: [{idx: 0, value: 7}, {idx: 1, value: 9}]}}]});
  const bytes = buildBase({
    slides: [{charts: [{id: 'rIdC'}], rels: [{id: 'rIdC', type: `${REL}/chart`, target: '../charts/c1.xml'}]}],
    charts: {'c1.xml': chart},
    extraParts: {'ppt/charts/_rels/c1.xml.rels': rels([{id: 'rIdBook', type: `${REL}/package`, target: '../embeddings/ledger.xlsx'}, {id: 'rIdX', type: `${REL}/hyperlink`, target: 'https://supplier.invalid/x', external: true}])},
  });
  const result = collectPresentation(bytes, {packageName: 'literal.pptx'});
  const parsed = Object.values(result.entities.charts)[0];
  assert.equal(parsed.series[0].value.sourceKind, 'numLit');
  assert.equal(parsed.series[0].value.workbook.status, 'literal');
  assert.equal(parsed.series[0].value.points[0].status, 'literal');
  assert.ok(parsed.warnings.some(value => value.code === 'literalData'));
  assert.equal(parsed.confirmed, false, '字面量数据没有工作簿可核对，不能确认');
  assert.ok(parsed.reasons.some(value => value.code === 'literalValueUncorroborated'));
  assert.equal(result.externalLinks[0].target, 'https://supplier.invalid/x');
  assert.equal(result.externalLinks[0].followed, false);
});

test('行优先与列优先方向均按引用区域逐点对应', () => {
  const workbook = workbookParts({
    sheets: [{name: '产能', xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row></sheetData></worksheet>'}],
  });
  const chart = chartXml({series: [{val: {formula: '产能!$A$1:$C$1', points: [{idx: 0, value: 1}, {idx: 1, value: 2}, {idx: 2, value: 3}]}}]});
  const bytes = buildBase({
    slides: [{charts: [{id: 'rIdC'}], rels: [{id: 'rIdC', type: `${REL}/chart`, target: '../charts/c1.xml'}]}],
    charts: {'c1.xml': chart},
    workbooks: {'ppt/embeddings/ledger.xlsx': workbook},
    extraParts: {'ppt/charts/_rels/c1.xml.rels': rels([{id: 'rIdBook', type: `${REL}/package`, target: '../embeddings/ledger.xlsx'}])},
  });
  const result = collectPresentation(bytes, {packageName: 'horizontal.pptx'});
  const series = Object.values(result.entities.charts)[0].series[0];
  assert.deepEqual(series.value.points.map(point => point.workbook.cell), ['A1', 'B1', 'C1']);
  assert.equal(series.value.counts.match, 3);
});

test('引用多行多列时不臆测对应关系', () => {
  const workbook = workbookParts({
    sheets: [{name: '产能', xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row><row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row></sheetData></worksheet>'}],
  });
  const chart = chartXml({series: [{val: {formula: '产能!$A$1:$B$2', points: [{idx: 0, value: 1}, {idx: 1, value: 2}]}}]});
  const bytes = buildBase({
    slides: [{charts: [{id: 'rIdC'}], rels: [{id: 'rIdC', type: `${REL}/chart`, target: '../charts/c1.xml'}]}],
    charts: {'c1.xml': chart},
    workbooks: {'ppt/embeddings/ledger.xlsx': workbook},
    extraParts: {'ppt/charts/_rels/c1.xml.rels': rels([{id: 'rIdBook', type: `${REL}/package`, target: '../embeddings/ledger.xlsx'}])},
  });
  const result = collectPresentation(bytes, {packageName: 'matrix.pptx'});
  const parsed = Object.values(result.entities.charts)[0];
  const series = parsed.series[0];
  assert.equal(series.value.workbook.status, 'ambiguousRange');
  assert.ok(parsed.reasons.some(value => value.code === 'ambiguousRange'));
  assert.ok(series.points.every(point => point.value.workbook.present === false && point.value.workbook.cell === null));
  assert.ok(series.points.every(point => point.value.status === 'unmapped'));
});

test('损坏的工作簿不阻止图表结构被记录', () => {
  const goodWorkbook = LEDGER();
  const chartGood = chartXml({series: [{cat: {formula: '产能!$A$1:$A$2', points: [{idx: 0, value: '产量'}, {idx: 1, value: '乙线'}]}, val: {formula: '产能!$B$1:$B$2', points: [{idx: 0, value: 10}, {idx: 1, value: 20}]}}]});
  const bytes = buildBase({
    slides: [{charts: [{id: 'rIdC'}], rels: [{id: 'rIdC', type: `${REL}/chart`, target: '../charts/c1.xml'}]}],
    charts: {'c1.xml': chartGood},
    workbooks: {
      'ppt/embeddings/ledger.xlsx': goodWorkbook,
      'ppt/embeddings/other.bin': Buffer.from('not a zip'),
    },
    extraParts: {'ppt/charts/_rels/c1.xml.rels': rels([{id: 'rIdBook', type: `${REL}/package`, target: '../embeddings/other.bin'}])},
  });
  const result = collectPresentation(bytes, {packageName: 'badworkbook.pptx'});
  const parsed = Object.values(result.entities.charts)[0];
  assert.equal(parsed.series[0].value.workbook.status, 'unreadable');
  assert.ok(parsed.reasons.some(value => value.code === 'workbookUnavailable'));
  assert.ok(parsed.series[0].points.every(point => point.value.status === 'unmapped'));
  assert.equal(parsed.series[0].value.points[0].display.source.part, 'ppt/charts/c1.xml');
});

test('包结构损坏时抛出明确错误，而不是返回空结果', () => {
  assert.throws(() => collectPresentation(Buffer.from('not a package')), /ZIP|中央目录/);
  const zip = new ZipArchive(fixture);
  assert.ok(zip.has('ppt/slides/slide1.xml'));
  assert.throws(() => zip.read('ppt/nope.xml'));
});

test('外部链接与外部工作簿只记录，不发起任何网络请求', async () => {
  const http = await import('node:http');
  let hits = 0;
  const server = http.createServer((request, response) => { hits++; response.end('secret'); });
  await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
  const url = `http://127.0.0.1:${server.address().port}/ledger.xlsx`;
  try {
    const bytes = buildBase({
      slides: [{charts: [{id: 'rIdC'}], rels: [{id: 'rIdC', type: `${REL}/chart`, target: '../charts/c1.xml'}, {id: 'rIdS', type: `${REL}/hyperlink`, target: url, external: true}]}],
      charts: {'c1.xml': chartXml({series: [{val: {formula: '产能!$B$1:$B$2', points: [{idx: 0, value: 1}, {idx: 1, value: 2}]}}]})},
      extraParts: {'ppt/charts/_rels/c1.xml.rels': rels([{id: 'rIdBook', type: `${REL}/hyperlink`, target: url, external: true}])},
    });
    const result = collectPresentation(bytes, {packageName: 'external.pptx'});
    assert.equal(hits, 0, '采集过程不得访问外部链接');
    assert.ok(result.externalLinks.length >= 2);
    assert.ok(result.externalLinks.every(link => link.followed === false));
    assert.ok(result.externalLinks.every(link => link.target === url));
    const chart = Object.values(result.entities.charts)[0];
    assert.equal(chart.series[0].value.workbook.status, 'external');
    assert.ok(chart.reasons.some(value => value.code === 'externalWorkbook'));
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});

test('XML 不执行脚本、不展开 DOCTYPE 或外部实体', async () => {
  const chart = `<?xml version="1.0"?><!DOCTYPE foo [<!ENTITY xxe SYSTEM "file:///etc/passwd">]><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart/></c:chartSpace>`;
  const bytes = buildBase({
    slides: [{charts: [{id: 'rIdC'}], rels: [{id: 'rIdC', type: `${REL}/chart`, target: '../charts/c1.xml'}]}],
    charts: {'c1.xml': chart},
  });
  const result = collectPresentation(bytes, {packageName: 'xxe.pptx'});
  const parsed = Object.values(result.entities.charts)[0];
  assert.equal(parsed.status, 'corrupt');
  assert.match(parsed.reasons[0].message, /DOCTYPE/);
});

test('ZIP 成员级损坏只影响该成员，其他图仍可核实', () => {
  const good = chartXml({series: [{cat: {formula: '产能!$A$1:$A$2', points: [{idx: 0, value: '甲线'}, {idx: 1, value: '乙线'}]}, val: {formula: '产能!$B$1:$B$2', points: [{idx: 0, value: 10}, {idx: 1, value: 20}]}}]});
  const goodWorkbook = LEDGER();
  const bytes = buildBase({
    slides: [{charts: [{id: 'rIdG', name: '好图'}, {id: 'rIdB', name: '坏图'}]}],
    charts: {'good.xml': good, 'bad.xml': good.replace('产能!$B$1:$B$2', '产能!$B$1:$B$2')},
    workbooks: {'ppt/embeddings/ledger.xlsx': goodWorkbook},
  });
  // add slide rels for both charts
  const zip = new ZipArchive(bytes);
  const parts = Object.fromEntries(zip.names().map(name => [name, zip.read(name)]));
  parts['ppt/slides/_rels/slide1.xml.rels'] = rels([
    {id: 'rIdG', type: `${REL}/chart`, target: '../charts/good.xml'},
    {id: 'rIdB', type: `${REL}/chart`, target: '../charts/bad.xml'},
  ]);
  parts['ppt/charts/_rels/good.xml.rels'] = rels([{id: 'rIdBook', type: `${REL}/package`, target: '../embeddings/ledger.xlsx'}]);
  parts['ppt/charts/_rels/bad.xml.rels'] = rels([{id: 'rIdBook', type: `${REL}/package`, target: '../embeddings/ledger.xlsx'}]);
  const packed = buildZip(parts);
  // Flip a byte inside bad.xml's compressed payload to make that member unreadable.
  const archive = new ZipArchive(packed);
  const entry = archive.entry('ppt/charts/bad.xml');
  const localNameSize = packed.readUInt16LE(entry.localOffset + 26);
  const localExtraSize = packed.readUInt16LE(entry.localOffset + 28);
  const dataStart = entry.localOffset + 30 + localNameSize + localExtraSize;
  const damaged = Buffer.from(packed);
  damaged[dataStart + 5] ^= 0xff;
  const result = collectPresentation(damaged, {packageName: 'zipdamage.pptx'});
  const charts = Object.values(result.entities.charts);
  const goodChart = charts.find(chart => chart.part.endsWith('good.xml'));
  const badChart = charts.find(chart => chart.part.endsWith('bad.xml'));
  assert.equal(goodChart.status, 'parsed');
  assert.equal(goodChart.confirmed, true);
  assert.equal(badChart.status, 'corrupt');
  const badReference = result.references.find(reference => reference.chartPart.endsWith('bad.xml'));
  assert.equal(badReference.status, 'corrupt');
  assert.ok(badReference.reasons.some(reason => reason.code === 'corruptChartPart'));
});

test('共享字符串与大小写不一致的工作表名按实际内容核对', () => {
  const workbook = workbookParts({
    sharedStrings: ['甲线', '乙线'],
    sheets: [{name: 'Capacity', xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1"><v>5</v></c><c r="C1" t="s"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c><c r="B2"><v>6</v></c></row></sheetData></worksheet>'}],
  });
  const chart = chartXml({series: [{name: {formula: 'capacity!$C$1', points: [{idx: 0, value: '乙线'}]}, cat: {formula: 'capacity!$A$1:$A$2', points: [{idx: 0, value: '甲线'}, {idx: 1, value: '乙线'}]}, val: {formula: 'capacity!$B$1:$B$2', points: [{idx: 0, value: 5}, {idx: 1, value: 6}]}}]});
  const bytes = buildBase({
    slides: [{charts: [{id: 'rIdC'}], rels: [{id: 'rIdC', type: `${REL}/chart`, target: '../charts/c1.xml'}]}],
    charts: {'c1.xml': chart},
    workbooks: {'ppt/embeddings/ledger.xlsx': workbook},
    extraParts: {'ppt/charts/_rels/c1.xml.rels': rels([{id: 'rIdBook', type: `${REL}/package`, target: '../embeddings/ledger.xlsx'}])},
  });
  const result = collectPresentation(bytes, {packageName: 'shared.pptx'});
  const parsed = Object.values(result.entities.charts)[0];
  const series = parsed.series[0];
  assert.equal(series.value.counts.match, 2);
  assert.equal(series.name.cellValue, '乙线');
  assert.equal(series.name.comparison, 'match');
  assert.equal(parsed.confirmed, true, '全部数值与类别均核对一致时图表可确认');
  assert.ok(parsed.warnings.some(value => value.code === 'sheetNameCaseMismatch'));
});

test('隐藏工作表与公式无缓存值会阻止确认', () => {
  const workbook = workbookParts({
    sheets: [{name: '产能', state: 'hidden', xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="B1"><f>SUM(1,2)</f></c></row></sheetData></worksheet>'}],
  });
  const chart = chartXml({series: [{val: {formula: '产能!$B$1', points: [{idx: 0, value: 3}]}}]});
  const bytes = buildBase({
    slides: [{charts: [{id: 'rIdC'}], rels: [{id: 'rIdC', type: `${REL}/chart`, target: '../charts/c1.xml'}]}],
    charts: {'c1.xml': chart},
    workbooks: {'ppt/embeddings/ledger.xlsx': workbook},
    extraParts: {'ppt/charts/_rels/c1.xml.rels': rels([{id: 'rIdBook', type: `${REL}/package`, target: '../embeddings/ledger.xlsx'}])},
  });
  const result = collectPresentation(bytes, {packageName: 'hidden.pptx'});
  const parsed = Object.values(result.entities.charts)[0];
  const series = parsed.series[0];
  assert.equal(series.value.points[0].workbook.formula, 'SUM(1,2)');
  assert.equal(series.value.points[0].workbook.present, false);
  assert.equal(series.value.points[0].status, 'displayOnly');
  assert.ok(parsed.warnings.some(value => value.code === 'worksheetHidden'));
  assert.ok(parsed.reasons.some(value => value.code === 'formulaWithoutCachedValue'));
  assert.equal(parsed.confirmed, false);
});

test('宏与嵌入程序只登记不执行', () => {
  const bytes = buildBase({
    slides: [{charts: []}],
    charts: {},
    extraParts: {
      'ppt/vbaProject.bin': Buffer.from('macro bytes'),
      'ppt/embeddings/oleObject1.bin': Buffer.from('ole bytes'),
    },
  });
  const result = collectPresentation(bytes, {packageName: 'macro.pptx'});
  assert.equal(result.activeContent.length, 2);
  assert.ok(result.activeContent.every(part => part.executed === false && part.read === false));
  assert.ok(result.activeContent.some(part => part.kind === 'macro'));
  assert.equal(result.summary.activeContentParts, 2);
});

test('重复引用同一实体在多次补页后仍只有一份实体和一份工作簿', () => {
  const first = collectPresentation(fixture, {packageName: 'capacity.pptx', pages: [1]});
  const second = collectPresentation(fixture, {packageName: 'capacity.pptx', pages: [2]});
  const third = collectPresentation(fixture, {packageName: 'capacity.pptx', pages: [1]});
  const merged = mergeCollection(mergeCollection(first, second), third);
  assert.equal(merged.references.length, 2);
  assert.equal(Object.keys(merged.entities.charts).length, 1);
  assert.equal(Object.keys(merged.entities.workbooks).length, 1);
  assert.deepEqual(merged.pages.map(page => page.page), [1, 2]);
  assert.equal(merged.summary.counts.value.conflict, 1, '重复补页不重复累计冲突');
});
