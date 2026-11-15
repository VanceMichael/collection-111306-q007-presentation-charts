import test from 'node:test';
import assert from 'node:assert/strict';
import {mkdtemp, readFile, rm, writeFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {CaptureStore} from '../src/store.mjs';
import {receivePresentation, describeCollection} from '../src/receive.mjs';
import {collectPresentation} from '../src/collect.mjs';
import {ZipArchive} from '../src/zip.mjs';
import {buildZip} from './helpers/package.mjs';

const fixture = await readFile(new URL('../fixtures/capacity.pptx', import.meta.url));

async function withStore(run) {
  const root = await mkdtemp(join(tmpdir(), 'receive-'));
  try {
    const path = join(root, 'captures.json');
    return await run(new CaptureStore(path), path);
  } finally {
    await rm(root, {recursive: true, force: true});
  }
}

test('接收记录包含每页来源：图表摘要、工作簿摘要、单元格与缓存指针', async () => {
  await withStore(async (store) => {
    const {record, collection} = await receivePresentation({store, bytes: fixture, packageName: 'capacity.pptx'});
    assert.equal(record.kind, 'presentation-charts');
    assert.equal(record.revision, 1);
    const chart = Object.values(collection.entities.charts)[0];
    const series = chart.series[0];
    assert.equal(series.value.points[0].display.source.part, 'ppt/charts/chart1.xml');
    assert.equal(series.value.points[0].workbook.source.cell, 'B1');
    assert.equal(series.value.points[0].workbook.source.package, Object.keys(collection.entities.workbooks)[0]);
    const onDisk = JSON.parse(await readFile(join(store.path.split('/').slice(0, -1).join('/'), 'captures.json'), 'utf8'));
    const stored = JSON.parse(onDisk.body);
    assert.equal(stored.records[0].collection.references.length, 2);
    assert.equal(stored.records[0].collection.entities.charts[chart.id].series[0].points[0].value.workbook.source.cell, 'B1');
    assert.equal(stored.records[0].collection.entities.charts[chart.id].series[0].points[0].category.display.value, '一号产线');
  });
});

test('分次接收：先第 1 页后补第 2 页，记录合并且每页引用可查', async () => {
  await withStore(async (store) => {
    const first = await receivePresentation({store, bytes: fixture, packageName: 'capacity.pptx', pages: [1]});
    assert.deepEqual(first.collection.pages.map(page => page.page), [1]);
    const second = await receivePresentation({store, bytes: fixture, packageName: 'capacity.pptx', pages: [2]});
    assert.deepEqual(second.collection.pages.map(page => page.page), [1, 2]);
    assert.equal(second.record.revision, 2);
    assert.equal(second.collection.references.length, 2);
    assert.deepEqual(second.collection.references.map(reference => reference.page), [1, 2]);
    const state = await store.read();
    assert.equal(state.records.length, 2, '每次接收各留一条记录，不覆盖历史');
    assert.deepEqual(state.records[1].collection.pages.map(page => page.page), [1, 2], '最新记录已合并两页');
    const text = describeCollection(second.collection);
    assert.match(text, /每页引用 2 处/);
    assert.match(text, /page:2\/frame:1\/chart:1/);
  });
});

test('重复接收同一页时替换该页而不重复计数', async () => {
  await withStore(async (store) => {
    await receivePresentation({store, bytes: fixture, packageName: 'capacity.pptx'});
    const again = await receivePresentation({store, bytes: fixture, packageName: 'capacity.pptx', pages: [1]});
    assert.equal(again.collection.references.length, 2);
    assert.equal(again.collection.pages.length, 2);
    assert.equal(again.collection.summary.charts, 1);
    assert.equal(again.record.revision, 2);
  });
});

test('新文稿补页保留旧页但标注旧来源', async () => {
  await withStore(async (store) => {
    await receivePresentation({store, bytes: fixture, packageName: 'capacity.pptx', pages: [1]});
    const zip = new ZipArchive(fixture);
    const parts = {};
    for (const name of zip.names()) parts[name] = zip.read(name);
    parts['ppt/slides/slide3.xml'] = '<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>';
    const changed = buildZip(parts);
    const second = await receivePresentation({store, bytes: changed, packageName: 'capacity.pptx', pages: [2]});
    assert.deepEqual(second.collection.pages.map(page => page.page), [1, 2]);
    assert.equal(second.collection.summary.stalePages.length, 1);
    assert.equal(second.collection.summary.stalePages[0].page, 1);
    assert.equal(second.collection.packageChanges.length, 1);
  });
});

test('页码解析与空选择不会静默当成全部', async () => {
  await withStore(async (store) => {
    await assert.rejects(receivePresentation({store, bytes: fixture, packageName: 'capacity.pptx', pages: [99]}), /页码均不存在/);
    const partial = await receivePresentation({store, bytes: fixture, packageName: 'capacity.pptx', pages: [1, 99]});
    assert.deepEqual(partial.collection.selection.missing, [99]);
    assert.deepEqual(partial.collection.pages.map(page => page.page), [1]);
    const missing = collectPresentation(fixture, {packageName: 'capacity.pptx', pages: [99]});
    assert.deepEqual(missing.selection.missing, [99]);
    assert.equal(missing.references.length, 0);
    assert.equal(missing.summary.charts, 0);
    assert.equal(missing.summary.selected, 0);
  });
});

test('描述输出列出每个数据点的缓存指针与工作簿单元格', async () => {
  const {collection} = await (async () => {
    const root = await mkdtemp(join(tmpdir(), 'describe-'));
    try {
      return await receivePresentation({store: new CaptureStore(join(root, 'c.json')), bytes: fixture, packageName: 'capacity.pptx'});
    } finally { await rm(root, {recursive: true, force: true}); }
  })();
  const text = describeCollection(collection);
  assert.match(text, /\[0\] conflict: 缓存 105 @ \/chartSpace\/chart\/plotArea\/barChart\[1\]\/ser\[0\]\/val\/numRef\/numCache\/pt\[@idx=0\] \| 工作簿 108 @ 产能!B1/);
  assert.match(text, /\[1\] missing: 缓存 空 @ — \| 工作簿 空 @ 产能!B2/);
  assert.match(text, /外部链接 1 个（均未访问）/);
});

test('接收文件损坏时不静默覆盖或返回空结果', async () => {
  const root = await mkdtemp(join(tmpdir(), 'receive-'));
  try {
    const path = join(root, 'captures.json');
    await writeFile(path, JSON.stringify({body: '{"revision":0,"records":[]}', checksum: 'x'}));
    const store = new CaptureStore(path);
    await assert.rejects(receivePresentation({store, bytes: fixture, packageName: 'capacity.pptx'}), /校验/);
  } finally { await rm(root, {recursive: true, force: true}); }
});
