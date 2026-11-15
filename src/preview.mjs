import {resolve} from 'node:path';
import {startSource} from './source.mjs';
import {CaptureStore} from './store.mjs';
import {receivePresentation, describeCollection} from './receive.mjs';
import {referencesFor, chartDetail} from './collect.mjs';
import {preview} from './domain.mjs';

/**
 * Short demonstration against the real local source: receive page 1, then
 * supplement with page 2, showing shared chart/workbook entities, the exact
 * per-point provenance, and what could not be confirmed. Nothing is replayed
 * from old logs; each round re-collects from the package bytes.
 */
async function run() {
  const source = await startSource();
  const path = resolve(process.env.CAPTURE_PATH ?? 'var/captures.json');
  try {
    const observation = preview(await (await fetch(source.url + '/api/snapshot')).json());
    console.log('== 原始资料预览（保留原有预览调用）==');
    console.log(JSON.stringify(observation));
    const response = await fetch(source.url + '/api/package');
    if (!response.ok) throw new Error(`资料源返回 ${response.status}`);
    const bytes = Buffer.from(await response.arrayBuffer());
    const store = new CaptureStore(path);
    const first = await receivePresentation({store, bytes, packageName: 'capacity.pptx', pages: [1]});
    console.log('\n== 第一轮：有选择地接收第 1 页 ==');
    console.log(`资料源: ${source.url}`);
    console.log(`接收文件: ${path}（第 ${first.record.revision} 次记录）`);
    console.log(`原始预览记录: 包 ${observation.package}，源端列出的页 ${observation.selected.join(',')}（仅作观察，不代替采集）`);
    console.log(describeCollection(first.collection));
    const supplement = await receivePresentation({store, bytes, packageName: 'capacity.pptx', pages: [2]});
    console.log('\n== 补页：随后只接收第 2 页 ==');
    console.log(`接收文件: ${path}（第 ${supplement.record.revision} 次记录；旧页保留）`);
    for (const reference of referencesFor(supplement.collection, {page: 2})) {
      const chart = supplement.collection.entities.charts[reference.entity];
      const verdict = chart ? (chart.confirmed ? '已核实' : '未确认') : reference.status;
      console.log(`第 2 页 ${reference.frame.name} → ${reference.chartPart}，与第 1 页共用图表实体 ${reference.entity}（${verdict}）`);
    }
    const entityId = Object.keys(supplement.collection.entities.charts)[0];
    console.log(`图表实体 ${entityId} 被引用于第 ${chartDetail(supplement.collection, entityId).references.map(reference => reference.page).join('、')} 页`);
    console.log(`图表实体 ${Object.keys(supplement.collection.entities.charts).length} 个；工作簿实体 ${Object.keys(supplement.collection.entities.workbooks).length} 个；记录 ${(await store.read()).records.length} 条（每次接收一条，不覆盖）`);
    console.log(`外部链接: ${supplement.collection.externalLinks.map(link => `${link.target}（followed=${link.followed}）`).join('；') || '无'}`);
    const confirmed = Object.values(supplement.collection.entities.charts).filter(chart => chart.confirmed);
    const unconfirmed = Object.values(supplement.collection.entities.charts).filter(chart => !chart.confirmed);
    console.log(`\n可确认的图: ${confirmed.length ? confirmed.map(chart => chart.part).join('，') : '无'}`);
    console.log(`不能确认的图: ${unconfirmed.length ? unconfirmed.map(chart => `${chart.part}（${chart.reasons.map(reason => reason.code).join(',')}）`).join('；') : '无'}`);
    console.log('\n== 汇总 ==');
    console.log(JSON.stringify(supplement.collection.summary, null, 2));
  } finally {
    await source.close();
  }
}

run().catch(error => {
  console.error(`演示失败: ${error.message}`);
  process.exitCode = 1;
});
