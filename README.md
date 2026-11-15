# 路演文件中图表与嵌入台账的联合接收

供应商只交一份演示文稿，采购助理会把里面的产能曲线抄进台账。这里保存现有资料预览工具和本地供方模拟服务，便于离线开发和复查领取行为。

## 当前能力

`src/collect.mjs` 沿“页面 → 图形 → 图表 → 嵌入工作簿”的关系采集图表证据，`src/receive.mjs` 将结果写入带校验和的接收文件。每个数据点同时保留显示缓存指针与嵌入工作簿单元格来源；缓存与工作簿不一致、缺格、隐藏列、长度不同、外部链接和多页共用同一张图都会在结果中标出。详见 `docs/collection.md`。

`npm run demo` 从实际本地资料源读取真实文稿，先做原有原始预览，再只接收第 1 页、随后补收第 2 页，打印每轮实际结果、可确认与不可确认的图及原因。`src/domain.mjs` 保留原有领域预览函数；`src/store.mjs` 提供带完整性校验的单进程顺序接收文件，不是跨进程事务数据库。

## 运行

固定 Node.js 22.22.1，不依赖外部服务。所有 HTTP 资料源由进程在回环随机端口启动并关闭。

```sh
npm ci
npm run build
npm test
npm run demo
npm run receive -- --file fixtures/capacity.pptx --pages 1
```

推荐使用固定镜像构建，首次构建会安装锁文件中的依赖；运行时不需联网。

```sh
docker build -t plant-presentation-charts .
docker run --rm --network none --ipc=host plant-presentation-charts npm test
docker run --rm --network none --ipc=host plant-presentation-charts npm run demo
```

开发时 `npm start` 打印本地源地址；设置 `PORT` 可指定监听端口，未设置时自动分配回环端口。浏览器页面或 HTTP 接口可直接查看。`CAPTURE_PATH` 可改变接收文件路径。重复运行预览追加记录；删除 `var` 可从空接收状态开始。

## 模块

- `src/source.mjs`：该业务的本地资料源与演示控制，独立于接收行为。
- `src/domain.mjs`：原有资料预览和领域数据结构（保留原有调用）。
- `src/zip.mjs`：惰性 ZIP 读取，成员级解压与 CRC 校验；一个成员损坏不影响其他成员。
- `src/xml.mjs`：不执行脚本、拒绝 DOCTYPE 的小型 XML 解析器。
- `src/opc.mjs`：OPC 关系解析、包内路径解析与内容类型。
- `src/spreadsheet.mjs`：嵌入 XLSX 的工作表、共享字符串、隐藏行列与单元格读取。
- `src/chart.mjs`：图表系列解析，显示缓存与工作簿单元格的逐点对照。
- `src/collect.mjs`：按页采集、共享实体、外部链接登记、补页合并与汇总。
- `src/receive.mjs`：接收入口（文件/本地源、页码选择、JSON 或摘要输出）。
- `src/preview.mjs`：`npm run demo` 的一次运行入口。
- `fixtures/source.json` 与 `fixtures/controls.json`：本地供方资料及可切换情形。
- `docs/source.md`、`docs/collection.md`：源端约定与采集记录说明。
- `test`：预览、源端、接收文件完整性，以及图表采集、补页与损坏隔离测试。

## 业务核对

以下情况都能在接收结果中独立观察：

- 多页引用同一图资源：每页各有一条引用，指向同一个按内容摘要标识的图表/工作簿实体。
- 缓存与嵌入工作簿不一致：每个数据点分别给出缓存值、单元格值及各自来源，状态为 `conflict`。
- 类别缺格及系列长度不同：`gaps`、`lengths` 与类别/数值分开的 `counts` 如实记录，不补零。
- 外部链接不实际访问：`externalLinks` 登记目标且 `followed: false`，测试用本地 HTTP 服务确认请求数为 0。
- 选择补页及单图损坏隔离：`--pages` 选择页码，补页合并保留旧页；损坏的单图只影响自身实体。

测试既覆盖源端行为，也覆盖接收器对真实封装结构（真实 ZIP、关系、图表缓存与嵌入 XLSX）的处理。

## 短时操作演示

从空 `var` 运行 `npm run demo`，展示真实读取、按页接收、补页合并、共享实体、每个数据点来源以及不能确认的图和原因；再运行 `npm test` 查看全部检查。固定环境中的这两步应在 90 秒内完成。演示每次重新采集，不回放旧日志，也不以源端控制操作代替接收结果。
