# 图表资料采集

`src/collect.mjs` 沿演示文稿的封装关系采集图表证据：页面 → 图形 → 图表 → 嵌入工作簿。`src/receive.mjs` 把结果写入现有 `CaptureStore`（每次接收一条带校验和的记录）。

## 用法

```sh
# 全部页面，接收本地文稿
npm run receive -- --file fixtures/capacity.pptx

# 只接收第 1 页，输出人读摘要
npm run receive -- --file fixtures/capacity.pptx --pages 1

# 后续补页（同一接收文件内合并页面与实体）
npm run receive -- --file fixtures/capacity.pptx --pages 2

# 查看某一页或某张图在每个数据点上的来源
npm run receive -- --file fixtures/capacity.pptx --page 2
npm run receive -- --file fixtures/capacity.pptx --chart 'chart:ppt/charts/chart1.xml#…'

# 从本地资料源读取，或直接输出 JSON
npm run receive -- --source http://127.0.0.1:PORT --pages 1-2 --json
```

`--capture` 或 `CAPTURE_PATH` 指定接收文件（默认 `var/captures.json`）。页码写 `1`、`1,3` 或 `1-3`。请求的页码在文稿中不存在时，会写入 `selection.missing`；若一页都不存在则报错，不会悄悄按全部页采集。

## 记录结构

每次接收写入 `{kind:"presentation-charts", receivedAt, package, selection, summary, collection, previousRevision}`。`collection` 是合并后的当前视图：

- `package`：文件名、`sha256` 摘要、字节数。
- `pages[]`：页号、所属 `slide` 实体、该页的引用列表、页面状态与来源（含采集时的文稿摘要，便于识别补页后未重采的旧页）。
- `references[]`：每页每个图表图形一条，包含图形名称/序号、关系 Id、图表资料路径、图表实体 Id、状态和未确认原因。共享图表在每一页都有独立引用，但指向同一个实体。
- `entities.slides/charts/workbooks`：以“路径 + 内容摘要”为 Id 的共享实体。重复引用复用同一实体；内容变化即成为新实体。
- `externalLinks[]`：所有 `TargetMode="External"` 的关系（幻灯片、图表、工作簿），记录声明位置与目标，`followed` 恒为 `false`。
- `summary`：页数、图表/引用/工作簿数、已核实与未确认图、`counts`（数值与类别分开）、阻断原因与警告。

## 每个数据点的来源

系列按角色（`name`、`category`、`value`）分别保存：

- `display`：图表 XML 内的显示缓存，`source.part` 为图表资料，`source.pointer` 形如 `/chartSpace/chart/plotArea/barChart[1]/ser[0]/val/numRef/numCache/pt[@idx=0]`。
- `workbook`：嵌入工作簿中的单元格，`source` 至少含 `package`（工作簿实体 Id）、`part`、`sheet`、`cell`。
- `status`：`match`、`conflict`、`displayOnly`、`workbookOnly`、`missing`、`literal`、`unmapped`。

`missing` 表示该位置在显示缓存和工作簿中都没有内容；不会补零，也不会选一个值当真值。`unmapped` 表示引用无法逐点对应（多行多列区域、外部工作簿、工作簿不可用等），此时只记录显示缓存。

引用记录（`references[]`）的状态为 `confirmed`、`unconfirmed`、`corrupt`、`missing`、`external` 或 `unresolved`；与其他页共用实体的引用另有 `shared: true`，但仍带本页自己的引用 Id 与图形位置。

## 冲突、空缺与引用范围

- 每个系列保存声明点数、缓存点数、引用区域大小与工作簿实际单元格数（`lengths`），不一致时给出警告。
- `gaps` 逐点列出缺失的是显示缓存还是工作簿单元格；`conflicts` 同时给出两边的值及来源。
- 引用为单行或单列时按偏移逐点映射；引用覆盖多行多列时不猜测对应关系，整段标记 `ambiguousRange` 并置为 `unmapped`。
- 隐藏列、隐藏行、隐藏工作表会作为警告报告（来源可定位到单元格），但本身不阻止确认；数值冲突、缺格、未映射、字面量数据、外部引用、工作簿不可用以及“只有公式没有缓存值”的单元格会阻止确认。
- 系列名称若引用多个单元格，只按第一个缓存点核对并给出警告。
- 工作表名大小写与引用不同但唯一匹配时按实际工作表读取并记录警告。

## 隔离与安全

- ZIP 中央目录先整体校验；单个成员在读取时才解压和校验 CRC。损坏的图表或工作簿只影响自身实体，其他页和图表继续采集。
- XML 由本项目的解析器读取：不执行脚本、宏或外部 DTD，遇到 `DOCTYPE` 直接拒绝。
- 外部关系只登记，不发起网络请求；`externalData` 指向外部工作簿时状态为 `external`，相关点标记 `unmapped`。
- `activeContent[]` 登记宏（`vbaProject.bin`）、OLE 对象与可执行嵌入件，只记录路径和大小，`read`/`executed` 恒为 `false`。
- 接收文件沿用 `CaptureStore` 的校验和与原子替换；校验失败时报错，不覆盖原文件。

## 确认规则

一张图只有同时满足以下条件才是 `confirmed`：图表资料可读、至少一个系列、所有数值点（以及存在的类别点）都能与工作簿逐点核对为 `match`，且没有冲突、缺格、未映射、字面量数据、外部引用或无法读取的工作簿。否则 `reasons` 给出可定位到资料/指针/单元格的原因。

`stats`/`summary` 中的 `counts` 始终把数值与类别分开统计；`unconfirmed` 列出每张不能确认的图及其原因，便于采购助理知道哪些曲线可以抄、哪些必须回问供应商。
