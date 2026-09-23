# 追加代码排查：2026-09-22

基线：`dd6adc7`，已包含上一轮 R1–R3 补修和覆盖拒绝的 `not_committed` 标记。本轮继续检查原清单之外的实际调用链。

**结论：新增 6 处可复现的问题，尚未修复业务代码。** 前 5 处影响数据、地址或恢复正确性，第 6 处是失败状态与操作队列的判断不一致。它们不是对上一轮 14 项问题的重复计数，也不能据此断言旧 40 题的错误全部由这些问题造成。

工作区同时存在其他界面开发改动，以及工作簿、依赖文件的改动，本轮均未修改。本报告针对下列源码路径，不将正在开发的 `taskpane/app/` 界面视为已完成验收。

## 修复状态

六项已修复。N01–N05 的复现夹具已转为 [正式回归测试](../tests/additional-audit-regressions.test.mjs)，断言正确行为；N06 的回归在 [taskpane.test.mjs](../tests/taskpane.test.mjs)（“a turn that reports is_error …”）。对照修复前的源码，这 6 条测试全部失败；修复后全部通过。完整 Node 测试 191/191 通过。

- **N04**：同一批写入中只改样式的单元格不再写入内容。数据格按同一行的连续段分段写入，因此实际写入的单元格就是覆盖检查查过的单元格；全部为数据格时仍是一次整体写入。修改位于 `scripts/office-agents-patches.mjs` 的 “preserve sparse cell contents” 补丁，重新生成了 vendor 产物。
- **N05**：无表头建表的恢复点记录 `generatedHeader`。撤销时先转为普通区域，再删除 Excel 自动插入的表头行（上移），数据和下方单元格回到原位；反向恢复点按原方式无表头重建，因此重做、再撤销都成立。回归按官方契约模拟行位移，逐步核对撤销 → 重做 → 撤销，并包含表格下方的已有值。
- **N01**：地址拆分统一改为最后一个 `!` 之后的部分（新增 `rangePart()`），覆盖使用区域、分页起点、写入回执、选区解析；备注地址原本已用最后一个 `!`。
- **N02**：导出按页的行数保留页面，不再按序列化文本是否为空。
- **N03**：CSV 字段含独立 `\r` 时同样加引号。
- **N06**：界面迁移时已统一为 `turnFailed()`，错误展示与是否续跑排队共用同一判断；主动 Stop 仍会续跑下一条。

[复查入口](../scripts/diagnostics/audit-additional-2026-09-22.mjs) 现在运行上述回归测试。范围：N04 的公式求值和 N05 的表格位移仍依据 Office 接口契约的模拟，未在真实 Excel 中验收；下一步应做这两项的定向验收。

## 结果和优先顺序

这里 P1 表示应在下一轮完整评测前修复的数据正确性问题；P2 表示随后修复的交互状态问题。优先级不代表所有输入都会触发。

| 编号 | 优先级       | 问题                                        | 最小触发条件                                  |
| ---- | ------------ | ------------------------------------------- | --------------------------------------------- |
| N04  | P1，优先处理 | 只改样式的单元格被重新写入公式接口          | 同批次其他格写值，原格保存以 `=` 开头的文本   |
| N05  | P1，优先处理 | 无表头建表后撤销残留一行数据                | `has_headers:false`，随后恢复创建表格的恢复点 |
| N01  | P1           | 工作表名中的 `!` 导致读回地址、使用区域错误 | 例如读取 `Sales!2026` 的 D5                   |
| N02  | P1           | CSV 分页丢失末尾单行空白页                  | 单列 20,001 行，最后一行为空                  |
| N03  | P1           | CSV 未转义独立回车符，单元格被拆成多行      | 一个文本单元格包含 `\r`，不包含 `\n`          |
| N06  | P2           | 已显示失败，仍自动执行下一条排队指令        | `subtype:"success"` 且 `is_error:true`        |

## N04：样式修改隐式重写了原有数据

位置：[office-agents-excel-api.js](../taskpane/shared/vendor/office-agents-excel-api.js)，覆盖检查 589–605 行，写入矩阵构造 607–632 行。

原状态：A1 为空，B1 是文本 `=1+1`，例如在常规格式单元格中用前导单引号录入。请求：

```js
setCellRange(1, "A1:B1", [[{ value: 7 }, { cellStyles: { fontWeight: "bold" } }]], {
  allowOverwrite: false,
});
```

预期只写 A1，B1 只加粗。实际覆盖检查确实跳过 B1，但矩阵构造把 B1 原值填回，然后执行整个区域的 `range.formulas = [[7, "=1+1"]]`。

`Range.formulas` 在没有公式时也会返回单元格的值，因此不能仅凭返回字符串以 `=` 开头就认定它是公式；通过该属性写入公式表达式会触发公式语义。[Microsoft Range 文档](https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-preview#excel-excel-range-formulas-member)

**证据：实际 `setCellRange()` 返回成功，隔离的 Office 接口记录到了对 B1 的公式赋值，即使 `allowOverwrite:false`。** 脚本没有模拟公式求值；文本变公式的结论依据实际接口调用与上述 API 契约，尚未在真实 Excel 中验收。

修复方向：不改内容的单元格不要重写内容。使实际写入集合与覆盖检查集合一致；不要为保留原值而将值重新送入公式解析。修改需要进入 vendor 构建补丁，再生成产物。

## N05：无表头建表的恢复范围少了一行

位置：[tools-excel.js](../taskpane/shared/tools-excel.js) 242–250 行；[recovery.js](../taskpane/shared/recovery.js) 589–612、624–633、735–741 行。

`toolExcelCreateTable()` 直接调用 `tables.add(address, false)`。此时 Excel 会自动生成表头并把数据下移一行，这是该参数的正式语义。[Microsoft TableCollection.add 文档](<https://learn.microsoft.com/en-us/javascript/api/excel/excel.tablecollection?view=excel-js-preview#excel-excel-tablecollection-add-member(1)>)

但恢复准备只保存输入地址，撤销结构时只是 `convertToRange()`，随后只写回输入区域的原值。

| 步骤                     | 单列内容      |
| ------------------------ | ------------- |
| 原始 A1:A2，A3 为空      | A、B、空      |
| 对 A1:A2 无表头建表      | Column1、A、B |
| 转回普通区域并恢复 A1:A2 | **A、B、B**   |

新增表头带来的位移没有被撤销，A3 的 B 残留。上一轮修复的组内恢复顺序不能补齐这次漏掉的范围。

**证据：执行实际建表工具、恢复准备和自定义结构恢复函数；Office 桩按官方契约模拟下移，然后按恢复计划已有的值快照写回，得到 A、B、B。** 数据复现不依赖格式快照。本例没有推断区域下方其他已有内容的全部移动规则，也没有冒充真实 Excel 验收。

修复方向：恢复计划必须包含自动表头引起的实际区域变化及对应逆操作。不能仅把表格转换为普通区域；修复后需核对建表、撤销、重做时的区域边界和原数据位置。

## N01：合法工作表名使读取结果标错地址

位置：[office-agents-excel-api.js](../taskpane/shared/vendor/office-agents-excel-api.js) 231、245 行。

`getCellRanges()` 使用 `address.split("!")[1]` 提取区域。`Sales!2026` 是合法工作表名，但 `'Sales!2026'!D5` 被拆开后取到的是 `2026'`，不是 D5。工作表命名的禁用字符不包含 `!`。[Microsoft 工作表命名说明](https://support.microsoft.com/en-us/excel/rename-a-worksheet)

实际函数的隔离复现：

```text
工作表：Sales!2026；使用区域：D5:D6；D5 值：42
请求读取：D5
返回 cells：{ A1: 42 }
返回 dimension：2026'
```

这会给模型错误的位置依据；不指定范围的 `sheet-to-csv` 还会拿到错误的 `dimension`。

同文件 766 行的写入回执、1063 行的选区解析也使用同类拆分，这是源码定位出的同根路径；本轮执行复现的是 `getCellRanges()`。

修复方向：统一使用能正确分离工作表和单元格区域的地址解析，不要混用多个字符串拆分版本。已有 `parseRangeAddress()` 使用最后一个分隔符，应核对并复用适当的公共逻辑。

## N02：CSV 丢失末尾空白页，报告行数却仍正确

位置：[compute-tool.mjs](../daemon/compute-tool.mjs) 98–104 行。

导出按页累加 `rowCount`，但只有 `page.csv` 非空字符串时才把它加入输出。一行一列的空白页正好序列化为 `""`；它有一行数据形状，却被当成没有内容。

使用实际导出 API 与实际计算沙箱命令串联：

```bash
sheet-to-csv 1 A1:A20001 data.csv && csv-to-sheet data.csv 1 B1 --force --text
```

输入前 20,000 行为 `row`，最后一行为空。**导出日志报告 20,001 行，实际导入仅写 20,000 行，最终写入范围停在 B20000。** B20001 不会被写空；如果已有值就会残留。

修复方向：依据页的行数保留空白记录，不能根据序列化文本的真假决定是否保留整页。这与此前 CSV 类型保护是不同问题，`--text` 无法解决。

## N03：独立回车符破坏 CSV 行边界

位置：[office-agents-excel-api.js](../taskpane/shared/vendor/office-agents-excel-api.js) 347 行；导入解析见 [compute-tool.mjs](../daemon/compute-tool.mjs) 的 `csvToSheet`。

CSV 转义仅检查逗号、双引号和 `\n`，漏掉 `\r`。因此一个值为 `alpha\rbeta` 的单元格被导出为未加引号的文本；PapaParse 把独立回车识别为记录分隔符。

**实际导出、导入链路将 A1 的一个单元格写成 B1:B2 两行**，内容分别是 alpha、beta，即使启用 `--text`。这里指独立 `\r`，不是已含 `\n`、会进入现有转义分支的 CRLF。

修复方向：让 CSV 字段转义覆盖回车和换行，保证一个字段中的控制字符不改变记录边界。将修复纳入 vendor 构建补丁。

## N06：错误展示与队列续跑采用不同成功条件

位置：[taskpane.js](../taskpane/shared/taskpane.js) 1574–1582 行；上游转发位于 [daemon/index.mjs](../daemon/index.mjs) 2424–2434 行。

错误提示把 `is_error:true` 视为失败；但决定 `drainQueue` 的条件只检查 `interrupted` 和 `subtype`，忽略了 `is_error`。

输入 `subtype:"success", is_error:true` 时，实际消息处理代码一边显示错误，一边传入 `drainQueue:true`，下一条排队任务会继续执行。

这不是假定 SDK 两个字段必须互相矛盾：当前安装 SDK 的 `SDKResultSuccess` 类型包含 `subtype:'success'` 和独立的 `is_error:boolean`（`sdk.d.ts` 5394、5414 行），daemon 也单独转发两者。复现验证了该消息组合的实际处理结果，没有声称所有失败都会续跑。

修复方向：错误展示和自动续跑共用同一轮次成功判断；保留产品对主动 Stop 后是否运行下一条的既有规则。

## 验证范围与后续工作

- 现有 Node 测试本轮 **169/169 通过**。它们未覆盖上述具体路径，不能用来排除这些问题。
- [最小复现脚本](../scripts/diagnostics/audit-additional-2026-09-22.mjs) 执行当前函数或当前源码切片，仅替换 Office、传输和存储边界；六个问题均复现。[保存的输出](additional-bug-audit-2026-09-22.repro.json)
- 执行：`node scripts/diagnostics/audit-additional-2026-09-22.mjs --save`。该脚本断言的是**缺陷存在**，退出 0 表示复现成功，不表示系统正确。修复时应把对应断言改成正确行为。
- 没有启动真实 Excel、改动真实工作簿或重跑 40 题。N04 的公式求值与 N05 的宿主表格行为在修复后需要少量真实 Excel 定向验收。
- 本轮业务代码未修改。建议先修 N04、N05，再修 N01–N03，最后统一 N06 的结束状态判断；完成后再运行必要回归和固定评测。

结构上，这些问题集中在三个接缝：写入集合与覆盖保护不一致、恢复范围与实际 Excel 操作不一致、分页/CSV 文本与二维数据形状不一致。修复应统一这些契约，避免只给六个示例各加一个特判。
