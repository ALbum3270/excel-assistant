# 全项目代码审计：2026-09-22

> 修复状态（2026-09-23）：F01–F09 已全部修复。原隔离复现脚本已改为修复后验证，JavaScript 7/7、Python 2/2 通过；完整 Node 测试 199/199、Python 测试 12/12 和 React 生产构建通过。

审计对象是当前工作区（基线提交 `dd6adc7` 加尚未提交的 Claude/用户改动），不是只检查最近几个修复文件。审计覆盖 daemon 会话与生命周期、WebSocket/HTTP 桥、Office.js 工具、工作簿协调器、恢复日志、计算沙箱、COM 网关与文件备份、Electron 托盘、React 任务窗格、构建及 vendor 补丁、会话持久化和 SpreadsheetBench 评测程序。

第一方 JavaScript/TypeScript/Python 共 83 个文件、约 20,994 行；第三方 vendor 共 32 个文件、约 20,488 行。第三方代码按锁定版本、补丁生成器、导入边界和实际使用路径核对，没有把未修改的上游 UI 组件逐行当作本项目代码重复审计。图片、xlsx、lockfile 和构建产物不按源码逐行审计。

下面的问题描述保留首次审计时的触发条件和证据；当前修复结果以文首状态和文末验证结果为准。此次修复没有调用模型、没有运行 40 题评测，也没有打开或修改真实 Excel 工作簿。

## 先前 6 项的复核

用户列出的 N01–N06 已经由 Claude 修复，并通过原触发路径复核：

| 编号 | 结果 | 复核内容 |
| --- | --- | --- |
| N01 | 已修复 | 工作表名含 `!` 时，D5 仍返回 D5 |
| N02 | 已修复 | 20,001 行、末页为空白单元格时仍保留最后一行 |
| N03 | 已修复 | 独立 `\r` 经 CSV 导出、导入仍留在一个单元格内 |
| N04 | 已修复 | 只改样式的单元格不再被送入内容写入接口 |
| N05 | 已修复 | 无表头建表的撤销与重做恢复相同的单元格布局 |
| N06 | 已修复 | `subtype:"success"` 且 `is_error:true` 会停止队列 |

复核命令执行 29 个相关子测试，29 个通过。全套 Node 测试随后执行 191 个，191 个通过。

## 已修复的问题

以下 9 项都曾有可重复的最小证据，现已按原触发路径修复并转为回归验证。

| 编号 | 优先级 | 问题 | 实际影响 |
| --- | --- | --- | --- |
| F01 | P1 | 公式预检把合法引用中的括号当作公式括号 | 合法公式在到达 Excel 前被拒绝 |
| F02 | P1 | 大 CSV 导入用展开参数求最大列数 | 20 万行导入直接栈溢出，0 行写入 |
| F03 | P1 | 活跃任务期间重载模型/上下文，没有结束前端旧轮次 | 后续指令永久留在队列，界面看似仍在执行 |
| F04 | P1 | 本地工作簿 URL 把 `#` 后内容当 URL fragment 删除 | 两个不同工作簿共享自定义恢复历史，可能恢复错工作簿的数据 |
| F05 | P1 | 恢复快照先按 120 条截断，再按一次操作分组 | 一次恢复只执行半组快照，却返回 `committed` |
| F08 | P1 | 评测保存了被删除的工作表名，但汇总只统计单元格数 | 删除整张非目标工作表仍被统计为 0 个越界修改 |
| F09 | P1 | 评测在进入 `try/finally` 前打开工作簿 | 打开失败后 Excel 的警告保持关闭、可见性也不恢复 |
| F06 | P2 | 任务窗格 CSV 预览用 `split` 解析 CSV | 带逗号或换行的合法字段显示成错误行列；模型输入不受影响 |
| F07 | P2 | 手动“重启”注册的退出回调不受随后“停止”约束 | 用户点重启后马上点停止，daemon 仍会再次启动 |

### F01：合法公式被预检拒绝

位置：`daemon/office-tools.mjs:125`。

`formulaProblem()` 跳过双引号字符串，却仍统计单引号工作表名和结构化引用方括号内的括号。以下合法形状均在调用任务窗格前返回 “parentheses are unbalanced”：

```text
='Plan (draft'!A1
='Plan )'!A1
=Table1[Cost (USD]
```

复现确认 bridge 调用次数为 0。修复应让括号扫描与已有 `topLevelComma()` 一样识别单引号、方括号和大括号，并正确处理 Excel 的重复引号转义。

### F02：大 CSV 导入在写入前栈溢出

位置：`daemon/compute-tool.mjs:145`。

`Math.max(...rows.map(...))` 把每一行都展开成函数参数。隔离沙箱生成 200,000 行单列 CSV 后执行 `csv-to-sheet`，结果是 `Maximum call stack size exceeded`，写入调用次数为 0。200,000 行本身在 Excel 1,048,576 行限制以内。

修复应改成循环或 `reduce` 计算宽度，并在第一次写入前同时验证最终行、列没有超过 Excel 边界，避免更大文件分块写到一半才失败。

### F03：配置重载后下一条消息不再发送

位置：`daemon/index.mjs:2294`、`taskpane/app/core/controller.js:587`。

活跃轮次中切换模型或保存上下文时，daemon 取消旧 session、发 `config_reloaded` 并启动新 session；它没有发旧轮次的 `turn_complete`。前端处理 `config_reloaded` 和随后的 `session_init` 时也没有清除 `turnInFlight`。

复现序列：发送第一条消息 → 模型重载 → 收到新 `session_init` → 发送第二条消息。结果第二条进入本地队列，实际只向 daemon 发出第一条。修复需要为被配置重载取消的轮次发送明确终态，或让前端在重载事件中结束旧轮次；两端只选一个权威路径，避免重复结束。

### F04：文件名含 `#` 的工作簿发生恢复历史碰撞

位置：`taskpane/shared/recovery.js:90`、`taskpane/shared/recovery.js:117`。

恢复身份直接对 `document.url` 执行 `split(/[?#]/, 1)`。Windows 文件名允许 `#`，因此下面两个不同路径得到相同名称 `Budget` 和相同 `workbookId`：

```text
C:\Work\Budget#A.xlsx
C:\Work\Budget#B.xlsx
```

自定义恢复快照直接使用这个哈希作为存储键，没有再绑定文档 token。因此第二本工作簿可能看到并应用第一本工作簿的表格、筛选、隐藏、冻结等自定义快照。

修复时要先区分真正的 `http/file` URL 与原生 Windows 路径。对 URL 使用 URL 解析器；对原生路径保留 `#` 和合法文件名内容。

### F05：恢复历史截断会拆散一次操作

位置：`taskpane/shared/recovery.js:1299`。

`allSnapshots()` 合并两种存储后先 `slice(0, 120)`，之后 `resolveSnapshotGroup()` 才按 `toolCallId` 找同组快照。复现存储中同一次操作仍有“结构”和“值”两条快照，但 119 条较新的自定义快照把较旧的“值”快照挤出 120 条视图。恢复只执行“结构”快照，仍返回 `success:true`、`commitStatus:"committed"`。

修复应先按操作分组，再按组限制列表数量；恢复或删除已选操作时，应从两个底层存储按 `toolCallId` 取完整组，而不是从展示用的截断列表反查。

### F06：CSV 预览不遵守 CSV 引号规则

位置：`taskpane/app/core/csv-preview.js:14`。

`"a,b",c` 被显示为 3 列；含引号内换行的两行数据被显示为 3 行。这里是 UI 预览错误，Office 工具返回给模型的原 CSV 没有被修改。项目已经依赖 PapaParse，可直接复用其解析结果并只截取预览范围。

### F07：停止不能取消已经请求的手动重启

位置：`app/main.mjs:180`。

`restartDaemon()` 给旧进程注册一次性 `exit` 回调，回调无条件执行 `startDaemon()`。用户在旧进程真正退出前再点“停止”，`stopDaemon()` 虽然增加 lifecycle 并清理自动重启定时器，却不能取消这个回调。复现结果是停止后又 spawn 1 个 daemon，最终状态回到 `starting`。

修复可在手动重启时捕获 lifecycle generation，并在退出回调中核对；`stopDaemon()` 后 generation 不匹配就不启动。

### F08：删除整张工作表不计入评测破坏

位置：`evals/workbook_diff.py:65`、`evals/run_spreadsheetbench.py:318`、`evals/run_spreadsheetbench.py:392`。

`unauthorized_edits()` 遇到输出中缺失的原工作表会直接跳过单元格比较，虽然结果另有 `sheets_removed`，`run_task()` 没把它写入扁平记录，`summarize()` 也只判断 `unauthorized_cells`。

最小工作簿删除含数据的 `SourceData` 后得到：

```json
{"unauthorized_cells":0,"sheets_removed":["SourceData"]}
```

汇总将其视为没有越界破坏。这会使“高错误率”的分析遗漏一种严重失败，不一定改变 SpreadsheetBench 官方答案区评分，但会高估项目自己的数据保全指标。

### F09：评测打开失败后不恢复 Excel 全局状态

位置：`evals/run_spreadsheetbench.py:244`、`evals/run_spreadsheetbench.py:247`。

代码先把 `DisplayAlerts=False`、`Visible=True`，再调用 `Workbooks.Open()`，之后才进入 `try/finally`。用抛出 `open failed` 的 COM 桩复现后，这两个全局状态仍为 `False/True`。下一题或用户手工操作可能在警告被关闭的 Excel 实例中继续。

应把状态修改和打开工作簿都纳入外层 `try/finally`，并只在成功取得 workbook 后关闭它。

## 验证结果与边界

- 原 6 项复核：29/29 通过。
- Node 全套：199/199 通过。
- Python 评测单元测试：12/12 通过。
- React 任务窗格生产构建：通过。
- 修复后原触发路径：JavaScript 7/7、Python 2/2 通过。
- 未运行 40 题模型评测；本次结果确认代码级缺陷已关闭，不代表模型任务正确率已经重新测量。
- Office.js 行为相关的旧 6 项仍以隔离 Office 契约桩为主；F01–F09 的修复均在不连接真实 Excel 的边界验证。F04 的触发条件是宿主返回原生 Windows 路径且文件名包含 `#`。

修复后验证证据：

- `scripts/diagnostics/audit-full-2026-09-22.mjs`
- `scripts/diagnostics/audit-full-2026-09-22.py`
- `docs/full-project-audit-2026-09-22.repro.json`
- `docs/full-project-audit-2026-09-22.python-repro.json`

`docs/additional-bug-audit-2026-09-22.md` 保留了更早一轮问题的审计过程；其中 N01–N06 的当前状态以该文档“修复状态”一节和本报告的复核结果为准。
