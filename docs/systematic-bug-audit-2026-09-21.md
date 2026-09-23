# 系统排错记录（2026-09-21）

审查基线：`335364b7dc461192ff35e4983afe0d3bd5b0ae84`。本轮只排查，不修改业务代码。

**结论：当前存在 14 处已复现的代码问题，不能把现阶段描述成“只剩跑评测”。最优先的是写入范围、数据类型、部分提交状态和恢复点真实性。** 另外，任务语义验证仍主要依靠提示词；这与工具执行成功是两件事。

## 方法和证据范围

- 对照 40 题评测的 manifest、summary、逐题 attempt 和 transcript，区分旧版本问题与当前代码问题。
- 审阅 daemon 会话/权限/协调/COM、Office.js 工具及包装层、恢复持久化、任务窗格状态与历史记录、计算沙箱、Electron 生命周期、评测归因代码。
- 建立 [独立复现脚本](../scripts/diagnostics/audit-2026-09-21.mjs)，运行实际模块或从实际源码提取的函数。Office、COM、IndexedDB 和进程边界用隔离桩替代；文件备份问题使用临时目录里的真实文件验证。
- [复现输出](systematic-bug-audit-2026-09-21.repro.json)：14/14 触发当前错误行为。这里的“复现成功”不是“修复通过”。本轮基线的既有 148 项测试通过，说明既有用例没有覆盖这些组合路径。
- 未重新执行 40 题，也未对真实用户工作簿做破坏性验证。涉及 Office 的发现，其代码路径已复现，但真实宿主表现仍应在修复后的少量定向验收中确认。

复现命令（约数秒，无模型调用）：

```powershell
node scripts/diagnostics/audit-2026-09-21.mjs
# 同时更新本报告配套的 JSON 证据：
node scripts/diagnostics/audit-2026-09-21.mjs --save
```

以下 P0 表示应在下一轮大规模评测前修复的数据正确性/恢复安全问题；P1 为功能、执行稳定性及权限链路问题；P2 为较低频的边界兼容问题。不是将 14 处发现都归因于那 14 道错题。

## 修复状态（2026-09-22）

后续复查发现 B10/B11/B12 的三处遗漏，现已补修，详见 [复查与补修记录](bug-fix-recheck-2026-09-22.md)。当前 Node 测试为 169/169；仍未重跑真实 Excel 40 题。以下表格记录的是首轮修复，不代表所有边界均已验收。

14 项已全部修复，每项都有回归测试。复现脚本 `scripts/diagnostics/audit-2026-09-21.mjs` 保留原样，它复现的是 `335364b` 上**修复前**的行为，是这些问题存在过的证据；它从源码切片执行，部分切片锚点已随修复改名，不再用作现行代码的检验。现行检验以下表的测试为准。标"旧代码失败"的测试已确认在修复前的代码上失败；其余测试依赖修复时新增的代码结构，无法在旧代码上运行，旧行为的证据来自上面的复现脚本。

| 编号 | 提交    | 回归测试                                                                             | 旧代码失败                 |
| ---- | ------- | ------------------------------------------------------------------------------------ | -------------------------- |
| B01  | d794f22 | `tests/pi-context.test.mjs`：a write that fails part-way still consumes its revision | 是                         |
| B02  | 845afec | `tests/com-backup.test.mjs`：same name in different folders / same millisecond       | 是                         |
| B03  | 845afec | `tests/compute-tool.test.mjs`：keeps lossy numbers as text…                          | 是                         |
| B04  | 47a3e1a | `tests/compute-tool.test.mjs`：a write from a saved script still asks…               | 是                         |
| B05  | 3b68c10 | `tests/archive.test.mjs`：exports and imports with each result paired                | 否（新结构）               |
| B06  | 3b68c10 | `tests/archive.test.mjs`：long Chinese conversation fits the import limit            | 否（新结构）               |
| B07  | 845afec | `tests/taskpane.test.mjs`：a backup that cannot be saved does not swallow…           | 是                         |
| B08  | 6be6bcd | `tests/office-tools.test.mjs`：ignores transient failures / starts over each turn    | 是                         |
| B09  | d794f22 | 由 `captureRange()` 结构保证：先只加载尺寸、判限后再加载数据                         | 否（未单测）               |
| B10  | bb47d32 | `tests/thepexcel-gateway.test.mjs`：stopped while it waits… never starts             | 是（撤掉修复验证）         |
| B11  | 02198f6 | `tests/recovery-restore.test.mjs`：one grouped restore leaves inverses that group    | 是                         |
| B12  | d794f22 | `tests/office-agents-patches.test.mjs`：checked against the range it expands to      | 是                         |
| B13  | 845afec | `tests/taskpane.test.mjs`：失败分支向恢复点传递失败信息；结构恢复收到后不发布逆操作  | 部分（结构函数本身未单测） |
| B14  | c7b1136 | `tests/app-lifecycle.test.mjs`：Stop cancels a pending restart 等 3 条               | 否（新结构）               |

修复中顺带处理：`copyToRange` 的格式快照原先按公式模式本身的尺寸截取，只覆盖填充区的第一块（随 B12 修复）；Excel 把 `3-4` 解析成日期等输入转换已在临时 Excel 实例中核实（随 B03）。

仍需真实 Excel 定向验收：复制扩展、部分提交后的 revision、失败结构操作不再产生恢复点、Stop 与 COM、分组恢复及其再撤销的顺序。验收后再冻结版本、重跑同一 40 题。

## 已确认问题总表

| 编号 | 优先级 | 问题                                    | 可观察后果                                       |
| ---- | ------ | --------------------------------------- | ------------------------------------------------ |
| B01  | P0     | 部分写入失败但已结束时，revision 不增长 | 已改变工作簿仍接受旧 revision 的写入             |
| B02  | P0     | COM 备份文件名冲突并覆盖                | 不同工作簿/同秒不同版本共用一个备份，旧备份丢失  |
| B03  | P0     | CSV 回写强制推断类型，缺少保真路径      | 前导零、长整数、文本布尔值和等号文本被改变       |
| B12  | P0     | 复制自动扩展后的范围未做覆盖检查        | `allow_overwrite=false` 仍覆盖未检查的已有数据   |
| B13  | P0     | 结构操作失败后仍产生可执行的反向恢复点  | 恢复一次失败操作可能删除原有空行、移动其他数据   |
| B04  | P1     | 沙箱审批只扫描最外层命令文本            | 执行已保存脚本可绕过当前写入审批                 |
| B05  | P1     | 回放事件与导出格式不兼容                | 含工具结果的对话导出直接报错                     |
| B07  | P1     | 恢复持久化异常逃出工具错误处理          | Excel 已修改，但没有任何 `tool_result`，最终超时 |
| B08  | P1     | 重复失败预算永久绑定同一会话参数        | 环境已恢复、重新读取后，合法调用仍被拒绝         |
| B09  | P1     | 恢复数据在限制检查之前已全量加载        | 大区域写入仍会在恢复准备阶段发生巨量读取         |
| B10  | P1     | COM 预检查等待期间取消，之后仍发起写入  | Stop 已返回，尚未发出的写操作随后开始            |
| B11  | P1     | 分组恢复产生互不相关的反向恢复点        | 一次操作的“撤销恢复”被拆成多个独立操作           |
| B14  | P1     | 自动重启定时器未被手动启停取消          | 停止后再次自启动，或启动多个 daemon 竞争端口     |
| B06  | P2     | 导出按字符限量，导入按字节拒绝          | 系统导出的较长中文对话无法从界面重新导入         |

## P0：数据正确性与恢复安全

### B01：把执行结束误当成没有发生部分提交

位置：[workbook-execution.mjs](../daemon/workbook-execution.mjs)，`run()` 的 catch，约 150 行；[bridge.mjs](../daemon/bridge.mjs)，`tool_result/tool_settled` 处理，约 419–456 行。

桥接收到任务窗格错误后统一附加 `executionSettled:true`。协调器只有在 `commitStatus !== "not_committed" && !executionSettled` 时增长 revision。于是，Office 已执行前面的写入、后续同步/格式/读回失败，只要错误正常返回，revision 就不增长。

复现：模拟写入已改变单元格，随后抛出 `{commitStatus:"unknown", executionSettled:true}`。结果 revision 仍为 0、blocked 为 null，另一个持有 `expectedRevision:0` 的调用继续成功派发。

应修正：把“执行是否仍在运行”和“工作簿是否可能改变”分别处理。已结束可以解除等待/占用，却不证明没有修改；`not_committed` 才能支撑不增长版本。修复也要覆盖取消后晚到的 `tool_settled`。

### B02：备份目录只按文件名，文件只精确到秒

位置：[com-backup.mjs](../daemon/com-backup.mjs)，`folderName()`、`stamp()`、`backupWorkbookFile()`。

`a/Budget.xlsx` 和 `b/Budget.xlsx` 使用同一个备份目录；时间戳丢弃毫秒，`copyFile` 允许覆盖。相同秒内第二次备份直接替换第一次。同一个工作簿在一秒内保存两次也一样。`lastBackup` 仍可能返回 `reused`，但对应文件已经是另一个工作簿的内容。

复现使用真实临时文件：两个不同目录的 Budget.xlsx 返回相同备份路径；第一次备份读出的内容变成第二本工作簿；再次备份第一本仍返回 reused。另一次同秒保存也覆盖旧副本。

应修正：目录包含规范化完整路径的稳定摘要；每份备份采用不会覆盖既存文件的唯一名称。按真实工作簿身份淘汰历史。这里不要求改变“COM 仅备份最后落盘状态”的既有边界。

### B03：CSV 回写不可保真

位置：[compute-tool.mjs](../daemon/compute-tool.mjs)，`cellInput()`（约 42 行）及 `csv-to-sheet`；[office-agents-excel-api.js](../taskpane/shared/vendor/office-agents-excel-api.js)，`getRangeAsCsv()`。

导出把原值转为字符串；回写逐个执行 Number/布尔/公式推断。CSV 引号无法保留原始单元格类型，也没有 raw/schema 选项。

实际沙箱命令 `sheet-to-csv ... && csv-to-sheet ... --force` 在数据未加工时即产生：

| 源文本             | 传给写工具的内容        |
| ------------------ | ----------------------- |
| `00123`            | 数字 `123`              |
| `9007199254740993` | 数字 `9007199254740992` |
| `TRUE`             | 布尔 `true`             |
| `=1+1`（原为文本） | 公式 `=1+1`             |

应修正：保留现有便捷推断的同时，提供明确的类型/文本保真方式；工作表导出后原样回写应有可用的保真路径。公式输出和文本输出应由调用者明确区分，不能只靠内容猜测。不要把模型提示词当成类型系统。

### B12：复制范围的检查与实际写入不一致

位置：[office-agents-excel-api.js](../taskpane/shared/vendor/office-agents-excel-api.js)，`copyTo()`（约 790 行）、`setCellRange()` 中 `copyToRange` 的检查（约 578 行）；恢复范围相关逻辑见 [recovery.js](../taskpane/shared/recovery.js) 的 `captureRangeSnapshot()` 和 `recoveryPlan()`。

触发：源为 `A1:B2`，目标只写 `D1`，D1 为空而 E2 有旧数据。代码只检查 D1，随后 `copyFrom` 自动扩展到 D1:E2。因此 `allow_overwrite=false` 也能覆盖 E2。

Microsoft 明确规定目标小于源时会自动扩展；这不是凭空假设的宿主行为。[Range.copyFrom 文档](<https://learn.microsoft.com/en-us/javascript/api/excel/excel.range?view=excel-js-preview#excel-excel-range-copyfrom-member(1)>)。

复现使用实际覆盖检查和 `copyTo()`，以符合该文档的复制桩模拟 Office：E2 从 KEEP 变成 4，工具返回成功。该复现没有声称跑过真实 Excel。

同一问题影响恢复覆盖：`sizeFromAddress` 只在目标不含冒号时扩展；小于源的显式矩形目标也会自动扩展，但恢复捕获不扩展。`set_cell_range.copyToRange` 的值快照也没有统一使用实际最终范围。

应修正：在执行前解析实际影响区域，覆盖检查、恢复、回执和复制共用该区域，避免各自推算。

### B13：失败的结构修改生成“幽灵恢复点”

位置：[taskpane.js](../taskpane/shared/taskpane.js)，`runOfficeTool()` catch（约 2140 行）；[recovery.js](../taskpane/shared/recovery.js)，`prepareStructureRecovery()` / `commitStructureRecovery()`；[pi-recovery.js](../taskpane/shared/vendor/pi-recovery.js)，`applyRowsState()`（约 11658 行）。

插入前记录 `rows_absent`，含义是恢复时删除这次新插入的行。但工具失败后，catch 仍无条件调用 `commitRecovery()`，结构恢复闭包不核实插入是否发生，照样发布此恢复点。

复现：准备在第 2 行插入，未发生插入，却执行失败分支的 commit。系统创建 rows_absent 恢复点；若原第 2 行是空行，Pi 的“有数据不删除”保护允许删除，原下方数据随之上移。若目标行有数据，Pi 会拒绝恢复，所以不是任何情况下都会删除非空数据。失败删除的反向“插入”也存在同类真实性风险。

应修正：结构恢复点必须对应已确认的结构变化。明确未提交时不发布可执行的结构逆操作；未知结果先核对实际结构，不能把请求参数直接当成提交证据。

## P1/P2：执行链路与产品功能

### B04：写入审批可被保存的脚本绕过

位置：[approval.mjs](../daemon/approval.mjs)，`needsApproval()` 的 excel_bash 分支；[compute-tool.mjs](../daemon/compute-tool.mjs)，持久化沙箱及 `csv-to-sheet` 的桥接调用。

审批判断只是对顶层 command 搜索 `csv-to-sheet`。第一轮存下 later.sh，里面包含回写命令；下一轮执行 `bash later.sh`，判断为不需要审批，但脚本实际调用 `excel_set_cell_range`。前一轮批准保存脚本不等于后续每次执行都获得当前轮的写入批准。变量构造命令也有同类问题。

复现中真实 just-bash 执行脚本，审批为 false，却捕获到写调用。应把审批放到沙箱实际执行写命令的边界；不能扩充一串文本正则解决间接执行。

### B05：带工具结果的会话无法导出

位置：[transcript.mjs](../daemon/transcript.mjs)，`eventsFromLine()` 约 110 行；[index.mjs](../daemon/index.mjs)，`portableTranscriptEvents()` 约 681 行。

回放已生成 `kind:"tool_result"`；导出只接受 user/assistant/tool，因此正常工具结果直接触发 `Conversation archive contains an unsupported event`。复现连接两个实际函数即报错。修复需统一事件格式，并保留 tool 与 tool_result 的关联 id；当前 portable tool 也会丢弃 id。

### B06：中文导出文件超出自己的导入限制（P2）

位置：[index.mjs](../daemon/index.mjs)，`exportConversation()` 按 JSON 字符长度累计；[taskpane.js](../taskpane/shared/taskpane.js) 约 573 行按 `file.size` 拒绝超过 2,000,000 字节。

复现：导出 17 条各 10 万汉字的消息，JSON 约 170 万字符、510 万 UTF-8 字节。导出允许，界面导入拒绝。应按最终文件编码后的字节预算裁剪，并统一导入、传输限制与导出限制。

### B07：恢复保存失败使成功写入失去结果

位置：[recovery.js](../taskpane/shared/recovery.js)，`commitCustomRecovery()` 约 611 行；[taskpane.js](../taskpane/shared/taskpane.js)，约 2109、2140 行。

普通 range 恢复会捕获保存异常，但自定义恢复直接 await appendCustomSnapshot。若 IndexedDB quota/事务写入失败，成功执行的工具进入 catch，catch 再调用一次同一 commit，又抛错，后面的错误回执完全不执行。

实际两个源码函数组合复现：工作簿写入 1 次，保存恢复点尝试 2 次，WebSocket 结果消息 0 条。应独立报告“写入结果”和“备份保存结果”，让存储故障不能吞掉写操作已发生的事实，也不应在 catch 盲目二次提交。

### B08：失败预算阻止恢复后的合法调用

位置：[office-tools.mjs](../daemon/office-tools.mjs)，`repeatedFailures` / `guarded()`，约 340–384 行。

计数属于 MCP server 闭包，不按用户轮次或状态变化重置，也不区分失败原因。一次调用因正在编辑单元格而失败三次，即使用户按 Enter、重新读到了新 revision，第四次同参数调用也永远到不了 Excel；只能改参数/重建 server，或等不相关调用填满计数表后清空。

复现使用实际 MCP handler：前 3 次底层失败；模拟环境恢复并读回后，第 4 次仍被拒绝，底层调用数停在 3。应限定预算生效范围，并允许外部状态改变后的正常重试；不要迫使模型通过改写参数规避限制。

### B09：恢复上限没有挡住大范围读取

位置：[recovery.js](../taskpane/shared/recovery.js)，`captureRangeSnapshot()` 约 235–260 行。

先加载维度，随后加载并 sync 全部 values/formulas，最后才判断 20,000-cell 上限。为 A1:A1000000 准备恢复时，代码已请求百万格的数据才报“过大”。正常读工具的分页优化并没有覆盖这条写前路径。

复现记录实际函数 load 顺序确认上述行为，不通过真分配百万格制造卡顿。应在最终区域维度确定后、数据加载前判限；变更大小的路径也要先读取维度再决定捕获策略。

### B10：停止之后才开始新的 COM 写入

位置：[thepexcel-gateway.mjs](../daemon/thepexcel-gateway.mjs)，传入 `execution.run()` 的回调、info 和 backup 后的调用，约 198–250 行。

回调不接收协调器提供的 signal。若用户在异步 info/备份期间 Stop，外层已返回 TOOL_CANCELLED，但 info 完成后代码仍继续 backup 和写工具。

复现实际 gateway 和协调器：让 info 等待，取消，拿到取消结果；释放 info 后出现新的 excel_range 写调用。这与“已经派发的 Office/COM 操作无法立即中断”不同，应在首次派发写操作前检查取消，并在能够取消的上游调用处传递 signal。

### B11：一次恢复的反向操作失去分组

位置：[recovery.js](../taskpane/shared/recovery.js)，`groupSnapshots()` / `workbookHistory(restore)`；[pi-recovery.js](../taskpane/shared/vendor/pi-recovery.js)，`restoreWorkbookRecoverySnapshot()`，约 11990 行起。

原工具的值、格式、结构快照以同一 toolCallId 成组。逐个恢复时，Pi 为每份逆快照生成 `restore:<各自 snapshot.id>`；自定义路径则为 `restore_<id>`。结果一次恢复的多个逆快照互不成组。界面再点击一个恢复点，只会恢复其中一部分。

复现实际 Pi restore 函数：原快照 1 组，恢复后的逆快照 2 组。应让一次分组 restore 生成共同操作标识，并核对结构与值/格式的逆向应用顺序。脚本仅确认分组丢失，不声称所有恢复顺序组合都已实测。

### B14：自动重启与手动启停竞争

位置：[app/main.mjs](../app/main.mjs)，`startDaemon()` exit 回调约 118–135 行；`stopDaemon()` / `restartDaemon()`。

崩溃时设置的一秒自动重启定时器没有保存/取消。用户在这一秒内手动启动再停止，旧定时器仍启动新的 daemon；直接手动启动则可能叠加启动。startDaemon 本身也没有阻止并发启动。

实际函数配进程/时钟桩复现：初始进程、手动新进程、Stop 后旧定时器启动的进程，共 3 次 spawn，最终状态 starting。应让生命周期维护唯一的在途启动和重启定时器；手动停止/重启淘汰旧调度。

## 40 题成绩能说明什么

来源：[manifest](../evals/runs/sample40-flash-e1f2622/manifest.json)、[summary](../evals/runs/sample40-flash-e1f2622/summary.json)。评测版本是 `e1f2622`，并非当前 HEAD。

- 40 道不同题目，26 通过，14 未通过；单元格任务 21/28，工作表任务 5/12。
- 39 次 agent completed，1 次 timeout。基础设施完成率 100% 是重试后的最终汇总；results.jsonl 有 45 条尝试，不能不去重就当成 45 道题。
- 133 次工具错误中包括 23 次覆盖保护拒绝。26 道通过题中，20 道也发生过工具错误。
- `177-6`、`183-8`、`50521`、`57989`：旧任务验证 passed，但独立评测 failed。模型自己选的检查条件不能证明满足完整题意。
- `54085`：记录的唯一工具是 ToolSearch，最终文字却声称已写公式、填充并经 verify_task 验证。这是执行证据与最终陈述不一致的实际样本，不能靠补一个数组边界检查解决。
- `118-50`：最终按前四个字母旋转解题，与题面“第五个字母移到开头”不同；旧验证也 failed，仍被口头解释为已通过主要检查。此题还存在题面要求排序 A 列而 answer_position 仅为 C:D 的范围冲突，不能把所有范围外变化一律算成系统代码缺陷。
- preservation 可核对的 34 题中，1 题有 19 个范围外单元格变化；另 6 题因金标本身修改范围外区域跳过该指标。不能将“26 通过”直接解读成完整的数据保全保证。

[run_spreadsheetbench.py](../evals/run_spreadsheetbench.py) 中 failure_class 的判断只是“失败题出现过非 overwrite_guard 工具错误，就归 tool_or_recovery”。**11 道 tool_or_recovery 不是经因果分析确认的 11 道代码 bug。** 应保留为粗分组，不能据此宣称修几个工具就能把这 11 题全部变成通过。

## 当前结构缺口与审查边界

1. **语义正确性尚未闭环。** `5e524c3` 删除了 excel_verify_task，当前 [system-prompt-excel.md](../daemon/system-prompt-excel.md) 的 Verify before reporting 是提示词规则；[index.mjs](../daemon/index.mjs) 的 result 分支没有校验是否读回、独立重算或覆盖目标范围。提示词可以改善行为，但不是执行门槛。[priority-roadmap.md](priority-roadmap.md) 的“强制核对”不能解释为系统已保证任务正确。也不应简单恢复旧自声明验证器：40 题已证明它会漏检。
2. **多处代码分别解释一次修改。** 审批按工具/字符串分类，覆盖检查按传参范围，Office 按实际 API 扩展，恢复按另一套 plan 捕获，revision 按结果字段判断。B01/B04/B07/B12/B13 都来自这些接口语义不一致。优先统一写入的实际目标、提交结果、恢复结果；不需要重写整个 Agent 框架。
3. **三个历史事件协议没有同步演进。** SDK transcript、窗格 replay、portable archive 各自处理事件类型，造成 B05/B06。复用同一事件表示比继续补三份相似分支更可靠。
4. 手动编辑、其他程序写入、多个 daemon 不属于当前 revision 的保证范围，[协调设计](workbook-coordination-design.md) 已明确说明；本报告不将这个已声明边界冒充新发现的实现 bug。
5. 首次保存/另存为后的工作簿身份更新、真实多窗格和跨 COM 实例、各种 Excel 版本的图表/透视表恢复、真实超长上下文压缩，尚缺真实宿主验证。本轮不对这些未复现路径作“没有问题”的保证。

## 建议修复顺序与验收范围

1. **先修数据与提交语义：B01、B12、B13、B03、B02。** 对应部分提交、真实目标范围、结构恢复真实性、类型保真、唯一备份。
2. **补齐取消、审批和错误返回：B10、B04、B07、B08、B09、B14。** 每项沿现有边界修复，避免引入新的庞大防御层。
3. **修产品闭环：B05、B06、B11。** 会话正常导出导入，分组恢复能 够完整撤回。
4. 对上述复现改为期望正确行为的少量回归；真实 Excel 只做复制扩展、部分提交、结构失败恢复、Stop/COM、配对恢复等关键路径的定向验收。没有必要先重复跑 40 题。
5. 最后冻结版本、重跑同一 40 题，对比独立成绩、数据保全和真实写入证据。语义错误单独归因，不能只报告工具错误数量减少。

本轮产物仅为本报告、隔离复现脚本及输出；上述业务缺陷尚未修复。
