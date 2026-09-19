# Excel Assistant 与 Pi for Excel 源码对照（2026-09-19）

对照对象：本仓库当前 `fix/lifecycle-hangs` 分支，与同级目录 `pi-for-excel-main`。这份记录只把已在源码中确认的差距列为待办；Pi 的设计并非默认正确，也不以代码量或测试数为目标。

## 已经补强的基础

- 本项目的单元格读取、CSV 读取和全簿搜索在 Office.js 取数前分页，并提供继续游标；Pi 的 `read_range` 和 `search_workbook` 仍会一次性加载目标区域或整张表的值与公式。见本项目 `taskpane/shared/vendor/office-agents-excel-api.js`，以及 Pi 的 `src/tools/read-range.ts`、`src/tools/search-workbook.ts`。
- 本项目已有工作簿隔离的会话身份、串行原子持久化、提交时选区快照、bridge 取消与超时状态、单元格和复制写入的读回与公式错误报告。不要为了对齐 Pi 而重写这些路径。
- 最近的 118-50 真实 Excel 复测：基础设施完成、13 次工具调用、0 次工具错误，但官方判题失败。它说明当前瓶颈已不只是工具参数形状；模型输出了 967 对候选词，未验证转换后的词也在原始词表中。

## 优先补的能力

| 优先级 | 差距与直接影响 | Pi 证据 | 本项目现状 |
| --- | --- | --- | --- |
| P0 | **任务语义验证。** 写入成功和公式无错误不等于任务完成。应在写前形成可检查的预期，在写后读取关键结果、核对约束，并把无法自动判断的项显式报告。 | `docs/research/live-eval-learnings-2026-07.md` §A3 也承认 Pi 的写入验证主要是机械验证。 | `daemon/system-prompt-excel.md` 仅要求模型自觉重读；118-50、183-8 均在无工具错误时失败。 |
| P0 | **统一写入回执与恢复。** 每种写操作都应说明提交状态、实际范围、验证、变更摘要和恢复点；未知或部分提交不得盲目重试。 | `src/tools/write-cells.ts`、`src/tools/mutation/finalize.ts`、`src/workbook/recovery-log.ts`、`src/tools/workbook-history.ts`。 | 仅 `set_cell_range`、`copy_to` 有完整 `commitStatus`；格式、排序、结构等返回形状不统一。`excel_bash` 分块写入可部分提交但不能回滚。 |
| P0 | **执行层编辑范围约束。** 模型给出的 `allow_overwrite` 不是用户授权证明；破坏性工具也需要统一范围和工作簿身份。 | `src/tools/with-workbook-coordinator.ts` 集中区分读写并串行 mutation；Pi 的范围授权仍主要依赖上层策略，不能照搬为完整解法。 | `daemon/office-tools.mjs` 的多个写工具直接透传到 taskpane；权限范围仍是提示词。 |
| 已移植 | **更强的起始上下文。** 自动提供有长度上限的工作簿概览、提交时选中的区域附近的实际值/公式和最近修改，减少模型先猜后读。 | `src/tools/get-workbook-overview.ts`、`src/context/selection.ts`、`src/context/change-tracker.ts`。 | `taskpane/shared/vendor/pi-context.js` 已接入，每轮从面板读取；选区固定为消息携带的地址，单格文本和注入文本均有上限。修改记录可能包含代理写入，不能据此判断编辑者。 |
| P1 | **任务型工具。** 公式填充使用独立的 `range + formula` 入口；范围变换由工具完成“读取 → 计算 → 写回 → 校验”；公式解释和依赖追踪提供取证信息。 | `src/tools/fill-formula.ts`、`src/tools/python-transform-range.ts`、`src/tools/explain-formula.ts`、`src/tools/trace-dependencies.ts`。 | 通用 `excel_set_cell_range` 需要复杂矩阵参数；`excel_bash` 需要模型自己组织 CSV 与 Python 命令；无依赖追踪工具。本轮先加入独立公式填充入口。 |
| P1 | **工作簿协调与长会话上下文预算。** 同一工作簿的写入集中排队，长任务的工具输出和旧结果受控。 | `src/workbook/coordinator.ts`、`src/tools/output-truncation.ts`、`src/compaction/auto-compaction.ts`。 | 当前隔离单位主要是 pane/session；工具各自限制输出，缺少统一的工作簿 revision 和结果压缩策略。 |
| P2 | **产品面扩展。** 多会话标签、恢复 UI、可视化 diff、持久化约定、扩展沙箱、多 provider/WPS 支持。 | `src/taskpane/session-runtime-manager.ts`、`src/ui/tool-renderers.ts`、`src/extensions/sandbox-runtime.ts`。 | 这些提升产品成熟度，但不是当前 SpreadsheetBench 低通过率的首因。 |

## 评测解释力

评测提示保持 SpreadsheetBench 官方原文，不为单题改写（见优化分析第三十二节）；题目要求的答案区外修改由标准答案校准排除。随后让评测保存首个偏差单元格、预期与实际值/公式、每次写入的提交和验证状态。已有 7 题结果是 2/7 通过，不能把这一小样本当全量成功率；其中一部分失败是任务理解或模型推理，不能再归咎于基础设施。

实施顺序：先统一写入回执和未知/部分提交恢复，再做执行层范围约束与语义验证；同时补起始上下文和任务型工具。固定小开发集用于确认每项能力，最后再跑全量评测。不要通过持续堆参数正则或反复跑同一题来替代这些工作。
