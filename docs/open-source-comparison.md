# 与公开开源方案的对标（2026-09-19）

目标：公开开源的 Excel 智能体里别人做得更好的部分，我们都借鉴或移植过来，至少达到同一水准。不靠反复刷题修补。以下事实均来自源码或官方数据，仓库都已 clone 到 `../_sdks`。

## 一、榜单（SpreadsheetBench 官方，2026-09-13 更新）

数据来源：`spreadsheetbench/spreadsheetbench.github.io` 仓库的 `data/leaderboard-*.json`，由基准团队独立评测。

- **V1 Verified-400**（与本项目评测同一题集）：共 17 个条目，82.5%–99.25%，几乎全是商业产品：ByteDance、Kingsoft WPS、GPT for Excel、GRID、Shortcut 等。
  - 唯一开源：**fabric-rlm**（MiniMax M3，开源权重），82.5%，仓库 `pawarbi/fabric-rlm-core`（MIT）。
  - 同一数字也写在该项目自己的 `docs/usage-guide.md` 中：单次作答，用官方 `evaluation.py` 判分，每题约 0.0065 美元。
- 参照：DealGlass 公开了结果仓库 `arthursolwayne/spreadsheet-agents`，但引擎 Tetra 闭源。
  - Opus 4.6 仅用三行提示词：80.2%；
  - Opus 4.6 接入 Tetra：95.2%；
  - GPT-5.4 加严格提示：78.2%。
- 高分方案的共同做法：基于文件和 Python（openpyxl/pandas），先算出结果，写入后重新载入核对。
- 本项目：用 qwen3.7-flash 在 10 题样本上得 40%。样本量和模型都不同，不能与上述数字直接比较。

## 二、对比对象（源码已核实）

| 项目 | 形态 | Star / 最近推送 / 许可 | 与我们的关系 |
| --- | --- | --- | --- |
| tmustier/pi-for-excel | Office.js 侧边栏智能体 | 427 / 2026-09-17 / MIT | 最成熟的同类，已移植其上下文模块 |
| hewliyang/office-agents | Office.js 智能体与 Excel API | 578 / 2026-05-12 / README 声明 MIT | Excel API 层已打包移植 |
| pawarbi/fabric-rlm-core | Python 工作区 + openpyxl（文件） | 2 / 2026-09-18 / MIT | 榜单上唯一的开源方案 |
| cmarathe1/MS-Excel-AI-plugin | Office.js + 本地 sidecar | 1 / 2026-06-11 / MIT | 变更集：预览、批准、撤销、自动回滚 |
| getcellm/cellm | 单元格里的 AI 函数（.NET） | 951 / 2026-09-07 / 许可未声明 | 另一类功能：`=PROMPT()` |
| ThepExcel/ThepExcelMCP | COM MCP | 13 / 2026-08-16 / MIT | 已接入 |
| LeonardHope/Draftspect | Claude Code 办公插件 | 27 / 2026-05-22 / MIT | 本项目的基础 |
| CrispStrobe/ExcelLLMAddin | Office.js：二十多个单元格 AI 函数加一个简单智能体 | 0 / 2026-07-16 / AGPL-3.0 | 单元格 AI 函数最完整（批量、流式、缓存、重试、用量统计）；借鉴设计，自己实现，不复制代码 |
| iOfficeAI/OfficeCLI | 无需 Office 的文件命令行工具 | 30.8k / Apache-2.0 | 处理文件，不操作打开着的 Excel，暂不需要 |

## 三、我们已经领先或持平的地方

- 读取和搜索按页进行，能继续翻页；Pi 的 `read_range` 和 `search_workbook` 会一次载入整个区域或整个已用区域。
- 起始上下文：已移植 Pi 的模块，并修掉它的两个问题：选中整列会读入上百万格；把代理自己的写入报成用户修改。
- 桥接层：支持取消，带结果归属检查；超时或断线标记为 `unknown`。所有写工具都有统一的 `commitStatus`。
- COM 高级能力：Power Query、数据模型、数据透视、切片器、数据验证、页面设置等，Pi 和 office-agents 都没有。
- 模型可替换（任何兼容 Anthropic 接口的服务）；可以读本地文件夹作为上下文。
- 评测：有可信评测框架，带运行清单、状态分离、范围外修改统计和标准答案校准。

## 四、差距与移植计划

| 优先级 | 差距 | 谁做得更好（源码位置） | 计划 |
| --- | --- | --- | --- |
| 已完成 | 解题规程：先复述目标区域、变换和输出类型；大范围写入先建完整矩阵，并断言尺寸与目标一致；写后抽查首尾单元格；列出反模式（把宏或说明文字写进单元格、用占位符凑数、差一行） | fabric-rlm `fabric_rlm/skills/excel_modify.md`（MIT），带它上榜的核心配方 | 已改写进系统提示词（87ae843） |
| 已完成 | 撤销与恢复点：值、公式、格式、排序、清空、行列尺寸、插删行列、增删和改名工作表 | Pi `src/workbook/recovery-log.ts` + `recovery/*` | 已移植，并在真实 Excel 中验证恢复（第三十七、三十九、四十一节）。表格、图表、批注和 COM 写入仍没有恢复点 |
| 不做 | 写后发现新出现的错误值就自动回滚 | MS-Excel-AI-plugin `sidecar/src/changeset/manager.ts` | `#N/A` 等错误常是预期结果，自动回滚会撤掉正确写入；改为在回执中报告 `formulaErrors`，由模型判断，需要时用恢复点撤销 |
| 已完成 | 公式解释、依赖追踪 | Pi `explain-formula.ts`、`trace-dependencies.ts` | 已移植并在真实 Excel 中验证；修复了区域引用只追踪左上角的上游问题（优化分析第四十节） |
| 已完成 | 写入前预览和审批（可选，默认关闭） | MS-Excel-AI-plugin 的变更集、ExcelLLMAddin 的先批准再执行 | 通过 SDK `canUseTool` 逐次审批，模型每步拿到真实结果（优化分析第四十二节）；侧边栏已有恢复点列表与恢复操作，仍可通过 `excel_workbook_history` 查看 |
| P2 | 单元格 AI 函数（分类、抽取、拆字段、翻译、`MAP` 批量等） | ExcelLLMAddin `officejs/src/functions/functions.ts`、`core/tasks.ts`；cellm | 需要在 manifest 里加自定义函数；另一类功能，按需求决定 |
| 待定 | 执行原始 Office.js 的万能工具 | Pi `execute_office_js`、office-agents `eval-officejs` | 会绕过覆盖保护和回执；若有恢复点兜底可再考虑 |
| 待定 | 沙箱 Python 缺 pandas/openpyxl | Pi 用 Pyodide（需从 CDN 下载约 15MB）；fabric-rlm 用本机 Python | 目前靠 CSV 加标准库已能完成，先不动 |

## 五、不移植的及原因

许可证只约束原样复制代码，不约束思路和做法。所有来源（包括 AGPL 项目）都可以参考后自行实现；只有原样复制代码时，才受对方许可证约束。


- DealGlass Tetra：闭源。
- OfficeCLI、Bread_Excel_Agent：处理文件，不操作打开着的 Excel，与本产品形态不同。
- fabric-rlm 的 `verified_task`（两次独立求解再对账）：作者注明只适用于只读分析，写入类任务不能重复执行。
