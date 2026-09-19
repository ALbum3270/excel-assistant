# Excel Assistant — 项目说明（供 Claude Code / Codex 阅读）

个人使用、同时用作作品展示的 **Excel 侧边栏 AI 助手**。只支持 Windows 本机安装的 Microsoft Excel，模型可替换。产品名是 **Excel Assistant**：不要用 Draftspect（本仓库最初 fork 自它）或 Claude Code 作为产品名。

## 工作方式（先读）

- **组合现成项目，只写胶水。** 动手前先广泛搜索，clone 下来读源码核实。先列出三样东西再实现：“模块 → 现有实现”对照表、放弃的候选及原因、需要自写的胶水清单。对标与移植计划见 `docs/open-source-comparison.md`。
- **目标是达到或超过公开开源方案的水准**，不靠反复跑题、逐题修补。
- **许可证不是排除参考的理由。** 任何项目的做法都可以参考后自行实现。原样复制进仓库的代码必须附上许可文本，并在 `NOTICE.md` 登记。
- **只保留影响数据正确性和可恢复性的必要边界。** 不堆防御性代码和重复测试，不用正则猜测用户意图，不为单题调参。反例是已撤销的按消息文本推断修改范围的限制，见优化分析第三十八节。
- **结论以源码或实测为准。** 费用、用时等估算要明确标注为估算。

## 架构

- **daemon**（`daemon/`）：Claude Agent SDK 会话、WebSocket 桥（`bridge.mjs`）、进程内 MCP 工具（`office-tools.mjs`）。
  - 每个工作簿（宿主加文档身份）有独立会话，见 `sessions.mjs`。
  - 每轮开始前从任务窗格拉取自动上下文（`autoContext`）。
  - 系统提示由 `system-prompt.md` 和 `system-prompt-excel.md` 组成，后者包含从 fabric-rlm 改写的解题规程。
- **任务窗格**（`taskpane/`）：`excel_*` 工具在这里通过 Office.js 执行。`taskpane/shared/vendor/` 下是打包的上游代码，重建时须锁定 commit 且工作区干净，补丁必须恰好匹配一处：
  - `office-agents-excel-api.js`：hewliyang/office-agents，锁定 95fb654。由 `scripts/vendor-office-agents.mjs` 加补丁生成。
  - `pi-context.js`、`pi-recovery.js`：tmustier/pi-for-excel，锁定 fd6c9e3。由 `scripts/vendor-pi-context.mjs` 加补丁生成。
  - `recovery.js` 是 Pi 恢复日志的胶水：IndexedDB 存储、工作簿身份、各写工具的快照计划。
- **计算沙箱**（`daemon/compute-tool.mjs`）：基于 just-bash 的 `excel_bash`，内置 `sheet-to-csv` 和 `csv-to-sheet`。只有 Python 标准库，没有网络，也碰不到本地文件。
- **COM 高级工具**：ThepExcelMCP，经 `agent.config.json` 配置，提供 Power Query、数据透视、数据模型等。
- **托盘**（`app/`）：Electron 程序，负责启动 daemon 并注册加载项。

## 配置

- **模型**：在项目 `.env`（已被 git 忽略）中设置 `ANTHROPIC_BASE_URL`、`ANTHROPIC_AUTH_TOKEN` 和各档模型名，不影响用户全局的 Claude Code。
  - 当前接的是 DashScope 的 Anthropic 兼容接口加 Qwen：haiku 对应 qwen3.7-flash，sonnet 对应 qwen3.7-plus，opus 对应 qwen3.7-max。
- **`agent.config.json`**（已被 git 忽略；模板是 `agent.config.example.json`），可配置项：
  - `mcpServers`、`plugins` / `skills`；
  - `builtinTools`、`settingSources`；
  - `inheritUserMcpServers`：默认 `false`，即不继承 `~/.claude.json` 中的全局 MCP；
  - `env`：例如 `ENABLE_TOOL_SEARCH`。

## 硬性约束

- 工作簿只能通过 `excel_*` 工具修改。`canUseTool` 禁止用 `Write` / `Edit` 写 Office 文件；VBA 和 Python in Excel 已禁用。
- 覆盖已有数据需要 `allow_overwrite`。所有写工具都返回 `commitStatus`；超时或断线时为 `unknown`，必须先重读再重试。
- 写入前自动创建恢复点，可用 `excel_workbook_history` 恢复。以下操作目前没有恢复点：
  - 工作表的增删和改名；
  - 表格、图表、透视表、批注；
  - 所有 COM 写入。
- 评测提示保持 SpreadsheetBench 官方原文，外加一句 Excel 适配说明，不为单题修改。
- 端口：WebSocket 用 47833，HTTP 用 47834。

## 协作

- Codex 也在本仓库工作，分支是 `fix/lifecycle-hangs`。开始前先看 `git log` 和共享进度日志 `docs/optimization-analysis.md`（按节追加）。
- 相关文档：
  - 对标：`docs/open-source-comparison.md`；
  - Pi 对照：`docs/pi-for-excel-comparison.md`；
  - 第三方许可：`NOTICE.md`。
- `README.md` 仍是 Draftspect 原文，待重写。

## 常用命令

```bash
npm run dev                  # 只启动 daemon（调试用）
npm start                    # 启动 Electron 托盘（daemon + 加载项注册）
npm test                     # Node 测试
npm run vendor:office-agents # 重建 office-agents 打包
npm run vendor:pi-context    # 重建 pi-context.js 和 pi-recovery.js
python -X utf8 evals/run_spreadsheetbench.py --dataset <.../spreadsheetbench_verified_400> --run <name> --model haiku --ids <...>
```

- 评测要求 Excel 空闲、电脑不休眠，结果写入 `evals/runs/`（已被 git 忽略）。
- 修改前端代码后，要在 Excel 里重新加载任务窗格才会生效。
