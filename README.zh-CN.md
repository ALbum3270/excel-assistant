# Excel Assistant

[English](README.md) | 简体中文

一个运行在 Excel 侧边栏里的 AI 助手。它能读取和修改你当前打开的工作簿，在本机运行。模型可以自选：你登录的 Claude Code，或任何兼容 Anthropic 接口的服务，比如 Qwen、DeepSeek、Kimi、GLM、MiniMax。

这个项目不是从零写起的，而是把多个开源项目组合在一起，每一块都取自已经把这件事做得最好的项目。本仓库自己写的部分包括：衔接它们的胶水代码、过程中发现的上游 bug 的修复，以及一套用来检验效果的评测。

## 能做什么

- **读写工作簿**：通过 Office.js 工具操作值、公式、格式、排序、筛选、表格、图表、行列和工作表。读取和搜索大表时分页进行，不会一次全部载入。
- **动手前先看清工作簿**：每条消息都会自动附上工作簿概览（工作表、表头、表格、命名区域）、你的选区及上下几行，以及上一轮之后你改过的单元格。
- **解释公式、追踪依赖**：能看出一个单元格由哪些单元格算出，也能看出改动它会影响哪些单元格。
- **撤销自己的修改**：每次修改都会建恢复点，覆盖值、公式、格式、排序、清空、插删行列，以及增删和改名工作表。恢复之后还能重做。
- **写入前征求同意（可选）**：打开后，每次修改都会在聊天里等你选择“批准”“本轮全部批准”或“拒绝”。
- **处理聊天装不下的大数据**：在沙箱 shell 里用 Python（标准库）、awk、jq、sqlite3 计算。数据在工作表和沙箱之间直接传递，不经过模型。
- **使用 Office.js 做不到的 Excel 功能**：通过 Windows COM 操作 Power Query、透视表布局、数据模型和 DAX、条件格式、数据验证等。
- **搭建财务模型**：接入 Anthropic 的财务分析技能，支持 DCF、LBO、三表模型、可比公司分析和模型审计。
- **参考本地文件**：你指定的文件夹里的笔记、需求文档、历史资料，它都可以读取。
- **过程可追溯**：每个工作簿有独立的对话历史、恢复点列表、指令队列，界面上还能看到当前模型提供方和 token 用量。

## 由哪些项目组成

| 部分 | 来源 | 许可证 |
| --- | --- | --- |
| 本地 daemon、WebSocket 桥、任务窗格、托盘程序、加载项注册 | [Draftspect](https://github.com/LeonardHope/Draftspect-Add-Ins-for-Word-and-Excel-Powered-by-Claude-Code) | MIT |
| Excel Office.js API 层（读取、写入、搜索、结构、对象） | [office-agents](https://github.com/hewliyang/office-agents) | MIT |
| 工作簿上下文、公式解释与追踪、恢复日志 | [pi-for-excel](https://github.com/tmustier/pi-for-excel) | MIT |
| 系统提示里的解题规程 | [fabric-rlm](https://github.com/pawarbi/fabric-rlm-core)，SpreadsheetBench Verified-400 榜单上唯一的开源方案 | MIT |
| 计算沙箱 | [just-bash](https://github.com/vercel-labs/just-bash) | Apache-2.0 |
| 高级 Excel 功能的 COM 工具 | [ThepExcelMCP](https://github.com/ThepExcel/ThepExcelMCP) | MIT |
| 财务技能 | [financial-services-plugins](https://github.com/anthropics/financial-services-plugins) | Apache-2.0 |
| 智能体循环 | [Claude Agent SDK](https://www.npmjs.com/package/@anthropic-ai/claude-agent-sdk) | Anthropic 条款 |

上游代码由 `scripts/` 下的脚本按锁定的 commit 打包进来，附带的小补丁每处都必须恰好匹配一次。本仓库自己做的部分：

- **安全写入**：覆盖保护；每个写工具都返回执行回执；返回结果有长度上限，大范围填充不会撑爆模型上下文；数据超出目标区域时直接拒绝写入。
- **可靠的工具调用**：支持取消，校验结果归属；报错时写明是哪个 Office.js 接口失败。
- **上游代码的修复**，都是在真实 Excel 里测出来的：
  - 区域内格式不一致时，格式恢复点建不成；
  - 恢复“无填充色”时被 Windows 版 Excel 拒绝；
  - 依赖追踪只保留区域左上角的一格；
  - 撤销删除行列后，依赖它们的公式仍是 `#REF!`。
- **写入前审批**：基于 SDK 的权限钩子，模型每一步拿到的都是真实结果。
- **评测框架**：用于 SpreadsheetBench，见下文。

和其他开源 Excel 助手的完整对比，见 [docs/open-source-comparison.md](docs/open-source-comparison.md)。

## 架构

```
 Excel (Windows)                          Your machine
 ┌─────────────────────────┐   WebSocket   ┌──────────────────────────────┐    HTTPS    ┌──────────────┐
 │ Task pane (Office.js)   │◄─────────────►│ Daemon (Node)                │◄───────────►│ Model API    │
 │ excel_* tools, context, │  127.0.0.1    │ Claude Agent SDK loop        │             │ (Claude, or  │
 │ restore log (IndexedDB) │  :47833       │ in-process MCP tools         │             │ compatible)  │
 └─────────────────────────┘               │ excel_bash (just-bash)       │             └──────────────┘
            ▲                              │ approval, permission guard   │
            │ COM                          └──────────────┬───────────────┘
            │                                             │ stdio MCP
            └──────────────── ThepExcelMCP (optional) ◄───┘
```

托盘程序（`app/`）负责启动 daemon，并把加载项注册到 Excel。任务窗格和 Office.js 由 daemon 在本机 47834 端口提供。

## 环境要求

- Windows 10 或 11，装有桌面版 Microsoft Excel（Microsoft 365 或 Excel 2021 及以上，需支持 ExcelApi 1.9）。
- Node.js 20.18.1 或更高版本。
- 一个可用的模型：已登录的 [Claude Code](https://docs.claude.com/en/docs/claude-code/overview)，或兼容 Anthropic 接口的服务商的 API key。
- 可选：[uv](https://docs.astral.sh/uv/) 加 ThepExcelMCP 源码，用于 COM 工具；Python 加 uv，用于评测。

## 安装

```bash
git clone <本仓库地址>
cd excel-assistant
npm install
npm start
```

启动后托盘会出现图标，首次运行时会提示把加载项装进 Excel，点 **Install**。然后彻底退出并重新打开 Excel，在 **插入 → 加载项**（或 **开始 → 加载项**）里打开 **Excel Assistant**。加载项通过注册表 `WEF\Developer` 键注册，不需要管理员权限，也不需要网络共享目录。

如果想在终端里看 daemon 的日志，用 `npm run dev` 代替 `npm start`。

## 选择模型

不做任何配置时，默认使用你登录的 Claude Code。想换成其他服务商，把 `.env.example` 复制为 `.env` 再填写。这个文件只对本项目生效，不会影响你全局的 Claude Code。以 Qwen 为例：

```bash
ANTHROPIC_BASE_URL=https://dashscope.aliyuncs.com/apps/anthropic
ANTHROPIC_AUTH_TOKEN=sk-你的密钥
ANTHROPIC_DEFAULT_HAIKU_MODEL=qwen3.7-flash
ANTHROPIC_DEFAULT_SONNET_MODEL=qwen3.7-plus
ANTHROPIC_DEFAULT_OPUS_MODEL=qwen3.7-max
```

任务窗格里的模型选择框用来在这三档之间切换。

## 可选：高级工具和技能

把 `agent.config.example.json` 复制为 `agent.config.json`，可以开启：

- **COM 工具**（`mcpServers`）：指向你本地的 ThepExcelMCP 目录；
- **财务技能**（`plugins`、`skills`）：指向 `financial-services-plugins`；
- **内置工具**（`builtinTools`）：默认只有只读的文件访问和网页搜索。

默认不继承你全局 `~/.claude.json` 里的 MCP 服务器（`inheritUserMcpServers: false`），这样上下文更小，行为也更可预期。

## 使用

- 打开一个已保存的工作簿，打开任务窗格，直接描述需求，比如“在销售表下面加一行合计”“为什么 D14 显示 #N/A？”。
- 每次写入都会显示一张卡片：改了哪个区域、是否已提交、如何验证、有几个公式错误，以及恢复入口。
- 助手还在执行时，可以先把下一条指令排进队列。
- **History**（聊天页顶部按钮）：列出这个工作簿的历史对话，可以重新打开、继续或删除。
- **Backups** 页：恢复点列表，可搜索、恢复、删除和清空。也可以直接对它说“撤销你刚才的修改”。
- **Setup** 页：选择工作区文件夹和参考文件，查看当前模型提供方和上一轮的 token 用量，并可打开 **Ask before changing the workbook** 来逐次审批修改。

## 安全与隐私

- 工作簿只通过 Office.js 工具修改。助手不能直接写磁盘上的 Office 文件，VBA 和 Python in Excel 已禁用。
- 覆盖已有数据需要显式标记，每次写入都会报告是否已执行。
- 恢复点只保存在本机任务窗格的存储里。
- 对话内容，包括模型读取的单元格内容，会发送给你配置的模型服务商；助手使用网页搜索或抓取时，请求会发到网上。

## 评测

`evals/run_spreadsheetbench.py` 在真实 Excel 中运行 [SpreadsheetBench](https://github.com/RUCKBReasoning/SpreadsheetBench) Verified-400 的题目，用基准自带的比较代码判分。题目提示使用官方原文，只多一句说明“直接在打开的工作簿里修改”。评测框架有以下特点：
- 锁定每次运行的配置；
- 把基础设施失败和助手自身的失败分开统计；
- 统计答案区以外的误改，并用标准答案做校准。

| 运行 | 模型 | 题数 | 通过 |
| --- | --- | --- | --- |
| 2026-09-19，固定 10 题样本 | qwen3.7-flash | 10 | 4 |
| 2026-09-19，同样 10 题，又一轮改进后 | qwen3.7-flash | 10 | 5 |

这 10 题里有 7 题在两轮之间结果翻转，两个方向都有。小模型加小样本，波动比两轮之间的差值还大，所以这两个数字都不能说明效果有提升。

作为参照，官方榜单上最好的开源方案是 fabric-rlm，用 MiniMax M3 在全部 400 题上得到 82.5%。但它是用 Python 直接改 `.xlsx` 文件，而不是操作打开着的 Excel，所以两边的数字不能直接比较。用小模型跑 10 题只是冒烟测试，不代表正式成绩。

## 开发

```bash
npm test                      # Node 测试
npm run vendor:office-agents  # 重建 office-agents 打包
npm run vendor:pi-context     # 重建 pi-for-excel 打包
python -X utf8 evals/run_spreadsheetbench.py --dataset <spreadsheetbench_verified_400 路径> --run <名称> --model haiku --ids <题号>
```

- 重建打包需要上游源码停在锁定的 commit，且工作区干净。
- 修改任务窗格代码后，需要在 Excel 里重新加载窗格。
- 设计记录和进度见 [docs/optimization-analysis.md](docs/optimization-analysis.md)。

## 已知限制

- 只在 Windows 上开发和测试过，COM 工具也只支持 Windows。
- 以下操作还没有恢复点：表格、图表、透视表、批注、复制工作表、隐藏行列、冻结窗格，以及通过 COM 工具做的所有修改。
- 删除行、列或工作表后再恢复，其他工作表里引用它们的公式仍是 `#REF!`。

## 许可证

MIT，见 [LICENSE](LICENSE)。第三方组件保留各自的许可证，见 [NOTICE.md](NOTICE.md)。
