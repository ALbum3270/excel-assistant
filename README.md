# Excel Assistant

English | [简体中文](README.zh-CN.md)

An AI agent in the Excel side panel. It reads and edits the workbook you have open, runs locally on your machine, and works with the model of your choice: your Claude Code login, or any Anthropic-compatible API such as Qwen, DeepSeek, Kimi, GLM or MiniMax.

It is assembled from open-source projects rather than written from scratch. Each part comes from the project that already does it best. This repository adds the glue between them, fixes for the upstream bugs found along the way, and an evaluation harness to check the result.

## What it can do

- **Read and edit the workbook** through Office.js tools: values, formulas, formatting, sorting, filters, tables, charts, rows, columns and sheets. Reads and searches page through large sheets instead of loading them whole.
- **See the workbook before it acts.** Every message carries the workbook overview (sheets, headers, tables, named ranges), your selection with the rows around it, and the cells you changed since the last turn.
- **Explain formulas and trace dependencies.** It can show what feeds a cell and what breaks if you change it.
- **Undo many of its edits.** Restore points cover values, formulas, formatting, sorting, clearing, rows and columns, sheets, tables, worksheet filters, hidden rows and columns, frozen panes, and chart creation or selected chart properties. Restoring a change can itself be redone.
- **Ask before it writes (optional).** Turn this on and every change waits for Approve, Approve rest of turn or Reject in the chat.
- **Crunch data too large for the chat** in a sandboxed shell with Python (standard library), awk, jq and sqlite3. Data moves between the sheet and the shell without passing through the model.
- **Use Excel features Office.js cannot reach** through Windows COM, such as Power Query, PivotTable layouts, the Data Model and DAX, conditional formatting and data validation.
- **Build finance models** with Anthropic's financial-analysis skills: DCF, LBO, three-statement, comps and model audits.
- **Read your local files** (notes, specs, prior work) from folders you point it at.
- **Keep track of the work**: one conversation history per workbook with portable read-only exports, a restore-point list, queued follow-ups, and provider, token and estimated cost displays.

## How it is built

| Part | Taken from | License |
| --- | --- | --- |
| Local daemon, WebSocket bridge, task pane, tray app, sideloading | [Draftspect](https://github.com/LeonardHope/Draftspect-Add-Ins-for-Word-and-Excel-Powered-by-Claude-Code) | MIT |
| Excel Office.js API layer (reads, writes, search, structure, objects) | [office-agents](https://github.com/hewliyang/office-agents) | MIT |
| Workbook context, formula explain and trace, restore log | [pi-for-excel](https://github.com/tmustier/pi-for-excel) | MIT |
| Task-solving protocol in the system prompt | [fabric-rlm](https://github.com/pawarbi/fabric-rlm-core), the only open-source entry on the SpreadsheetBench Verified-400 leaderboard | MIT |
| Sandboxed shell for computation | [just-bash](https://github.com/vercel-labs/just-bash) | Apache-2.0 |
| COM tools for advanced Excel features | [ThepExcelMCP](https://github.com/ThepExcel/ThepExcelMCP) | MIT |
| Finance skills | [financial-services-plugins](https://github.com/anthropics/financial-services-plugins) | Apache-2.0 |
| Agent loop | [Claude Agent SDK](https://www.npmjs.com/package/@anthropic-ai/claude-agent-sdk) | Anthropic terms |

Upstream code is vendored at pinned commits by scripts in `scripts/`, with small patches that must each match exactly once. What this repository adds:

- **Safe writes.** Overwrite protection, a commit receipt from every write tool, and bounded results so large fills don't overflow the model's context. Writes are rejected when the data would spill outside the target range.
- **Reliable tool calls.** Cancellation and ownership checks on tool calls, and error messages that name the Office.js API that failed.
- **Fixes to upstream code**, all found by testing in real Excel:
  - format restore points failed on ranges with mixed formatting;
  - restoring "no fill" was rejected by Windows Excel;
  - dependency traces kept only the top-left cell of a range;
  - undoing a row or column delete left dependent formulas as `#REF!`.
- **Approve before apply**, built on the SDK's permission hook so the model still sees each real result.
- **An evaluation harness** for SpreadsheetBench (see below).

The full comparison with other open-source Excel agents is in [docs/open-source-comparison.md](docs/open-source-comparison.md).

## Architecture

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

The tray app (`app/`) starts the daemon and registers the add-in with Excel. The daemon serves the task pane and Office.js locally on port 47834.

## Requirements

- Windows 10 or 11 with desktop Microsoft Excel (Microsoft 365, or Excel 2021 or later; ExcelApi 1.9 or later).
- Node.js 20.18.1 or later.
- A model: a signed-in [Claude Code](https://docs.claude.com/en/docs/claude-code/overview), or an API key for an Anthropic-compatible provider.
- Optional: [uv](https://docs.astral.sh/uv/) and a checkout of ThepExcelMCP for the COM tools; Python and uv for the evaluation harness.

## Install

```bash
git clone <this repository>
cd excel-assistant
npm install
npm start
```

A tray icon appears. On first launch it offers to install the add-in into Excel; click **Install**. Quit and reopen Excel, then open **Excel Assistant** from **Insert → Add-ins** (or **Home → Add-ins**). The add-in is registered through the `WEF\Developer` registry key, so no admin rights or network share are needed.

To watch the daemon's log in a terminal, run `npm run dev` instead of `npm start`.

## Choose a model

Without configuration the agent uses your Claude Code login. You can enter an Anthropic-compatible base URL, credential and tier model IDs in the pane's **Setup → Model connection** section. Saving restarts the local daemon. Alternatively, copy `.env.example` to `.env` and fill it in. This file is read only by this project and does not change your global Claude Code. Example for Qwen:

```bash
ANTHROPIC_BASE_URL=https://dashscope.aliyuncs.com/apps/anthropic
ANTHROPIC_AUTH_TOKEN=sk-your-key
ANTHROPIC_DEFAULT_HAIKU_MODEL=qwen3.7-flash
ANTHROPIC_DEFAULT_SONNET_MODEL=qwen3.7-plus
ANTHROPIC_DEFAULT_OPUS_MODEL=qwen3.7-max
```

The model picker in the pane switches between these three tiers.

## Optional: advanced tools and skills

Copy `agent.config.example.json` to `agent.config.json` to turn on:

- **COM tools** (`mcpServers`): point it at your ThepExcelMCP checkout.
- **Finance skills** (`plugins`, `skills`): point it at `financial-services-plugins`.
- **Built-in tools** (`builtinTools`): the default is read-only file access plus web search.

By default the agent does not inherit the MCP servers from your global `~/.claude.json` (`inheritUserMcpServers: false`), which keeps its context small and predictable.

## Using it

- Open a saved workbook, open the pane and describe what you want, for example "Add a total row under the sales table" or "Why is D14 showing #N/A?".
- Each write shows a card with its range, commit status, verification level, formula errors and restore link. Small cell edits also show before/after values.
- You can queue a follow-up message while the agent is still working.
- **History** (in the chat header) lists this workbook's past conversations; you can reopen, continue, export or delete them. Imported conversation archives are view-only and cannot be resumed by the agent.
- **Backups** tab lists the restore points: search them, restore one, or clear them. You can also just ask "undo your last change".
- **Setup** tab: choose the workspace folder and context files, configure the model provider, see the last turn and current agent run's token usage and estimated cost, and turn on **Ask before changing the workbook** to approve each edit.

## Safety and privacy

- The workbook is changed only through Office.js tools. The agent is blocked from writing Office files on disk, and VBA and Python in Excel are disabled.
- Overwriting cells that already hold data needs an explicit flag, and every write reports whether it was committed.
- Restore points are kept locally in the task pane's storage.
- The conversation, including the cell contents the model reads, goes to the model provider you configure. Web search and fetch, when the agent uses them, go to the web.

## Evaluation

`evals/run_spreadsheetbench.py` runs [SpreadsheetBench](https://github.com/RUCKBReasoning/SpreadsheetBench) Verified-400 tasks in live Excel and grades them with the benchmark's own comparison code. The task prompt is the official one, plus one line adapting it to a live workbook. The harness pins the configuration of each run, separates infrastructure failures from agent failures, and counts edits outside the answer range, calibrated against the reference solution.

| Run | Model | Tasks | Passed |
| --- | --- | --- | --- |
| 2026-09-19, fixed 10-task sample | qwen3.7-flash | 10 | 4 |
| 2026-09-19, same 10 tasks after the next round of work | qwen3.7-flash | 10 | 5 |

Seven of those ten tasks changed verdict between the two runs, in both directions. With a small model on a small sample, run-to-run variance is larger than the difference between the runs, so neither number shows an improvement.

For scale, the best open-source entry on the official leaderboard is fabric-rlm with MiniMax M3 at 82.5% on all 400 tasks. It edits `.xlsx` files with Python rather than driving live Excel, so the numbers are not directly comparable. A 10-task sample on a small model is a smoke test, not a score.

## Development

```bash
npm test                      # Node tests
npm run vendor:office-agents  # rebuild the office-agents bundle
npm run vendor:pi-context     # rebuild the pi-for-excel bundles
python -X utf8 evals/run_spreadsheetbench.py --dataset <path/to/spreadsheetbench_verified_400> --run <name> --model haiku --ids <task ids>
```

- The vendor scripts need the upstream checkouts at their pinned commits with clean working trees.
- After changing task-pane code, reload the pane in Excel.
- Design notes and progress are logged in [docs/optimization-analysis.md](docs/optimization-analysis.md).

## Limitations

- Built and tested on Windows only; the COM tools are Windows-only.
- Restore remains incomplete for chart deletion and data-source changes, PivotTables, comments, duplicated sheets, and writes through the COM tools. New table, filter, hidden-row/column and freeze-pane recovery paths still need live Excel acceptance.
- After rows, columns or a sheet are deleted and restored, formulas on other sheets that pointed at them stay `#REF!`.

## License

MIT; see [LICENSE](LICENSE). Third-party components keep their own licenses; see [NOTICE.md](NOTICE.md).
