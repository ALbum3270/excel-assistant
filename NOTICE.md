# Third-party notices

Excel Assistant is assembled from the open-source projects below. Each keeps its own license; copyright notices are preserved here and in the referenced files.

| Component | Used for | License |
|---|---|---|
| [LeonardHope/Draftspect](https://github.com/LeonardHope/Draftspect-Add-Ins-for-Word-and-Excel-Powered-by-Claude-Code) | Base of this repository: local daemon, WebSocket bridge, task pane, tray app, sideloading. See `LICENSE` (Copyright (c) 2026 Leonard Hope). | MIT |
| [hewliyang/office-agents](https://github.com/hewliyang/office-agents) | Excel Office.js API layer, vendored into `taskpane/shared/vendor/office-agents-excel-api.js` by `scripts/vendor-office-agents.mjs`; tool schemas in `daemon/office-tools.mjs`; `sheet-to-csv`/`csv-to-sheet` in `daemon/compute-tool.mjs`. License text: `taskpane/shared/vendor/office-agents-excel-api.LICENSE`. | MIT |
| [tmustier/pi-for-excel](https://github.com/tmustier/pi-for-excel) | Workbook overview, selection context, change tracker, workbook mutation coordinator, formula explain/trace tools and recovery log, vendored into `taskpane/shared/vendor/pi-context.js` and `pi-recovery.js` by `scripts/vendor-pi-context.mjs`; plain-value cell convention in `daemon/office-tools.mjs`. License text: `taskpane/shared/vendor/pi-context.LICENSE` (Copyright (c) 2026 Thomas Mustier). | MIT |
| [TypeBox](https://github.com/sinclairzx81/typebox) | Runtime validation inside the vendored Pi recovery log. Bundled into `taskpane/shared/vendor/pi-recovery.js`. License text: `taskpane/shared/vendor/typebox.LICENSE` (Copyright (c) 2017-2026 Haydn Paterson). | MIT |
| [pawarbi/fabric-rlm-core](https://github.com/pawarbi/fabric-rlm-core) | Task protocol (inspect, restate target/transformation/output type, build and size-check the full result, anti-patterns, boundary spot-checks) adapted from its `excel_modify` skill into `daemon/system-prompt-excel.md`. Copyright (c) 2026 fabric-rlm contributors. | MIT |
| [vercel-labs/just-bash](https://github.com/vercel-labs/just-bash) | Sandboxed shell and Python for `excel_bash`. npm dependency, not redistributed. | Apache-2.0 |
| [PapaParse](https://github.com/mholt/PapaParse) | CSV parsing in `csv-to-sheet`. npm dependency, not redistributed. | MIT |
| [RUCKBReasoning/SpreadsheetBench](https://github.com/RUCKBReasoning/SpreadsheetBench) | Evaluation only: its grader is imported at run time and its task prompt wording is adapted in `evals/run_spreadsheetbench.py`. Dataset not redistributed. | CC BY-SA 4.0 |
| [ThepExcel/ThepExcelMCP](https://github.com/ThepExcel/ThepExcelMCP) | COM-based Excel MCP server (Power Query, PivotTables, Data Model, snapshots), configured in `agent.config.json`. Not redistributed. | MIT |
| [anthropics/financial-services-plugins](https://github.com/anthropics/financial-services-plugins) | Finance skills (audit, DCF, 3-statement, LBO, comps), loaded as a plugin via `agent.config.json`. Not redistributed. | Apache-2.0 |
| [@anthropic-ai/claude-agent-sdk](https://www.npmjs.com/package/@anthropic-ai/claude-agent-sdk) | Agent harness. | Anthropic Commercial Terms |
| [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) | Office.js, served locally by the daemon. | Microsoft license (see package) |
| [marked](https://github.com/markedjs/marked) | Markdown rendering in the task pane. | MIT |
| [DOMPurify](https://github.com/cure53/DOMPurify) | Sanitizing rendered model output. | MPL-2.0 OR Apache-2.0 |
