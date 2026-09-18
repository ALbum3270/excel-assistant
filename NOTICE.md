# Third-party notices

Excel Assistant is assembled from the open-source projects below. Each keeps its own license; copyright notices are preserved here and in the referenced files.

| Component | Used for | License |
|---|---|---|
| [LeonardHope/Draftspect](https://github.com/LeonardHope/Draftspect-Add-Ins-for-Word-and-Excel-Powered-by-Claude-Code) | Base of this repository: local daemon, WebSocket bridge, task pane, tray app, sideloading. See `LICENSE` (Copyright (c) 2026 Leonard Hope). | MIT |
| [hewliyang/office-agents](https://github.com/hewliyang/office-agents) | Excel Office.js API layer, vendored into `taskpane/shared/vendor/office-agents-excel-api.js` by `scripts/vendor-office-agents.mjs`; tool schemas in `daemon/office-tools.mjs`. | MIT |
| [ThepExcel/ThepExcelMCP](https://github.com/ThepExcel/ThepExcelMCP) | COM-based Excel MCP server (Power Query, PivotTables, Data Model, snapshots), configured in `agent.config.json`. Not redistributed. | MIT |
| [anthropics/financial-services-plugins](https://github.com/anthropics/financial-services-plugins) | Finance skills (audit, DCF, 3-statement, LBO, comps), loaded as a plugin via `agent.config.json`. Not redistributed. | Apache-2.0 |
| [@anthropic-ai/claude-agent-sdk](https://www.npmjs.com/package/@anthropic-ai/claude-agent-sdk) | Agent harness. | Anthropic Commercial Terms |
| [@microsoft/office-js](https://www.npmjs.com/package/@microsoft/office-js) | Office.js, served locally by the daemon. | Microsoft license (see package) |
| [marked](https://github.com/markedjs/marked) | Markdown rendering in the task pane. | MIT |
| [DOMPurify](https://github.com/cure53/DOMPurify) | Sanitizing rendered model output. | MPL-2.0 OR Apache-2.0 |
