import { test } from "node:test";
import assert from "node:assert/strict";
import { createThepExcelGateway } from "../daemon/thepexcel-gateway.mjs";
import { createWorkbookExecution } from "../daemon/workbook-execution.mjs";

test("a COM write stopped while it waits on the workbook check never starts", async () => {
  // With backups on, the missing test file would fail the backup and stop the
  // write for an unrelated reason; turn them off so only cancellation can.
  process.env.EXCEL_COM_BACKUP = "off";
  const calls = [];
  let releaseInfo;
  const infoGate = new Promise((resolve) => (releaseInfo = resolve));
  const client = {
    async listTools() {
      return {
        tools: [
          {
            name: "excel_range",
            description: "range",
            inputSchema: {
              type: "object",
              properties: { action: { type: "string" }, workbook: { type: "string" } },
            },
          },
        ],
      };
    },
    getInstructions: () => "",
    async callTool({ name, arguments: args }) {
      calls.push(`${name}:${args.action}`);
      if (name === "excel_workbook") {
        await infoGate;
        return { content: [{ type: "text", text: JSON.stringify({ path: "C:/Books/A.xlsx" }) }] };
      }
      return { content: [{ type: "text", text: "{}" }] };
    },
  };
  const gateway = await createThepExcelGateway(null, createWorkbookExecution(), { client });
  const stop = new AbortController();
  const server = gateway.createSessionServer({
    workbookPath: "C:/Books/A.xlsx",
    signal: stop.signal,
  });
  const handler = server.instance._registeredTools.excel_range.handler;

  const pending = handler({ action: "write", workbook: "A.xlsx" });
  await new Promise((resolve) => setImmediate(resolve));
  stop.abort();
  const response = await pending;
  assert.match(response.content[0].text, /TOOL_CANCELLED/);

  // The workbook check finishes after Stop; nothing may be written after that.
  releaseInfo();
  await new Promise((resolve) => setTimeout(resolve, 20));
  assert.deepEqual(calls, ["excel_workbook:info"]);
});
