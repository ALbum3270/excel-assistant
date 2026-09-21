import { test } from "node:test";
import assert from "node:assert/strict";
import { ApprovalManager, needsApproval } from "../daemon/approval.mjs";

test("approval classifies Office and ThepExcel reads separately from writes", () => {
  assert.equal(needsApproval("mcp__office__excel_get_cell_ranges", {}), false);
  assert.equal(needsApproval("mcp__office__excel_set_cell_range", {}), true);
  assert.equal(needsApproval("mcp__thepexcel-excel__excel_range", { action: "read" }), false);
  assert.equal(needsApproval("mcp__thepexcel-excel__excel_range", { action: "write" }), true);
  assert.equal(needsApproval("mcp__thepexcel-excel__excel_screenshot", { action: "sheet" }), false);
  assert.equal(needsApproval("mcp__thepexcel-excel__excel_name", { action: "set" }), true);
});

test("a pending approval is replayed after reconnect and settled by the new pane", async () => {
  const sent = [];
  const manager = new ApprovalManager({ sendEvent: (event, key) => sent.push({ event, key }) });
  const controller = new AbortController();
  const pending = manager.request(
    "excel\0book.xlsx",
    { abortController: controller },
    "mcp__office__excel_copy_to",
    { sourceRange: "A1:A2", destinationRange: "B1:B2" },
  );
  const requestId = sent[0].event.request_id;

  manager.replay("excel\0book.xlsx");
  assert.equal(sent[1].event.request_id, requestId);
  assert.equal(manager.respond("excel\0book.xlsx", requestId, "approve"), true);
  assert.equal(await pending, "approve");
  assert.equal(sent.at(-1).event.event, "approval_resolved");

  const pendingAfterToggle = manager.request(
    "excel\0book.xlsx",
    { abortController: controller },
    "mcp__office__excel_set_format",
    { address: "A1", bold: true },
  );
  manager.setEnabled("excel\0book.xlsx", false);
  assert.equal(await pendingAfterToggle, "disabled");
});

test("approval timeout is distinct from user rejection", async () => {
  const manager = new ApprovalManager({ sendEvent() {}, timeoutMs: 1 });
  const decision = await manager.request(
    "excel\0book.xlsx",
    { abortController: new AbortController() },
    "mcp__office__excel_set_format",
    { address: "A1" },
  );
  assert.equal(decision, "timeout");
});

test("excel_bash is not gated by its text; csv-to-sheet asks when it writes", () => {
  // A regex on the command missed saved scripts and built commands (audit B04).
  assert.equal(
    needsApproval("mcp__office__excel_bash", { command: "csv-to-sheet out.csv 1 A1 --force" }),
    false,
  );
});
