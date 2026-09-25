import { test } from "node:test";
import assert from "node:assert/strict";
import { guardOfficeToolSearch } from "../daemon/tool-search-guard.mjs";

test("already-loaded Office tools are called directly while deferred searches remain available", () => {
  const office = guardOfficeToolSearch({
    tool_name: "ToolSearch",
    tool_input: { query: "select:mcp__office__excel_fill_formula,mcp__office__excel_set_format" },
  });
  assert.equal(office.hookSpecificOutput.permissionDecision, "deny");
  assert.match(office.hookSpecificOutput.permissionDecisionReason, /already loaded.*directly/);
  assert.deepEqual(
    guardOfficeToolSearch({
      tool_name: "ToolSearch",
      tool_input: { query: "select:mcp__thepexcel-excel__excel_format" },
    }),
    {},
  );
});
