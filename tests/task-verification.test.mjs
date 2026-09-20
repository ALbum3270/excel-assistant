import { test } from "node:test";
import assert from "node:assert/strict";
import { createTaskVerification } from "../daemon/task-verification.mjs";
import { createOfficeBridgeMcp } from "../daemon/office-tools.mjs";

function workbook(cells, formulas = {}) {
  return async () => ({
    success: true,
    hasMore: false,
    worksheet: { cells: { ...cells }, formulas: { ...formulas } },
  });
}
const target = (range) => ({ sheetId: 1, range });

test("checks find leftover rows, duplicates, blanks and wrong typed values without modifying Excel", async () => {
  const verification = createTaskVerification(
    workbook({ A2: "a", A3: "a", A4: "leftover", B2: "10", B3: 20 }),
  );
  await verification.define([
    { type: "row_count", label: "two results", target: target("A2:A5"), expected: 2 },
    { type: "unique", label: "unique IDs", target: target("A2:A5") },
    { type: "not_blank", label: "required amounts", target: target("B2:B4") },
    { type: "matches", label: "numeric boundary", target: target("B2"), expected: [[10]] },
  ]);
  const result = await verification.run();
  assert.equal(result.status, "failed");
  assert.equal(result.passed, 0);
  assert.equal(result.checks[0].actual, 3);
  assert.deepEqual(result.checks[1].examples, [{ row: 3, duplicateOfRow: 2 }]);
  assert.equal(result.checks[2].examples[0].address, "B4");
  assert.equal(result.checks[3].examples[0].actual, "10");
});

test("source snapshots preserve row tuples and duplicate counts through an in-place edit", async () => {
  const cells = { A1: "east", B1: 1, A2: "west", B2: 2, A3: "east", B3: 1 };
  const verification = createTaskVerification(workbook(cells));
  await verification.define([
    {
      type: "same_rows",
      label: "sort preserves rows",
      target: target("A1:B3"),
      source: target("A1:B3"),
    },
    {
      type: "rows_in_source",
      label: "valid source pairs",
      target: target("A1:B3"),
      source: target("A1:B3"),
    },
  ]);
  cells.A1 = "west";
  cells.B1 = 2;
  cells.A2 = "east";
  cells.B2 = 1;
  verification.markMutation();
  assert.equal((await verification.run()).status, "passed");
  cells.B1 = 1; // Each scalar exists in the source, but this pair never did.
  const result = await verification.finish();
  assert.equal(result.status, "failed");
  assert.equal(result.checks[1].failedCount, 1);
  assert.equal(result.checks[0].failedCount, 2, "one unexpected and one missing source row");
  assert.equal(result.definedBeforeChanges, true);
});

test("verification follows read continuations and never passes an incomplete read", async () => {
  let broken = false;
  const calls = [];
  const verification = createTaskVerification(async (_name, args) => {
    calls.push(args.ranges);
    if (args.ranges[0] === "A1:A5001")
      return {
        success: true,
        hasMore: true,
        remainingRanges: broken ? [] : ["A5001"],
        worksheet: {
          cells: Object.fromEntries(Array.from({ length: 5000 }, (_, i) => [`A${i + 1}`, i])),
        },
      };
    return { success: true, hasMore: false, worksheet: { cells: { A5001: 0 } } };
  });
  await verification.define([{ type: "unique", label: "all IDs", target: target("A1:A5001") }]);
  const result = await verification.run();
  assert.equal(result.status, "failed");
  assert.equal(result.checks[0].examples[0].row, 5001);
  assert.deepEqual(calls, [["A1:A5001"], ["A5001"]]);
  broken = true;
  assert.equal((await verification.run()).status, "incomplete");
});

test("formula coverage, numeric total and independent expectations use actual readback", async () => {
  const verification = createTaskVerification(
    workbook({ B2: 0.1, B3: 0.2 }, { B2: "=A2/10", B3: "=A3/10" }),
  );
  await verification.define([
    { type: "formulas", label: "derived values remain formulas", target: target("B2:B3") },
    {
      type: "sum",
      label: "source total",
      target: target("B2:B3"),
      expected: 0.3,
      tolerance: 1e-10,
    },
    {
      type: "matches",
      label: "independent values",
      target: target("B2:B3"),
      expected: [[0.1], [0.2]],
    },
  ]);
  assert.equal((await verification.run()).passed, 3);
});

test("MCP checks are rerun after later writes and reset for the next user turn", async () => {
  const cells = { A1: 10 };
  const verification = createTaskVerification(workbook(cells));
  const server = createOfficeBridgeMcp(
    {
      async callTaskpaneTool(name, args) {
        assert.equal(name, "excel_set_cell_range");
        cells.A1 = args.cells[0][0].value;
        return { success: true, commitStatus: "committed" };
      },
    },
    "excel",
    "book",
    { verification },
  );
  const tools = server.instance._registeredTools;
  await tools.excel_verify_task.handler({
    action: "define",
    checks: [{ type: "matches", label: "expected value", target: target("A1"), expected: [[10]] }],
  });
  assert.equal((await tools.excel_verify_task.handler({ action: "run" })).isError, undefined);
  await tools.excel_set_cell_range.handler({ sheetId: 1, range: "A1", cells: [[{ value: 11 }]] });
  assert.equal(
    (await verification.finish()).status,
    "failed",
    "a previous pass cannot survive a later wrong write",
  );
  verification.reset();
  assert.equal(await verification.finish(), null);
  await tools.excel_set_cell_range.handler({ sheetId: 1, range: "A1", cells: [[{ value: 12 }]] });
  assert.equal((await verification.finish()).status, "not_checked");
});
