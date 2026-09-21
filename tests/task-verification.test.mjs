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

test("a rejected plan names every problem at once and explains the range arithmetic", async () => {
  const verification = createTaskVerification(workbook({}));
  const error = await verification
    .define([
      { type: "matches", label: "pairs 1-2", target: target("C2:D4"), expected: [[1, 2], [3, 4]] },
      { type: "row_count", label: "no tail", target: target("C2:C50") },
      { type: "same_rows", label: "same rows", target: target("A2:B9") },
    ])
    .then(() => null, (e) => e);
  assert.match(error.message, /3 problems/);
  assert.match(error.message, /rows 2-4 .*3 x 2.*expected. is 2 x 2/s);
  assert.match(error.message, /set target to C2:D3/);
  assert.match(error.message, /row_count needs `expected`/);
  assert.match(error.message, /same_rows needs `source`/);
});

test("a single-cell target takes its size from the expected matrix", async () => {
  const verification = createTaskVerification(
    workbook({ C2: "a", D2: 1, C3: "b", D3: 2 }),
  );
  await verification.define([
    { type: "matches", label: "block", target: target("C2"), expected: [["a", 1], ["b", 2]] },
  ]);
  const result = await verification.run();
  assert.equal(result.status, "passed");
  assert.equal(result.checks[0].target.range, "C2:D3");
});

test("a sheet id sent as text is accepted rather than bouncing the call", async () => {
  const verification = createTaskVerification(workbook({ A2: "x" }));
  await verification.define([
    { type: "not_blank", label: "filled", target: { sheetId: "1", range: "A2" } },
  ]);
  assert.equal((await verification.run()).status, "passed");
});

test("a one-cell matches check takes a single value; a larger one still needs rows", async () => {
  const verification = createTaskVerification(workbook({ I2: 305, I3: 7 }));
  await verification.define([
    { type: "matches", label: "one cell", target: target("I2"), expected: 305 },
  ]);
  const result = await verification.run();
  assert.equal(result.status, "passed");

  const error = await verification
    .define([{ type: "matches", label: "two cells", target: target("I2:I3"), expected: 305 }])
    .then(() => null, (e) => e);
  assert.match(error.message, /single value only fits a one-cell target; I2:I3 is 2 x 1/);
});

test("a sum sent as plain numeric text is accepted, but not with separators or units", async () => {
  const verification = createTaskVerification(workbook({ B2: 300, B3: 5 }));
  await verification.define([
    { type: "sum", label: "total", target: target("B2:B3"), expected: "305" },
  ]);
  assert.equal((await verification.run()).status, "passed");

  for (const expected of ["1,234", "305 kg", "abc"]) {
    const error = await verification
      .define([{ type: "sum", label: "total", target: target("B2:B3"), expected }])
      .then(() => null, (e) => e);
    assert.match(error.message, /sum `expected` must be a number/, expected);
  }
});
