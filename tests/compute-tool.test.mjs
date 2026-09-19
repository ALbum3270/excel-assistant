import { test } from "node:test";
import assert from "node:assert/strict";
import { createComputeShell } from "../daemon/compute-tool.mjs";

test("csv-to-sheet writes parsed values through the workbook bridge", async () => {
  const calls = [];
  const shell = createComputeShell(async (name, args) => {
    calls.push({ name, args });
    return {
      success: true,
      commitStatus: "committed",
      writtenRange: args.range,
    };
  });

  const result = await shell({
    command: "printf 'Name,Amount\\nAlice,12.5\\nBob,7\\n' > out.csv && csv-to-sheet out.csv 1 F2 --force",
  });

  assert.equal(result.exitCode, 0);
  assert.match(result.stdout, /Committed 3 rows x 2 columns/);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].name, "excel_set_cell_range");
  assert.equal(calls[0].args.range, "F2:G4");
  assert.deepEqual(calls[0].args.cells, [
    [{ value: "Name" }, { value: "Amount" }],
    [{ value: "Alice" }, { value: 12.5 }],
    [{ value: "Bob" }, { value: 7 }],
  ]);
});

test("csv-to-sheet accepts a valid one-column CSV", async () => {
  const calls = [];
  const shell = createComputeShell(async (name, args) => {
    calls.push({ name, args });
    return { success: true, commitStatus: "committed", writtenRange: args.range };
  });

  const result = await shell({
    command: "printf 'Alpha\\nBeta\\n' > one.csv && csv-to-sheet one.csv 1 A1 --force",
  });

  assert.equal(result.exitCode, 0);
  assert.equal(calls[0].args.range, "A1:A2");
  assert.deepEqual(calls[0].args.cells, [[{ value: "Alpha" }], [{ value: "Beta" }]]);
});

test("csv-to-sheet finishes every chunk and then lists formula errors", async () => {
  const ranges = [];
  const shell = createComputeShell(async (name, args) => {
    ranges.push(args.range);
    const formulaErrors = ranges.length === 1 ? [{ address: "A1", value: "#N/A" }] : [];
    return { success: true, commitStatus: "committed", writtenRange: args.range, formulaErrors };
  });

  // 2001 one-column rows -> two chunks (2000 + 1); the first reports #N/A.
  const result = await shell({
    command: "seq 1 2001 > n.csv && csv-to-sheet n.csv 1 A1 --force",
  });

  assert.equal(result.exitCode, 0);
  assert.deepEqual(ranges, ["A1:A2000", "A2001:A2001"]);
  assert.match(result.stdout, /Committed 2001 rows x 1 columns .* 2 committed chunk/);
  assert.match(result.stdout, /Formula errors in 1 cell\(s\): A1=#N\/A/);
});
