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
  assert.match(result.stdout, /Wrote and verified 3 rows x 2 columns/);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].name, "excel_set_cell_range");
  assert.equal(calls[0].args.range, "F2:G4");
  assert.deepEqual(calls[0].args.cells, [
    [{ value: "Name" }, { value: "Amount" }],
    [{ value: "Alice" }, { value: 12.5 }],
    [{ value: "Bob" }, { value: 7 }],
  ]);
});
