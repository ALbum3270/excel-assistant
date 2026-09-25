import { test } from "node:test";
import assert from "node:assert/strict";
import { createComputeShell } from "../daemon/compute-tool.mjs";
import { getFormulaFlags } from "../taskpane/shared/formula-flags.js";

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
    command:
      "printf 'Name,Amount\\nAlice,12.5\\nBob,7\\n' > out.csv && csv-to-sheet out.csv 1 F2 --force",
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

test("csv-to-sheet imports a large valid row set without spreading it onto the JS call stack", async () => {
  let rowsWritten = 0;
  const shell = createComputeShell(async (_name, args) => {
    rowsWritten += args.cells.length;
    return { success: true, commitStatus: "committed", writtenRange: args.range };
  });
  const result = await shell({
    command: `python3 - <<'PY'
with open('large.csv', 'w') as stream:
    stream.write('x\\n' * 200000)
PY
csv-to-sheet large.csv 1 A1 --force --text`,
  });
  assert.equal(result.exitCode, 0, result.stderr);
  assert.equal(rowsWritten, 200_000);
});

test("csv-to-sheet rejects a target that would cross Excel's last row before writing", async () => {
  let writes = 0;
  const shell = createComputeShell(async () => {
    writes += 1;
    return { success: true, commitStatus: "committed" };
  });
  const result = await shell({
    command: `printf 'one\ntwo' > rows.csv
csv-to-sheet rows.csv 1 A1048576 --force --text`,
  });
  assert.equal(result.exitCode, 1);
  assert.match(result.stderr, /outside Excel's limits/);
  assert.equal(writes, 0);
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

test("csv-to-sheet keeps lossy numbers as text and offers exact text on request", async () => {
  const writes = [];
  const shell = createComputeShell(async (name, args) => {
    if (name === "excel_set_cell_range") writes.push(args.cells);
    return { success: true, commitStatus: "committed", writtenRange: args.range };
  });
  // The four values the audit round-tripped, plus a date and a plain number.
  await shell({
    command:
      "printf '00123,9007199254740993,TRUE,=1+1,2026-03-04,42\n' > t.csv && csv-to-sheet t.csv 1 A1 --force",
  });
  assert.deepEqual(writes[0], [
    [
      { value: "'00123" },
      { value: "'9007199254740993" },
      { value: true },
      { formula: "=1+1" },
      { value: "2026-03-04" },
      { value: 42 },
    ],
  ]);

  await shell({ command: "csv-to-sheet t.csv 1 A1 --force --no-formulas" });
  assert.deepEqual(writes[1][0][3], { value: "'=1+1" });

  await shell({ command: "csv-to-sheet t.csv 1 A1 --force --text" });
  assert.deepEqual(writes[2][0], [
    { value: "'00123" },
    { value: "'9007199254740993" },
    { value: "'TRUE" },
    { value: "'=1+1" },
    { value: "'2026-03-04" },
    { value: "'42" },
  ]);
});

test("typed JSON round-trip preserves mixed text, numbers, booleans and formulas", async () => {
  const writes = [];
  const shell = createComputeShell(async (name, args) => {
    if (name === "excel_get_formula_flags") {
      assert.deepEqual(args.addresses, ["B2", "D2"]);
      return { success: true, flags: { B2: false, D2: true } };
    }
    if (name === "excel_get_cell_ranges") {
      return args.ranges[0] === "A1:D2"
        ? {
            success: true,
            worksheet: { name: "Data", cells: { A1: "2026-03-04", B1: 42, C1: true } },
            remainingRanges: ["D1:D2"],
          }
        : {
            success: true,
            worksheet: {
              name: "Data",
              cells: { D1: 2, A2: "00123", B2: "=literal", D2: "=A1" },
              formulas: { D1: "=1+1", B2: "=literal", D2: "=A1" },
            },
            remainingRanges: [],
          };
    }
    writes.push(args);
    return { success: true, commitStatus: "committed", writtenRange: args.range };
  });
  const result = await shell({
    command: "sheet-to-json 1 A1:D2 data.json && json-to-sheet data.json 1 F1 --force",
  });
  assert.equal(result.exitCode, 0, result.stderr);
  assert.equal(writes.length, 1);
  assert.equal(writes[0].range, "F1:I2");
  assert.deepEqual(writes[0].cells, [
    [{ value: "'2026-03-04" }, { value: 42 }, { value: true }, { formula: "=1+1" }],
    [{ value: "'00123" }, { value: "'=literal" }, { value: "" }, { formula: "=A1" }],
  ]);
});

test("assert-sheet-json compares an independently prepared matrix with live cells", async () => {
  const assertions = [];
  const shell = createComputeShell(
    async (name) => {
      assert.equal(name, "excel_get_cell_ranges");
      return {
        success: true,
        worksheet: { name: "Data", cells: { C2: 10, C3: 20 } },
        remainingRanges: [],
      };
    },
    { onAssert: (target) => assertions.push(target) },
  );
  const prepared = await shell({
    command: `python3 - <<'PY'
import json
with open('expected.json', 'w') as f:
    json.dump({'format':'excel-typed-cells-v1','rows':[[{'value':10}],[{'value':20}]]}, f)
PY
assert-sheet-json expected.json 1 C2:C3`,
  });
  assert.equal(prepared.exitCode, 0, prepared.stderr);
  assert.deepEqual(assertions, [{ sheetId: 1, range: "C2:C3" }]);
  const mismatch = await shell({
    command: `python3 - <<'PY'
import json
with open('wrong.json', 'w') as f:
    json.dump({'format':'excel-typed-cells-v1','rows':[[{'value':10}],[{'value':21}]]}, f)
PY
assert-sheet-json wrong.json 1 C2:C3`,
  });
  assert.equal(mismatch.exitCode, 1);
  assert.match(mismatch.stderr, /C3: expected.*21.*got.*20/);
  assert.equal(assertions.length, 1);
});

test("formula-presence check distinguishes literal equals text from formulas", async () => {
  const original = globalThis.Excel;
  globalThis.Excel = {
    run: async (callback) =>
      callback({
        workbook: {
          worksheets: {
            getItem: () => ({
              getRange: (address) => ({
                getSpecialCellsOrNullObject: () => ({
                  isNullObject: address === "B2",
                  load() {},
                }),
              }),
            }),
          },
        },
        async sync() {},
      }),
  };
  try {
    const result = await getFormulaFlags("Data", ["B2", "D2"]);
    assert.deepEqual(result.flags, { B2: false, D2: true });
  } finally {
    globalThis.Excel = original;
  }
});

test("a write from a saved script still asks for approval at the moment it writes", async () => {
  const writes = [];
  const asked = [];
  let decision = "reject";
  const shell = createComputeShell(
    async (name, args) => {
      if (name === "excel_set_cell_range") writes.push(args.range);
      return { success: true, commitStatus: "committed", writtenRange: args.range };
    },
    {
      approveWrite: async (toolName, input) => {
        asked.push(input.range);
        return decision;
      },
    },
  );
  // The audit's bypass: save the write in one call, run the script in another.
  await shell({
    command: "printf 'a,b\n1,2\n' > t.csv && echo 'csv-to-sheet t.csv 1 A1 --force' > later.sh",
  });
  const rejected = await shell({ command: "bash later.sh" });
  assert.deepEqual(asked, ["A1:B2"], "the script's write asked for approval");
  assert.equal(writes.length, 0, "a rejected write never reaches Excel");
  assert.match(rejected.stderr, /rejected this workbook change/);

  decision = "approve";
  await shell({ command: "bash later.sh" });
  assert.deepEqual(writes, ["A1:B2"]);
});

test("sheet-to-csv keeps a trailing page that is one blank cell", async () => {
  // Additional audit N02: the last page is a single empty cell, which
  // serializes to "", and must still count as a row on the way back in.
  const writes = [];
  const reads = [];
  const shell = createComputeShell(
    async (name, args) => {
      if (name === "excel_get_range_as_csv") {
        return args.range === "A1:A3"
          ? {
              csv: "row\nrow",
              rowCount: 2,
              columnCount: 1,
              sheetName: "S",
              hasMore: true,
              nextRange: "A3:A3",
            }
          : {
              csv: "",
              rowCount: 1,
              columnCount: 1,
              sheetName: "S",
              hasMore: false,
              nextRange: null,
            };
      }
      writes.push(args);
      return { success: true, commitStatus: "committed", writtenRange: args.range };
    },
    { onRead: (target) => reads.push(target) },
  );
  const result = await shell({
    command: "sheet-to-csv 1 A1:A3 data.csv && csv-to-sheet data.csv 1 B1 --force --text",
  });
  assert.equal(result.exitCode, 0, result.stderr);
  assert.match(result.stdout, /Exported 3 rows/);
  assert.equal(writes[0].range, "B1:B3");
  assert.equal(writes[0].cells.length, 3);
  assert.deepEqual(reads, [{ sheetId: 1, range: "A1:A3" }]);
});

test("CSV exported from a range cannot silently overwrite that range with inferred types", async () => {
  const writes = [];
  const shell = createComputeShell(async (name, args) => {
    if (name === "excel_get_range_as_csv") {
      return {
        success: true,
        csv: "2026-03-04,42",
        rowCount: 1,
        columnCount: 2,
        sheetName: "Data",
        hasMore: false,
      };
    }
    writes.push(args.range);
    return { success: true, commitStatus: "committed", writtenRange: args.range };
  });
  const refused = await shell({
    command: "sheet-to-csv 1 A1:B1 data.csv && csv-to-sheet data.csv 1 A1 --force",
  });
  assert.equal(refused.exitCode, 1);
  assert.match(refused.stderr, /typed round trip/);
  assert.deepEqual(writes, []);
  const deliberate = await shell({
    command: "csv-to-sheet data.csv 1 A1 --force --allow-type-loss",
  });
  assert.equal(deliberate.exitCode, 0, deliberate.stderr);
  assert.deepEqual(writes, ["A1:B1"]);
});
