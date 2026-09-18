import { test, before } from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import { patchOfficeAgentsExcelApi } from "../scripts/office-agents-patches.mjs";

let api;
const settings = new Map();

before(async () => {
  globalThis.Office = {
    context: {
      document: {
        settings: {
          refreshAsync: (done) => done(),
          get: (key) => settings.get(key),
          set: (key, value) => settings.set(key, value),
          saveAsync: (done) => done(),
        },
      },
    },
  };
  api = await import("../taskpane/shared/vendor/office-agents-excel-api.js");
});

function range(address, initialValues, initialFormulas = initialValues) {
  let currentValues = initialValues.map((row) => [...row]);
  let currentFormulas = initialFormulas.map((row) => [...row]);
  const state = { valueAssignments: 0, formulaAssignments: 0, copies: 0 };
  const object = {
    address: `Sheet1!${address}`,
    rowCount: currentValues.length,
    columnCount: currentValues[0]?.length ?? 0,
    load() {},
    getCell() {
      return {
        address: "Sheet1!A1",
        load() {},
        format: {
          font: {},
          fill: {},
          borders: { getItem: () => ({}) },
        },
      };
    },
    getEntireColumn: () => ({ format: {} }),
    getEntireRow: () => ({ format: {} }),
    copyFrom() {
      state.copies++;
    },
    state,
  };
  Object.defineProperties(object, {
    values: {
      get: () => currentValues,
      set: (value) => {
        state.valueAssignments++;
        currentValues = value;
      },
    },
    formulas: {
      get: () => currentFormulas,
      set: (value) => {
        state.formulaAssignments++;
        currentFormulas = value;
      },
    },
  });
  return object;
}

function installExcel(ranges, { usedAddress = "A1:A3" } = {}) {
  let runCalls = 0;
  const sheet = {
    id: "sheet-guid",
    name: "Sheet1",
    load() {},
    notes: { add() {} },
    getUsedRangeOrNullObject: () => ({
      isNullObject: false,
      address: `Sheet1!${usedAddress}`,
      load() {},
    }),
    getRange(address) {
      const result = typeof ranges === "function" ? ranges(address) : ranges[address];
      assert.ok(result, `Unexpected range request: ${address}`);
      return result;
    },
  };
  const context = {
    workbook: {
      worksheets: {
        items: [sheet],
        load() {},
      },
    },
    async sync() {},
  };
  globalThis.Excel = {
    run: async (callback) => {
      runCalls++;
      return callback(context);
    },
    RangeCopyType: { all: "all" },
    BorderIndex: { edgeTop: "top", edgeBottom: "bottom", edgeLeft: "left", edgeRight: "right" },
    BorderLineStyle: { continuous: "continuous", dash: "dash", dot: "dot", double: "double" },
    BorderWeight: { thin: "thin", medium: "medium", thick: "thick" },
  };
  return {
    get runCalls() {
      return runCalls;
    },
  };
}

test("patches match the pinned upstream source and reject already-patched input", async () => {
  const upstream = await readFile(
    new URL("../../_sdks/office-agents/packages/excel/src/lib/excel/api.ts", import.meta.url),
    "utf8",
  );
  const patched = patchOfficeAgentsExcelApi(upstream);
  assert.match(patched, /function throwOverwriteError/);
  assert.match(patched, /scanRange:/);
  assert.match(patched, /allowOverwrite = false/);
  assert.throws(() => patchOfficeAgentsExcelApi(patched), /expected exactly one match/);
  const generated = await readFile(
    new URL("../taskpane/shared/vendor/office-agents-excel-api.js", import.meta.url),
    "utf8",
  );
  assert.match(generated, /office-agents @ 95fb654491a9d394dc85ea2b8c93dee2ca4546b9/);
});

test("getCellRanges reports truncation within one range", async () => {
  installExcel({
    "A1:A3": range("A1:A3", [[1], [2], [3]]),
  });
  const result = await api.getCellRanges(1, ["A1:A3"], {
    includeStyles: false,
    cellLimit: 2,
  });
  assert.equal(result.hasMore, true);
  assert.deepEqual(result.worksheet.cells, { A1: 1, A2: 2 });
});

// Generate values only when Office.js requests them, so an accidental
// full-range load fails before allocating a whole-sheet matrix.
function installReadGrid(valueAt, maxChunk) {
  const reads = [];
  const col = (letters) => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  installExcel((address) => {
    const match = address.match(/^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/);
    assert.ok(match, address);
    const firstRow = Number(match[2]);
    const firstCol = col(match[1]);
    const rowCount = Number(match[4] ?? match[2]) - firstRow + 1;
    const columnCount = col(match[3] ?? match[1]) - firstCol + 1;
    const target = { address: `Sheet1!${address}`, rowCount, columnCount };
    target.load = (properties) => {
      if (!properties.includes("values")) return;
      const size = rowCount * columnCount;
      assert.ok(size <= maxChunk, `Unbounded host read: ${size} cells`);
      reads.push(size);
      target.values = Array.from({ length: rowCount }, (_, r) =>
        Array.from({ length: columnCount }, (_, c) => valueAt(firstRow + r, firstCol + c)),
      );
      target.formulas = target.values;
    };
    return target;
  });
  return reads;
}

test("bounded sparse pages preserve every cell and stop scanning huge blank ranges", async () => {
  const reads = installReadGrid((r, c) => r * 10 + c, 2000);
  const cells = {};
  let ranges = ["A1:C1001"];
  let pages = 0;
  do {
    const page = await api.getCellRanges(1, ranges, { includeStyles: false, cellLimit: 700 });
    for (const [address, value] of Object.entries(page.worksheet.cells)) {
      assert.equal(cells[address], undefined, `Duplicate: ${address}`);
      cells[address] = value;
    }
    ranges = page.remainingRanges;
    assert.equal(page.hasMore, ranges.length > 0);
    assert.ok(++pages < 10, "Continuation must advance");
  } while (ranges.length);
  assert.equal(Object.keys(cells).length, 3003);
  assert.equal(cells.C1001, 10013);
  assert.ok(reads.every((size) => size <= 2000));

  const blankReads = installReadGrid(() => "", 2000);
  const blank = await api.getCellRanges(1, ["A1:XFD1048576"], { includeStyles: false });
  assert.equal(
    blankReads.reduce((a, b) => a + b, 0),
    20000,
  );
  assert.equal(blank.hasMore, true);
  assert.equal(blank.remainingRanges[0], "EIC2:XFD2");
  assert.deepEqual(blank.worksheet.cells, {});
});

test("CSV pages bound host reads and continue after the skipped header without losing rows", async () => {
  const reads = installReadGrid((r, c) => `${r},${c}`, 20000);
  const first = await api.getRangeAsCsv(1, "A1:D10001", {
    includeHeaders: false,
    maxRows: 20000,
  });
  assert.equal(first.rowCount, 5000);
  assert.equal(first.nextRange, "A5002:D10001");
  assert.equal(first.csv.split("\n")[0], '"2,1","2,2","2,3","2,4"');
  const second = await api.getRangeAsCsv(1, first.nextRange, { maxRows: 20000 });
  assert.equal(second.rowCount, 5000);
  assert.equal(second.hasMore, false);
  assert.equal(second.nextRange, null);
  assert.equal(second.csv.split("\n")[0], '"5002,1","5002,2","5002,3","5002,4"');
  assert.deepEqual(reads, [20000, 20000]);
});

test("a formula with an empty displayed value counts toward cellLimit", async () => {
  installExcel(
    {
      "A1:A2": range("A1:A2", [[""], [7]], [['=IF(TRUE,"",1)'], [7]]),
    },
    { usedAddress: "A1:A2" },
  );
  const result = await api.getCellRanges(1, ["A1:A2"], {
    includeStyles: false,
    cellLimit: 1,
  });
  assert.equal(result.hasMore, true);
  assert.deepEqual(result.worksheet.cells, {});
  assert.deepEqual(result.worksheet.formulas, { A1: '=IF(TRUE,"",1)' });
});

test("getCellRanges only reports more data when a later range contains a populated cell", async () => {
  installExcel({
    A1: range("A1", [[1]]),
    B1: range("B1", [[""]]),
  });
  const exact = await api.getCellRanges(1, ["A1", "B1"], {
    includeStyles: false,
    cellLimit: 1,
  });
  assert.equal(exact.hasMore, false);

  installExcel({
    A1: range("A1", [[1]]),
    B1: range("B1", [[2]]),
  });
  const truncated = await api.getCellRanges(1, ["A1", "B1"], {
    includeStyles: false,
    cellLimit: 1,
  });
  assert.equal(truncated.hasMore, true);
});

test("invalid range and cell matrix inputs fail before Excel.run", async () => {
  const excel = installExcel({});
  await assert.rejects(api.getCellRanges(1, [], {}), /non-empty array/);
  await assert.rejects(api.getCellRanges(1, ["A1"], { cellLimit: 0 }), /positive integer/);
  await assert.rejects(api.setCellRange(1, "A1", []), /non-empty rectangular/);
  await assert.rejects(api.setCellRange(1, "A1", [[{ value: 1 }], []]), /rectangular/);
  await assert.rejects(api.setCellRange(1, "A1", [[{ formula: "SUM(A1:A2)" }]]), /start with/);
  await assert.rejects(
    api.modifySheetStructure(1, { operation: "delete", dimension: "rows" }),
    /reference is required/,
  );
  await assert.rejects(
    api.modifySheetStructure(1, {
      operation: "hide",
      dimension: "rows",
      reference: "2x",
    }),
    /positive row number/,
  );
  await assert.rejects(
    api.modifySheetStructure(1, {
      operation: "insert",
      dimension: "columns",
      reference: "XFE",
    }),
    /between A and XFD/,
  );
  assert.equal(excel.runCalls, 0);
});

test("style-only edits are not blocked by an existing value", async () => {
  const target = range("A1", [[99]]);
  installExcel({ A1: target });
  await api.setCellRange(1, "A1", [[{ cellStyles: { fontWeight: "bold" } }]]);
  assert.equal(target.state.valueAssignments, 0);
  assert.equal(target.state.formulaAssignments, 0);
  assert.deepEqual(target.values, [[99]]);
});

test("sparse writes preserve existing values in cells that only change style", async () => {
  const target = range("A1:B1", [[10, 20]]);
  installExcel({ "A1:B1": target });
  await api.setCellRange(1, "A1:B1", [[{ value: 11 }, { cellStyles: { fontWeight: "bold" } }]], {
    allowOverwrite: true,
  });
  assert.equal(target.state.valueAssignments, 0);
  assert.equal(target.state.formulaAssignments, 1);
  assert.deepEqual(target.formulas, [[11, 20]]);
});

test("an explicit null is treated as a protected destructive write", async () => {
  const target = range("A1", [[99]]);
  installExcel({ A1: target });
  await assert.rejects(api.setCellRange(1, "A1", [[{ value: null }]]), /Would overwrite.*A1/);
  assert.equal(target.state.valueAssignments, 0);
  assert.equal(target.state.formulaAssignments, 0);
});

test("setCellRange checks copyToRange before writing the base range", async () => {
  const base = range("A1", [[""]]);
  const destination = range("A1:A3", [[""], ["occupied"], [""]]);
  installExcel({ A1: base, "A1:A3": destination });
  await assert.rejects(
    api.setCellRange(1, "A1", [[{ value: 1 }]], { copyToRange: "A1:A3" }),
    /Would overwrite.*A2/,
  );
  assert.equal(base.state.valueAssignments, 0);
  assert.equal(destination.state.copies, 0);
});

test("copyTo protects destination data unless overwrite is authorized", async () => {
  const source = range("A1", [[1]]);
  const destination = range("B1:B2", [["occupied"], [""]]);
  installExcel({ A1: source, "B1:B2": destination });
  await assert.rejects(api.copyTo(1, "A1", "B1:B2"), /Would overwrite.*B1/);
  assert.equal(destination.state.copies, 0);
  await api.copyTo(1, "A1", "B1:B2", true);
  assert.equal(destination.state.copies, 1);
});

test("copy protection permits source cells inside a larger destination", async () => {
  const source = range("A1", [[1]]);
  const destination = range("A1:A3", [[1], [""], [""]]);
  installExcel({ A1: source, "A1:A3": destination });
  await api.copyTo(1, "A1", "A1:A3");
  assert.equal(destination.state.copies, 1);
});
