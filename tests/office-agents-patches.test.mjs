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
      const result = ranges[address];
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
