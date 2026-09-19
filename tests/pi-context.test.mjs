import { test } from "node:test";
import assert from "node:assert/strict";
import { buildOverview, ChangeTracker, readSelectionContext } from "../taskpane/shared/vendor/pi-context.js";

test("selection context reads the address submitted with the turn", async () => {
  const originalExcel = globalThis.Excel;
  const requestedRanges = [];
  const sheetName = "O'Brien!2026";
  const selectedRange = {
    address: "'O''Brien!2026'!B2",
    rowIndex: 1,
    columnIndex: 1,
    rowCount: 1,
    columnCount: 1,
    worksheet: { name: sheetName },
    load() {},
    getCell() {
      return { getResizedRange: () => ({ formulas: [["=1+2"]], load() {} }) };
    },
  };
  const sheet = {
    getUsedRangeOrNullObject() {
      return { isNullObject: false, rowIndex: 0, columnIndex: 0, rowCount: 3, columnCount: 3, load() {} };
    },
    getRange(address) {
      requestedRanges.push(address);
      if (address === "B2") return selectedRange;
      if (address === "A1:C3") {
        return {
          address,
          values: [["Header", "Value", "Other"], [1, 3, ""], [2, "#DIV/0!", ""]],
          load() {},
        };
      }
      throw new Error(`Unexpected range ${address}`);
    },
  };
  globalThis.Excel = {
    async run(callback) {
      return callback({
        workbook: {
          getSelectedRange() { throw new Error("Live selection must not be read"); },
          worksheets: {
            getItem(name) {
              assert.equal(name, sheetName);
              return sheet;
            },
          },
        },
        async sync() {},
      });
    },
  };
  try {
    const result = await readSelectionContext("'O''Brien!2026'!B2");
    assert.deepEqual(requestedRanges, ["B2", "A1:C3"]);
    assert.match(result.text, /Selection:.*B2/);
    assert.match(result.text, /B2: =1\+2/);
    assert.match(result.text, /Errors nearby:.*B3=#DIV\/0!/);
  } finally {
    globalThis.Excel = originalExcel;
  }
});

test("overview bounds the header read and counts shapes without loading them", async () => {
  const originalExcel = globalThis.Excel;
  const ranges = [];
  const sheet = {
    name: "Data",
    position: 0,
    visibility: "Visible",
    getUsedRangeOrNullObject() {
      return { isNullObject: false, rowCount: 2, columnCount: 30, load() {} };
    },
    getRange(address) {
      ranges.push(address);
      return { getUsedRangeOrNullObject: () => ({ isNullObject: false, values: [["A", "B"]], load() {} }) };
    },
    tables: { items: [], load() {} },
    charts: { count: 0, load() {} },
    pivotTables: { getCount: () => ({ value: 0 }) },
    shapes: { getCount: () => ({ value: 2 }) },
  };
  globalThis.Excel = {
    async run(callback) {
      return callback({
        workbook: {
          name: "Book.xlsx",
          load() {},
          worksheets: { items: [sheet], load() {} },
          names: { items: [], load() {} },
        },
        async sync() {},
      });
    },
  };
  try {
    const overview = await buildOverview();
    assert.deepEqual(ranges, ["A1:T1"]);
    assert.match(overview, /2 shape\(s\)/);
  } finally {
    globalThis.Excel = originalExcel;
  }
});

test("change tracker keeps user edits and skips this add-in's own writes", async () => {
  const originalExcel = globalThis.Excel;
  let onChanged;
  const sheet = { id: "s1", name: "Data", visibility: "Visible", onChanged: { add: (handler) => (onChanged = handler) } };
  globalThis.Excel = {
    async run(callback) {
      return callback({ workbook: { worksheets: { items: [sheet], load() {} } }, async sync() {} });
    },
  };
  try {
    const tracker = new ChangeTracker();
    await tracker.start();
    await onChanged({ worksheetId: "s1", address: "B2", triggerSource: "Unknown" });
    await onChanged({ worksheetId: "s1", address: "C3", triggerSource: "ThisLocalAddin" });
    await onChanged({ worksheetId: "s1", address: "D4" });
    const flushed = tracker.flush();
    assert.match(flushed, /Data: B2, D4/);
    assert.doesNotMatch(flushed, /C3/);
  } finally {
    globalThis.Excel = originalExcel;
  }
});
