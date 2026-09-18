import { test } from "node:test";
import assert from "node:assert/strict";
import {
  toolExcelAddTableRows,
  toolExcelAutoFilter,
  toolExcelGetSelectedRange,
} from "../taskpane/shared/tools-excel.js";

function installExcel() {
  const calls = [];
  const sheets = new Map();
  const worksheet = (name) => {
    if (!sheets.has(name)) {
      sheets.set(name, {
        name,
        autoFilter: {
          apply: (target) => calls.push(["apply", name, target.address]),
          remove: () => calls.push(["remove", name]),
        },
        getRange: (address) => ({ address }),
      });
    }
    return sheets.get(name);
  };
  const context = {
    workbook: {
      worksheets: {
        getItem: worksheet,
        getActiveWorksheet: () => ({ name: "Active", load() {} }),
      },
      tables: { getItem: () => ({ rows: { add() {} } }) },
    },
    async sync() {},
  };
  globalThis.Excel = { run: async (callback) => callback(context) };
  return calls;
}

test("autofilter honors a sheet-qualified address, including escaped apostrophes", async () => {
  const calls = installExcel();
  const result = await toolExcelAutoFilter({ address: "'O''Brien Data'!A1:C8", sheet: "Wrong" });
  assert.deepEqual(calls, [["apply", "O'Brien Data", "A1:C8"]]);
  assert.deepEqual(result, {
    sheet: "O'Brien Data",
    autofilter: "applied",
    address: "A1:C8",
  });
});

test("table row input fails before Excel.run when it is empty, ragged, or has a bad index", async () => {
  let runCalls = 0;
  globalThis.Excel = {
    run: async () => {
      runCalls++;
    },
  };
  await assert.rejects(toolExcelAddTableRows({ table: "T", values: [[]] }), /non-empty/);
  await assert.rejects(toolExcelAddTableRows({ table: "T", values: [[1, 2], [3]] }), /rectangular/);
  await assert.rejects(toolExcelAddTableRows({ table: "T", values: [[1]], index: -1 }), /index/);
  assert.equal(runCalls, 0);
});

test("selected-range reads a bounded top-left preview instead of the full selection", async () => {
  const loads = [];
  const resizeCalls = [];
  const preview = {
    address: "Sheet1!A1:P125",
    values: [["bounded preview"]],
    load(properties) {
      loads.push(["preview", properties]);
    },
  };
  const selected = {
    address: "Sheet1!A1:P1048576",
    rowCount: 1_048_576,
    columnCount: 16,
    worksheet: { name: "Sheet1" },
    load(properties) {
      loads.push(["selection", properties]);
    },
    getCell(row, column) {
      assert.deepEqual([row, column], [0, 0]);
      return {
        getResizedRange(rowDelta, columnDelta) {
          resizeCalls.push([rowDelta, columnDelta]);
          return preview;
        },
      };
    },
  };
  let syncs = 0;
  globalThis.Excel = {
    run: async (callback) =>
      callback({
        workbook: { getSelectedRange: () => selected },
        async sync() {
          syncs++;
        },
      }),
  };

  const result = await toolExcelGetSelectedRange();
  assert.deepEqual(loads, [
    ["selection", "address, rowCount, columnCount, worksheet/name"],
    ["preview", "address, values"],
  ]);
  assert.deepEqual(resizeCalls, [[124, 15]]);
  assert.equal(syncs, 2);
  assert.equal(result.total_cell_count, 16_777_216);
  assert.equal(result.preview_address, "Sheet1!A1:P125");
  assert.equal(result.truncated, true);
  assert.deepEqual(result.values, [["bounded preview"]]);
});

test("selected-range rejects invalid preview limits before Excel.run", async () => {
  let runCalls = 0;
  globalThis.Excel = { run: async () => runCalls++ };
  await assert.rejects(toolExcelGetSelectedRange({ cellLimit: 0 }), /1 to 5000/);
  await assert.rejects(toolExcelGetSelectedRange({ cellLimit: 5001 }), /1 to 5000/);
  assert.equal(runCalls, 0);
});
