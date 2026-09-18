import { test } from "node:test";
import assert from "node:assert/strict";
import { toolExcelAddTableRows, toolExcelAutoFilter } from "../taskpane/shared/tools-excel.js";

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
