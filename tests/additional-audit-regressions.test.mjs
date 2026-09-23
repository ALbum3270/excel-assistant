// Regression tests for docs/additional-bug-audit-2026-09-22.md (N01–N05; N06 is
// in taskpane.test.mjs). The fixtures are the audit's repro fixtures, asserting
// the correct outcome instead of the defect. Office is an in-memory stand-in;
// the table fixture shifts rows the way TableCollection.add is documented to.
import { test, before } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import { createComputeShell } from "../daemon/compute-tool.mjs";
import { toolExcelCreateTable } from "../taskpane/shared/tools-excel.js";

let api;
before(async () => {
  const settings = new Map([["openexcel-sheet-id-map", { "audit-sheet": 1 }]]);
  globalThis.Office = {
    context: {
      document: {
        settings: {
          get: (key) => settings.get(key),
          set: (key, value) => settings.set(key, value),
          refreshAsync: (done) => done(),
          saveAsync: (done) => done(),
        },
      },
    },
  };
  api = await import("../taskpane/shared/vendor/office-agents-excel-api.js");
});

const col = (letters) => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) - 1;
const letters = (n) =>
  n < 0 ? "" : letters(Math.floor(n / 26) - 1) + String.fromCharCode(65 + (n % 26));

// One worksheet whose ranges read from and write to a shared cell map.
function gridHost({ name = "Sheet1", used = "A1", valueAt = () => "" } = {}) {
  const data = new Map();
  const writes = [];
  const get = (r, c) => (data.has(`${r},${c}`) ? data.get(`${r},${c}`) : valueAt(r, c));
  const range = (address) => {
    const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/.exec(address);
    assert.ok(match, address);
    const r = Number(match[2]) - 1;
    const c = col(match[1]);
    const rows = Number(match[4] ?? match[2]) - r;
    const columns = col(match[3] ?? match[1]) - c + 1;
    let cachedValues;
    return {
      address: `'${name.replaceAll("'", "''")}'!${address}`,
      rowCount: rows,
      columnCount: columns,
      isNullObject: false,
      load() {},
      format: { font: {}, fill: {}, borders: { getItem: () => ({}) } },
      get values() {
        return (cachedValues ??= Array.from({ length: rows }, (_, dr) =>
          Array.from({ length: columns }, (_, dc) => get(r + dr, c + dc)),
        ));
      },
      get formulas() {
        return this.values;
      },
      set formulas(matrix) {
        cachedValues = undefined;
        writes.push({ address, matrix });
        matrix.forEach((row, dr) =>
          row.forEach((value, dc) => {
            if (value !== null) data.set(`${r + dr},${c + dc}`, value);
          }),
        );
      },
      getCell(dr, dc) {
        return range(`${letters(c + dc)}${r + dr + 1}`);
      },
      getResizedRange(dr, dc) {
        return range(`${letters(c)}${r + 1}:${letters(c + columns + dc - 1)}${r + rows + dr}`);
      },
    };
  };
  const sheet = {
    id: "audit-sheet",
    name,
    load() {},
    getRange: range,
    getUsedRangeOrNullObject: () => range(used),
  };
  const context = { workbook: { worksheets: { items: [sheet], load() {} } }, sync: async () => {} };
  globalThis.Excel = { run: (fn) => fn(context), RangeCopyType: { all: "All" } };
  return { writes };
}

test("N01: a sheet named with '!' reads the requested cell at its own address", async () => {
  gridHost({
    name: "Sales!2026",
    used: "D5:D6",
    valueAt: (r, c) => (c === 3 && r === 4 ? 42 : ""),
  });
  const result = await api.getCellRanges(1, ["D5"], { includeStyles: false });
  assert.deepEqual(result.worksheet.cells, { D5: 42 });
  assert.equal(result.worksheet.dimension, "D5:D6");
});

function exportThenImport(command) {
  const writes = [];
  const shell = createComputeShell(async (name, args) => {
    if (name === "excel_get_range_as_csv") return api.getRangeAsCsv(args.sheetId, args.range, args);
    writes.push(args);
    return { commitStatus: "committed", writtenRange: args.range };
  });
  return shell({ command }).then((result) => ({ result, writes }));
}

test("N02: a final blank page still comes back as a row", async () => {
  gridHost({ used: "A1:A20001", valueAt: (r) => (r < 20_000 ? "row" : "") });
  const { result, writes } = await exportThenImport(
    "sheet-to-csv 1 A1:A20001 data.csv && csv-to-sheet data.csv 1 B1 --force --text",
  );
  assert.equal(result.exitCode, 0, result.stderr);
  assert.equal(
    writes.reduce((sum, args) => sum + args.cells.length, 0),
    20_001,
  );
  assert.match(writes.at(-1).range, /B20001$/);
});

test("N03: a lone carriage return stays inside its cell through CSV", async () => {
  gridHost({ valueAt: () => "alpha\rbeta" });
  const { result, writes } = await exportThenImport(
    "sheet-to-csv 1 A1 data.csv && csv-to-sheet data.csv 1 B1 --force --text",
  );
  assert.equal(result.exitCode, 0, result.stderr);
  assert.equal(writes.length, 1);
  assert.equal(writes[0].range, "B1:B1");
  assert.equal(writes[0].cells.length, 1);
  assert.match(JSON.stringify(writes[0].cells[0][0]), /alpha\\rbeta/);
});

test("N04: a cell that only gets a style is not written, a neighbour's value is", async () => {
  // Range.formulas returns the value of a constant, so text "=1+1" read from
  // it and written back would be parsed as a formula.
  const host = gridHost({ used: "A1:C1", valueAt: (_r, c) => (c === 1 ? "=1+1" : "") });
  const result = await api.setCellRange(
    1,
    "A1:C1",
    [[{ value: 7 }, { cellStyles: { fontWeight: "bold" } }, { value: 9 }]],
    { allowOverwrite: false },
  );
  assert.equal(result.success, true);
  assert.deepEqual(
    host.writes.map((write) => [write.address, write.matrix]),
    [
      ["A1:A1", [[7]]],
      ["C1:C1", [[9]]],
    ],
  );
  // A matrix with data in every cell stays one write.
  const whole = gridHost({ used: "A1:B1" });
  await api.setCellRange(1, "A1:B1", [[{ value: 1 }, { formula: "=A1*2" }]], {
    allowOverwrite: false,
  });
  assert.deepEqual(whole.writes, [{ address: "A1:B1", matrix: [[1, "=A1*2"]] }]);
});

// Column A of one sheet, top to bottom. tables.add(address, false) inserts a
// generated header above the data and shifts the column down one row, as
// TableCollection.add documents; Range.delete("Up") shifts it back.
function tableHost(initial) {
  const cells = [...initial];
  let table = null;
  const address = (top, rows) => `Sheet1!A${top + 1}${rows > 1 ? `:A${top + rows}` : ""}`;
  const cellRange = (top, rows = 1) => ({
    address: address(top, rows),
    rowCount: rows,
    columnCount: 1,
    load() {},
    getCell: () => cellRange(top),
    getResizedRange: (dr) => cellRange(top, rows + dr),
  });
  const parse = (a1) => {
    const [, first, last] = /^A(\d+)(?::A(\d+))?$/.exec(a1);
    return { top: Number(first) - 1, rows: Number(last ?? first) - Number(first) + 1 };
  };
  const tables = {
    add(a1, hasHeaders) {
      const { top, rows } = parse(a1);
      if (!hasHeaders) cells.splice(top, 0, "Column1");
      table = { name: "T", top, rows: hasHeaders ? rows : rows + 1 };
      return {
        set name(value) {},
        get name() {
          return "T";
        },
        load() {},
      };
    },
    getItem() {
      assert.ok(table, "the table exists");
      return {
        getHeaderRowRange: () => cellRange(table.top),
        getDataBodyRange: () => cellRange(table.top + 1, table.rows - 1),
        convertToRange() {
          table = null;
        },
      };
    },
  };
  const worksheet = {
    name: "Sheet1",
    tables,
    getRange: (a1) => ({
      delete(direction) {
        assert.equal(direction, "Up");
        const { top, rows } = parse(a1);
        cells.splice(top, rows);
      },
    }),
    load() {},
  };
  globalThis.Excel = {
    run: (fn) =>
      fn({
        workbook: {
          tables,
          worksheets: { getItem: () => worksheet, getActiveWorksheet: () => worksheet },
        },
        sync: async () => {},
      }),
  };
  return {
    worksheet,
    cells,
    tableExists: () => Boolean(table),
    tableState: () => ({
      kind: "table_present",
      name: "T",
      address: address(table.top, table.rows),
      hasHeaders: true,
      count: 1,
    }),
  };
}

test("N05: a table made without headers undoes and redoes back to the same cells", async () => {
  const source = readFileSync(
    new URL("../taskpane/shared/recovery.js", import.meta.url),
    "utf8",
  ).replace(/\r\n/g, "\n");
  const start = source.indexOf("async function prepareCustomRecovery(");
  const end = source.indexOf("async function restoreCustomSnapshot(", start);
  assert.ok(start > 0 && end > start);

  // Row 3 holds a value below the table; it must end up where it started.
  const host = tableHost(["A", "B", "X"]);
  const saved = [];
  const context = vm.createContext({
    Excel: globalThis.Excel,
    captureRangeSnapshot: async ({ address }) => ({
      address: `Sheet1!${address}`,
      cellCount: 2,
      beforeValues: host.cells.slice(0, 2).map((value) => [value]),
      beforeFormulas: host.cells.slice(0, 2).map((value) => [value]),
    }),
    captureFormatSnapshot: async () => {
      throw new Error("formats are outside this data fixture");
    },
    appendCustomSnapshot: async (args) => {
      const snapshot = { ...args, id: "table" };
      saved.push(snapshot);
      return snapshot;
    },
    recoveryLog: {
      append: async (args) => {
        const snapshot = { ...args, id: "values" };
        saved.push(snapshot);
        return snapshot;
      },
    },
    captureTableState: async () => host.tableState(),
    resolveWorksheet: async (_context, { address }) => ({
      worksheet: host.worksheet,
      a1: address.slice(address.lastIndexOf("!") + 1),
    }),
  });
  vm.runInContext(source.slice(start, end), context);

  const args = { address: "A1:A2", sheet: "Sheet1", has_headers: false, name: "T" };
  const commit = await context.prepareCustomRecovery("excel_create_table", args, "create-table");
  await commit(await toolExcelCreateTable(args));
  assert.deepEqual(host.cells, ["Column1", "A", "B", "X"]);

  // Undo: the structural restore runs first, then the saved values.
  const restoreValues = () =>
    saved
      .find((snapshot) => snapshot.id === "values")
      .beforeValues.forEach(([value], row) => {
        host.cells[row] = value;
      });
  const tableState = saved.find((snapshot) => snapshot.id === "table").state;
  const redo = await context.restoreCustomState(tableState);
  restoreValues();
  assert.equal(host.tableExists(), false);
  assert.deepEqual(host.cells, ["A", "B", "X"]);

  // Redo recreates the table the way it was made; undo again restores the cells.
  const undoAgain = await context.restoreCustomState(redo);
  assert.equal(host.tableExists(), true);
  assert.deepEqual(host.cells, ["Column1", "A", "B", "X"]);
  await context.restoreCustomState(undoAgain);
  assert.deepEqual(host.cells, ["A", "B", "X"]);
});
