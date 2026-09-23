/* global Excel */
//
// Excel tool implementations + Excel-specific helpers.
//
// Each `toolExcel*` function is invoked from the dispatcher in taskpane.js
// when a matching `excel_*` tool_call arrives from the daemon. All operations
// go through Excel.run for proper context lifecycle. Address strings use
// A1 notation, optionally sheet-qualified (e.g. "Sheet1!A1:C5").

function _splitSheetAddress(address, fallbackSheetName) {
  // Returns { sheetName, a1 } from "Sheet1!A1:B2" or "A1:B2" (uses fallback).
  if (!address) return { sheetName: fallbackSheetName, a1: null };
  if (address.startsWith("'")) {
    for (let i = 1; i < address.length; i++) {
      if (address[i] !== "'") continue;
      if (address[i + 1] === "'") {
        i++;
        continue;
      }
      if (address[i + 1] === "!") {
        return {
          sheetName: address.slice(1, i).replaceAll("''", "'"),
          a1: address.slice(i + 2),
        };
      }
      break;
    }
  }
  const separator = address.indexOf("!");
  if (separator > 0) {
    return { sheetName: address.slice(0, separator), a1: address.slice(separator + 1) };
  }
  return { sheetName: fallbackSheetName, a1: address };
}

async function _activeSheetName(context) {
  const ws = context.workbook.worksheets.getActiveWorksheet();
  ws.load("name");
  await context.sync();
  return ws.name;
}

export async function toolExcelSetFrozenPanes({ sheetId, operation, dimension, count = 1 }) {
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items");
    await context.sync();
    for (const worksheet of worksheets.items) worksheet.load("id,name");
    await context.sync();
    const stableMap = Office?.context?.document?.settings?.get("openexcel-sheet-id-map") ?? {};
    const worksheet = worksheets.items.find(
      (item) => item.id === sheetId || Number(stableMap[item.id]) === Number(sheetId),
    );
    if (!worksheet) throw new Error(`Worksheet with ID ${sheetId} not found.`);
    if (operation === "unfreeze") {
      worksheet.freezePanes.unfreeze();
    } else {
      const location = worksheet.freezePanes.getLocationOrNullObject();
      location.load("isNullObject,rowCount,columnCount");
      await context.sync();
      const rows = dimension === "rows" ? count : location.isNullObject ? 0 : location.rowCount;
      const columns =
        dimension === "columns" ? count : location.isNullObject ? 0 : location.columnCount;
      if (rows && columns) {
        const letters = (value) =>
          value > 0
            ? letters(Math.floor((value - 1) / 26)) + String.fromCharCode(65 + ((value - 1) % 26))
            : "";
        worksheet.freezePanes.freezeAt(worksheet.getRange(`A1:${letters(columns)}${rows}`));
      } else if (rows) {
        worksheet.freezePanes.freezeRows(rows);
      } else {
        worksheet.freezePanes.freezeColumns(columns);
      }
    }
    await context.sync();
    return { success: true, operation, dimension, count, sheet: worksheet.name };
  });
}

export async function toolExcelGetSelectedRange({ cellLimit = 2000 } = {}) {
  if (!Number.isInteger(cellLimit) || cellLimit <= 0 || cellLimit > 5000) {
    throw new Error("`cellLimit` must be an integer from 1 to 5000.");
  }
  return await Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load("address, rowCount, columnCount, worksheet/name");
    await context.sync();
    const totalCellCount = range.rowCount * range.columnCount;
    const truncated = totalCellCount > cellLimit;
    const previewColumns = Math.min(range.columnCount, cellLimit);
    const previewRows = Math.min(
      range.rowCount,
      Math.max(1, Math.floor(cellLimit / previewColumns)),
    );
    const previewRange = truncated
      ? range.getCell(0, 0).getResizedRange(previewRows - 1, previewColumns - 1)
      : range;
    previewRange.load("address, values");
    await context.sync();
    return {
      address: range.address,
      sheet: range.worksheet.name,
      row_count: range.rowCount,
      column_count: range.columnCount,
      total_cell_count: totalCellCount,
      values: previewRange.values,
      preview_address: previewRange.address,
      truncated,
    };
  });
}

export async function toolExcelSelectRange({ address, sheet = null }) {
  if (!address || typeof address !== "string") {
    throw new Error("`address` (an A1 range like 'B4' or 'A1:D9') is required.");
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const ws = context.workbook.worksheets.getItem(activeName);
    ws.activate(); // make sure the selection is visible on the right sheet
    const range = ws.getRange(address);
    range.select();
    range.load("address");
    await context.sync();
    return { sheet: activeName, address: range.address, selected: true };
  });
}

// ---------------------------------------------------------------------------
// Tier 1 editing tools — formulas / formatting / columns / sheets / clear
// ---------------------------------------------------------------------------

// Tool: excel_set_format — number format / font / fill / borders on a range.
export async function toolExcelSetFormat({
  address,
  sheet = null,
  number_format,
  bold,
  italic,
  font_size,
  font_name,
  font_color,
  fill_color,
  border,
}) {
  if (!address || typeof address !== "string") {
    throw new Error("`address` (an A1 range) is required.");
  }
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    range.load("rowCount, columnCount, address");
    await context.sync();
    if (number_format !== undefined) {
      // numberFormat expects a 2D array matching the range shape.
      range.numberFormat = Array.from({ length: range.rowCount }, () =>
        Array.from({ length: range.columnCount }, () => number_format),
      );
    }
    if (bold !== undefined) range.format.font.bold = bold;
    if (italic !== undefined) range.format.font.italic = italic;
    if (font_size !== undefined) range.format.font.size = font_size;
    if (font_name !== undefined) range.format.font.name = font_name;
    if (font_color !== undefined) range.format.font.color = font_color;
    if (fill_color !== undefined) range.format.fill.color = fill_color;
    if (border) {
      for (const edge of [
        "EdgeTop",
        "EdgeBottom",
        "EdgeLeft",
        "EdgeRight",
        "InsideHorizontal",
        "InsideVertical",
      ]) {
        const b = range.format.borders.getItem(edge);
        b.style = "Continuous";
        b.weight = "Thin";
      }
    }
    await context.sync();
    return { sheet: sheetName, address: range.address, formatted: true };
  });
}

// ---------------------------------------------------------------------------
// Tier 2 editing tools — sort / filter / tables / charts / dimensions
// ---------------------------------------------------------------------------

// Tool: excel_sort_range — sort a range by one column.
export async function toolExcelSortRange({
  address,
  sheet = null,
  key = 0,
  ascending = true,
  has_headers = false,
}) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  if (!Number.isInteger(key) || key < 0) throw new Error("`key` must be a 0-based column index.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const range = ws.getRange(a1);
    range.sort.apply([{ key, ascending: !!ascending }], false, !!has_headers);
    range.load("address");
    await context.sync();
    return {
      sheet: sheetName,
      address: range.address,
      sorted_by_column: key,
      ascending: !!ascending,
    };
  });
}

// Tool: excel_autofilter — apply an AutoFilter to a range, or clear it.
export async function toolExcelAutoFilter({ address, sheet = null, clear = false }) {
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    if (clear) {
      ws.autoFilter.remove();
      await context.sync();
      return { sheet: sheetName, autofilter: "removed" };
    }
    if (!address || typeof address !== "string") {
      throw new Error("`address` is required (or pass clear: true to remove the filter).");
    }
    ws.autoFilter.apply(ws.getRange(a1));
    await context.sync();
    return { sheet: sheetName, autofilter: "applied", address: a1 };
  });
}

// Tool: excel_create_table — turn a range into a named Excel table.
export async function toolExcelCreateTable({
  address,
  sheet = null,
  has_headers = true,
  name = null,
}) {
  if (!address || typeof address !== "string") throw new Error("`address` is required.");
  return await Excel.run(async (context) => {
    const activeName = sheet || (await _activeSheetName(context));
    const { sheetName, a1 } = _splitSheetAddress(address, activeName);
    const ws = context.workbook.worksheets.getItem(sheetName);
    const table = ws.tables.add(a1, !!has_headers);
    if (name) table.name = name;
    table.load("name");
    await context.sync();
    return { sheet: sheetName, table: table.name, range: a1 };
  });
}

// Tool: excel_add_table_rows — append rows to an existing table.
export async function toolExcelAddTableRows({ table, values, index = null }) {
  if (!table || typeof table !== "string") throw new Error("`table` (table name) is required.");
  if (!Array.isArray(values) || !values.length || !Array.isArray(values[0]) || !values[0].length) {
    throw new Error("`values` must be a non-empty rectangular 2D array.");
  }
  const width = values[0].length;
  if (values.some((row) => !Array.isArray(row) || row.length !== width)) {
    throw new Error("`values` must be rectangular; every row must have the same length.");
  }
  if (index !== null && (!Number.isInteger(index) || index < 0)) {
    throw new Error("`index` must be a non-negative integer or null.");
  }
  return await Excel.run(async (context) => {
    const t = context.workbook.tables.getItem(table);
    t.rows.add(Number.isInteger(index) ? index : null, values);
    await context.sync();
    return { table, added_rows: values.length };
  });
}

const CHART_TYPE = {
  column: "ColumnClustered",
  bar: "BarClustered",
  line: "Line",
  pie: "Pie",
  scatter: "XYScatter",
  area: "Area",
};
