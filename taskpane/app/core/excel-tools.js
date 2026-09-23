/* global Excel */
// The Office.js side of each tool the daemon can call, plus the reads the pane
// makes on its own (selection, per-turn context) and cell-link navigation.
import {
  toolExcelGetSelectedRange,
  toolExcelSelectRange,
  toolExcelSetFormat,
  toolExcelSortRange,
  toolExcelAutoFilter,
  toolExcelCreateTable,
  toolExcelAddTableRows,
  toolExcelSetFrozenPanes,
} from "../../shared/tools-excel.js";
import {
  clearCellRange,
  copyTo,
  getAllObjects,
  getCellRanges,
  getRangeAsCsv,
  getWorkbookMetadata,
  modifyObject,
  modifySheetStructure,
  modifyWorkbookStructure,
  resizeRange,
  searchData,
  setCellRange,
} from "../../shared/vendor/office-agents-excel-api.js";
import {
  buildOverview,
  ChangeTracker,
  createExplainFormulaTool,
  createTraceDependenciesTool,
  readSelectionContext,
} from "../../shared/vendor/pi-context.js";
import { workbookHistory } from "../../shared/recovery.js";

export { getWorkbookMetadata };

// pi-for-excel's read-only formula tools, used as-is. They report failures in
// their text output rather than throwing, so the text is the whole result.
const explainFormulaTool = createExplainFormulaTool();
const traceDependenciesTool = createTraceDependenciesTool();

async function runPiTool(piTool, toolCallId, args) {
  const output = await piTool.execute(toolCallId, args ?? {});
  return { text: output.content.map((part) => part.text ?? "").join("\n") };
}

// ---- Per-turn workbook context ---------------------------------------------
// Read with pi-for-excel's overview, selection and change-tracker readers. Each
// part is optional: a slow or failing read is dropped rather than delaying the turn.
function limitContextText(value, maxLength, preserveTail = false) {
  if (!value || value.length <= maxLength) return value;
  const notice = "\n[Context truncated; read the relevant range with an Excel tool.]\n";
  if (!preserveTail) return value.slice(0, maxLength - notice.length) + notice;
  const tailLength = Math.min(2000, Math.floor(maxLength / 4));
  return value.slice(0, maxLength - tailLength - notice.length) + notice + value.slice(-tailLength);
}

function settleWithin(promise, ms) {
  let timer;
  const timeout = new Promise((resolve) => {
    timer = setTimeout(() => resolve(null), ms);
  });
  return Promise.race([promise.catch(() => null), timeout]).finally(() => clearTimeout(timer));
}

export function createContextSnapshot({
  getMetadata = getWorkbookMetadata,
  overview = buildOverview,
  readSelection = readSelectionContext,
  changeTracker,
}) {
  let overviewRead = null;
  // Concurrent turns share one unfinished overview read.
  function readOverviewSingleFlight() {
    if (!overviewRead) {
      const read = Promise.resolve().then(overview);
      overviewRead = read;
      const clear = () => {
        if (overviewRead === read) overviewRead = null;
      };
      read.then(clear, clear);
    }
    return overviewRead;
  }

  return async function contextSnapshot({ selectionAddress = null } = {}) {
    const [metadata, overviewText, selection] = await Promise.all([
      settleWithin(getMetadata(), 1500),
      settleWithin(readOverviewSingleFlight(), 2500),
      selectionAddress ? settleWithin(readSelection(selectionAddress), 1500) : null,
    ]);
    const sheetIds = (metadata?.sheetsMetadata ?? []).map((sheet) => `${sheet.name}=${sheet.id}`);
    const workbookParts = [];
    if (sheetIds.length) {
      workbookParts.push(
        limitContextText(`sheetId for mcp__office__excel_* tools: ${sheetIds.join(", ")}`, 4000),
      );
    }
    if (overviewText) workbookParts.push(limitContextText(overviewText, 8000));
    return {
      workbook: workbookParts.join("\n\n") || null,
      selection: limitContextText(selection?.text ?? null, 12000, true),
      changes: limitContextText(changeTracker.flush(), 3000),
    };
  };
}

export const changeTracker = new ChangeTracker();
const contextSnapshot = createContextSnapshot({ changeTracker });

// ---- Tool dispatch ----------------------------------------------------------
export async function executeTool(name, args, id) {
  switch (name) {
    case "excel_get_selected_range":
      return toolExcelGetSelectedRange(args);
    case "excel_get_workbook_metadata":
      return getWorkbookMetadata();
    case "excel_context_snapshot":
      return contextSnapshot(args);
    case "excel_get_cell_ranges":
      return getCellRanges(args.sheetId, args.ranges, {
        includeStyles: args.includeStyles,
        cellLimit: args.cellLimit,
      });
    case "excel_get_range_as_csv":
      return getRangeAsCsv(args.sheetId, args.range, {
        includeHeaders: args.includeHeaders,
        maxRows: args.maxRows,
      });
    case "excel_search_data":
      return searchData(args.searchTerm, {
        sheetId: args.sheetId,
        range: args.range,
        offset: args.offset,
        cursor: args.cursor,
        ...args.options,
      });
    case "excel_get_all_objects":
      return getAllObjects({ sheetId: args.sheetId, id: args.id });
    case "excel_explain_formula":
      return runPiTool(explainFormulaTool, id, args);
    case "excel_trace_dependencies":
      return runPiTool(traceDependenciesTool, id, args);
    case "excel_set_cell_range":
      return setCellRange(args.sheetId, args.range, args.cells, {
        copyToRange: args.copyToRange,
        resizeWidth: args.resizeWidth,
        resizeHeight: args.resizeHeight,
        allowOverwrite: args.allow_overwrite,
      });
    case "excel_clear_cell_range":
      return clearCellRange(args.sheetId, args.range, args.clearType);
    case "excel_copy_to":
      return copyTo(args.sheetId, args.sourceRange, args.destinationRange, args.allow_overwrite);
    case "excel_modify_sheet_structure":
      return ["freeze", "unfreeze"].includes(args.operation)
        ? toolExcelSetFrozenPanes(args)
        : modifySheetStructure(args.sheetId, {
            operation: args.operation,
            dimension: args.dimension,
            reference: args.reference,
            count: args.count,
            position: args.position,
          });
    case "excel_modify_workbook_structure":
      return modifyWorkbookStructure({
        operation: args.operation,
        sheetId: args.sheetId,
        sheetName: args.sheetName,
        newName: args.newName,
        tabColor: args.tabColor,
      });
    case "excel_resize_range":
      return resizeRange(args.sheetId, {
        range: args.range,
        width: args.width,
        height: args.height,
      });
    case "excel_modify_object":
      return modifyObject({
        operation: args.operation,
        sheetId: args.sheetId,
        objectType: args.objectType,
        id: args.id,
        properties: args.properties,
      });
    case "excel_select_range":
      return toolExcelSelectRange(args);
    case "excel_set_format":
      return toolExcelSetFormat(args);
    case "excel_sort_range":
      return toolExcelSortRange(args);
    case "excel_autofilter":
      return toolExcelAutoFilter(args);
    case "excel_create_table":
      return toolExcelCreateTable(args);
    case "excel_add_table_rows":
      return toolExcelAddTableRows(args);
    case "excel_workbook_history":
      return workbookHistory(args);
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}

// ---- Selection ----------------------------------------------------------------
// One cell shows its value; a block shows its address and size.
export async function readSelection(excel = globalThis.Excel) {
  return excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    const firstCell = range.getCell(0, 0);
    range.load("address, rowCount, columnCount");
    firstCell.load("values");
    await context.sync();
    const cellCount = (range.rowCount || 0) * (range.columnCount || 0);
    const text =
      cellCount === 1
        ? String(firstCell.values?.[0]?.[0] ?? "")
        : `${range.address} (${range.rowCount}×${range.columnCount})`;
    return { text, address: range.address };
  });
}

export function watchSelection(handler) {
  return Excel.run(async (context) => {
    context.workbook.onSelectionChanged.add(handler);
    await context.sync();
  });
}

// ---- Cell links -------------------------------------------------------------
// Ported from pi-for-excel's cell-link.ts. It navigates with a native selection
// (activate the sheet, select the range): the viewport scrolls and Excel draws
// its own highlight, so nothing is written and no undo state is disturbed.
export async function navigateToRange(address) {
  // Disjoint areas cannot be scrolled to at once; go to the first.
  const first = String(address).split(",")[0].trim();
  const bang = first.lastIndexOf("!");
  const sheet = bang > 0 ? first.slice(0, bang).replace(/^'|'$/g, "").replaceAll("''", "'") : null;
  const local = bang > 0 ? first.slice(bang + 1) : first;
  await Excel.run(async (context) => {
    const worksheet = sheet
      ? context.workbook.worksheets.getItem(sheet)
      : context.workbook.worksheets.getActiveWorksheet();
    worksheet.activate();
    await context.sync();
    worksheet.getRange(local).select();
    await context.sync();
  });
}
