import { createSdkMcpServer, tool } from "@anthropic-ai/claude-agent-sdk";
import { z } from "zod";
import { diag } from "./diag.mjs";
import { COMPUTE_TOOL_DESCRIPTION, createComputeShell } from "./compute-tool.mjs";

// Wrap a bridge tool result for MCP. Handlers return {content: [...]}.
function asMcpResult(result, { isError = false } = {}) {
  const text = typeof result === "string" ? result : JSON.stringify(result, null, 2);
  return { content: [{ type: "text", text }], ...(isError ? { isError: true } : {}) };
}

function toCellInput(cell) {
  if (cell !== null && typeof cell === "object") return cell;
  return typeof cell === "string" && cell.startsWith("=") ? { formula: cell } : { value: cell };
}

function columnNumber(label) {
  return [...label.toUpperCase()].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
}

function parseA1RangeSize(address) {
  const local = String(address ?? "").split("!").pop().replaceAll("$", "").trim();
  const match = /^([A-Z]+)([1-9]\d*)(?::([A-Z]+)([1-9]\d*))?$/i.exec(local);
  if (!match) return null;
  const startColumn = columnNumber(match[1]);
  const startRow = Number(match[2]);
  const endColumn = columnNumber(match[3] ?? match[1]);
  const endRow = Number(match[4] ?? match[2]);
  if (endColumn < startColumn || endRow < startRow) return null;
  return {
    rows: endRow - startRow + 1,
    columns: endColumn - startColumn + 1,
    start: `${match[1].toUpperCase()}${startRow}`,
    startColumn,
    startRow,
  };
}

function columnLetters(number) {
  let result = "";
  for (let value = number; value > 0; value = Math.floor((value - 1) / 26)) {
    result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
  }
  return result;
}

function parseCellsPayload(raw) {
  if (typeof raw !== "string") return raw;
  const text = raw.trim();
  if (!(text.startsWith("[") || text.startsWith("{"))) return raw;
  try {
    return JSON.parse(text);
  } catch (error) {
    throw new Error(`cells is not valid JSON (${error.message}); send cells as an array, not a string`);
  }
}

function normalizeRanges(raw) {
  if (Array.isArray(raw)) return raw;
  const text = String(raw ?? "").trim();
  if (!text) return [];
  try {
    const parsed = JSON.parse(text);
    if (Array.isArray(parsed)) return parsed;
  } catch {}
  const unwrapped = text.replace(/^\s*\[/, "").replace(/\]\s*$/, "");
  return unwrapped
    .split(",")
    .map((part) => part.trim().replace(/^['"]|['"]$/g, ""))
    .filter(Boolean);
}

function normalizeCellMatrix(raw, range) {
  const parsed = parseCellsPayload(raw);
  if (!Array.isArray(parsed)) return [[toCellInput(parsed)]];
  if (parsed.length === 0) throw new Error("cells must not be empty");

  const nested = parsed.map(Array.isArray);
  if (nested.every(Boolean)) return parsed.map((row) => row.map(toCellInput));
  if (nested.some(Boolean)) throw new Error("cells cannot mix rows with individual cell values");

  // A flat list is a single row, or a single column when the target is one column wide.
  // For a multi-row, multi-column target the intended layout is ambiguous, so ask for rows.
  const size = parseA1RangeSize(range);
  if (size && size.columns === 1 && size.rows > 1) return parsed.map((cell) => [toCellInput(cell)]);
  if (size && size.columns > 1 && size.rows > 1 && parsed.length > 1) {
    throw new Error(`cells is a flat list but ${range} spans several rows and columns; send a 2D array of rows`);
  }
  return [parsed.map(toCellInput)];
}

function prepareCellWrite(args, cellMatrix) {
  const size = parseA1RangeSize(args.range);
  const width = cellMatrix[0]?.length ?? 0;
  // A formula pattern aimed at a larger explicit range means "fill it through
  // the range". Preserve Excel's relative-reference translation by writing the
  // pattern and using copyToRange. Values are never repeated.
  const height = cellMatrix.length;
  if (
    !args.copyToRange &&
    cellMatrix.every((row) => row.every((cell) => typeof cell.formula === "string")) &&
    size &&
    (height < size.rows || width < size.columns) &&
    size.rows % height === 0 &&
    size.columns % width === 0
  ) {
    const patternRange =
      height === 1 && width === 1
        ? size.start
        : `${size.start}:${columnLetters(size.startColumn + width - 1)}${size.startRow + height - 1}`;
    return { ...args, range: patternRange, copyToRange: args.range, cells: cellMatrix };
  }
  // The task pane resizes the range to the matrix. For an explicit multi-cell
  // target, growing past it writes cells the model never named (e.g. a 3x1
  // column sent for the row J3:L3 lands in J3:J5), so refuse instead.
  if (size && size.rows * size.columns > 1 && (height > size.rows || width > size.columns)) {
    throw new Error(
      `cells is ${height}x${width} but range ${args.range} is ${size.rows}x${size.columns}; ` +
        "the write would spill outside range. Reshape cells (outer array = rows) or fix range.",
    );
  }
  return { ...args, cells: cellMatrix };
}

// Write receipts read back the whole target, so a 5000-row fill would return
// every formula value and overflow the tool-output limit. Keep a sample of
// results and the first errors, with totals, so the receipt stays readable.
const MAX_FORMULA_RESULTS = 20;
const MAX_FORMULA_ERRORS = 50;

function boundWriteReceipt(result) {
  if (!result || typeof result !== "object") return result;
  const bounded = { ...result };
  const entries = Object.entries(result.formulaResults ?? {});
  if (entries.length > MAX_FORMULA_RESULTS) {
    bounded.formulaResults = Object.fromEntries(entries.slice(0, MAX_FORMULA_RESULTS));
    bounded.formulaResultCount = entries.length;
  }
  if (Array.isArray(result.formulaErrors) && result.formulaErrors.length > MAX_FORMULA_ERRORS) {
    bounded.formulaErrors = result.formulaErrors.slice(0, MAX_FORMULA_ERRORS);
    bounded.formulaErrorCount = result.formulaErrors.length;
  }
  return bounded;
}

function asMcpError(err) {
  return asMcpResult(
    {
      success: false,
      error: err?.message ?? String(err),
      ...(err?.code ? { code: err.code } : {}),
      ...(err?.commitStatus ? { commitStatus: err.commitStatus } : {}),
    },
    { isError: true },
  );
}

/**
 * Build the office-bridge MCP server. Tools forward calls over the WS bridge
 * to the Excel task pane, which executes them in Office.js and returns results.
 *
 * @param {{ callTaskpaneTool: (name: string, args: object) => Promise<any> }} bridge
 * @param {"excel"|null} host
 */
export function createOfficeBridgeMcp(
  bridge,
  host = null,
  paneKey = null,
  { signal } = {},
) {
  // `paneKey` routes every call to the exact workbook pane this session
  // belongs to (so two open workbooks don't cross-talk).
  const call = (name, args, options = {}) =>
    bridge.callTaskpaneTool(name, args ?? {}, paneKey, { signal: options.signal ?? signal });

  const excel_get_selected_range = tool(
    "excel_get_selected_range",
    "Return the user's current selection in the active Excel workbook. Includes its full address and dimensions plus a bounded top-left values preview. When truncated is true, use mcp__office__excel_get_cell_ranges to read the specific rows or columns needed. Call whenever the user refers to 'this', 'these cells', 'the selection', or asks to edit existing content without specifying location.",
    {
      cellLimit: z
        .number()
        .int()
        .positive()
        .max(5000)
        .optional()
        .describe("Maximum cells in the values preview. Default: 2000; maximum: 5000."),
    },
    async (args) => {
      try {
        return asMcpResult(await call("excel_get_selected_range", args));
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const excel_select_range = tool(
    "excel_select_range",
    "Select a cell or range in the spreadsheet, making it the user's active selection (and switching to its sheet). Use this when the user asks to 'select', 'highlight', 'go to', or 'jump to' a cell/range/result.",
    {
      address: z.string().describe("A1-style cell or range to select, e.g. 'B4' or 'A1:D9'."),
      sheet: z.string().optional().describe("Worksheet name; defaults to the active sheet."),
    },
    async (args) => {
      try {
        return asMcpResult(await call("excel_select_range", args));
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const wrap = (name) => async (args) => {
    try {
      return asMcpResult(boundWriteReceipt(await call(name, args ?? {})));
    } catch (e) {
      return asMcpError(e);
    }
  };

  const scalarCell = z.union([z.string(), z.number(), z.boolean(), z.null()]);
  const tableValues = z
    .array(z.array(scalarCell).min(1))
    .min(1)
    .superRefine((rows, ctx) => {
      const width = rows[0]?.length;
      rows.forEach((row, index) => {
        if (row.length !== width) {
          ctx.addIssue({
            code: "custom",
            path: [index],
            message: `Row ${index + 1} has ${row.length} cells; expected ${width}.`,
          });
        }
      });
    });

  const excel_set_format = tool(
    "excel_set_format",
    "Set number format and/or font/fill/border styling on a range.",
    {
      address: z.string().describe("A1 range to format."),
      sheet: z.string().optional(),
      number_format: z.string().optional().describe("e.g. '0.00', '$#,##0', 'yyyy-mm-dd', '0%'."),
      bold: z.boolean().optional(),
      italic: z.boolean().optional(),
      font_size: z.number().optional(),
      font_name: z.string().optional(),
      font_color: z.string().optional().describe("Hex color like '#1F4E79'."),
      fill_color: z.string().optional().describe("Cell fill hex color."),
      border: z.boolean().optional().describe("true = thin continuous borders on all edges."),
    },
    wrap("excel_set_format"),
  );

  const excel_sort_range = tool(
    "excel_sort_range",
    "Sort a range by one column (0-based index within the range).",
    {
      address: z.string().describe("A1 range to sort."),
      sheet: z.string().optional(),
      key: z.number().int().optional().describe("0-based column index within the range."),
      ascending: z.boolean().optional().describe("Default true."),
      has_headers: z.boolean().optional().describe("Treat the first row as headers."),
    },
    wrap("excel_sort_range"),
  );

  const excel_autofilter = tool(
    "excel_autofilter",
    "Apply an AutoFilter to a range, or pass clear:true to remove the sheet's filter.",
    {
      address: z.string().min(1).optional().describe("A1 range (required unless clear)."),
      sheet: z.string().min(1).optional(),
      clear: z.boolean().optional(),
    },
    wrap("excel_autofilter"),
  );

  const excel_create_table = tool(
    "excel_create_table",
    "Turn a range into a named Excel table (ListObject).",
    {
      address: z.string(),
      sheet: z.string().optional(),
      has_headers: z.boolean().optional().describe("Default true."),
      name: z.string().optional(),
    },
    wrap("excel_create_table"),
  );

  const excel_add_table_rows = tool(
    "excel_add_table_rows",
    "Append rows to an existing table by name.",
    {
      table: z.string().min(1).describe("Table name."),
      values: tableValues.describe("Non-empty rectangular 2D array of row values."),
      index: z.number().int().nonnegative().optional().describe("Insert position; omit to append."),
    },
    wrap("excel_add_table_rows"),
  );

  const excel_workbook_history = tool(
    "excel_workbook_history",
    "List or restore automatic recovery checkpoints for this workbook. Use restore only when the user asks to undo or recover a prior assistant edit. Delete removes one checkpoint; clear removes this workbook's checkpoints.",
    {
      action: z.enum(["list", "restore", "delete", "clear"]).default("list"),
      snapshot_id: z.string().optional().describe("Checkpoint id. Restore uses the latest checkpoint when omitted."),
      limit: z.number().int().min(1).max(50).optional().describe("Maximum checkpoints returned by list."),
    },
    wrap("excel_workbook_history"),
  );

  // ---- Excel tools ported from hewliyang/office-agents (MIT) ----
  // Schemas and descriptions mirror office-agents' packages/excel/src/lib/tools;
  // the task pane executes them via taskpane/shared/vendor/office-agents-excel-api.js.
  const sheetId = z
    .number()
    .int()
    .describe(
      "Worksheet ID from mcp__office__excel_get_workbook_metadata (stable per workbook, not the tab position).",
    );
  const explanation = z.string().optional().describe("Brief explanation (a few words).");
  const borderSide = z
    .object({
      style: z.enum(["solid", "dashed", "dotted", "double"]).optional(),
      weight: z.enum(["thin", "medium", "thick"]).optional(),
      color: z.string().optional(),
    })
    .optional();
  const cellInput = z.object({
    value: z.union([z.string(), z.number(), z.boolean(), z.null()]).optional(),
    formula: z
      .string()
      .startsWith("=", "Formula must start with '='.")
      .optional()
      .describe("Formula starting with '=', e.g. '=SUM(B2:B5)'."),
    note: z.string().optional(),
    cellStyles: z
      .object({
        fontWeight: z.enum(["normal", "bold"]).optional(),
        fontStyle: z.enum(["normal", "italic"]).optional(),
        fontLine: z.enum(["none", "underline", "line-through"]).optional(),
        fontSize: z.number().optional(),
        fontFamily: z.string().optional(),
        fontColor: z.string().optional(),
        backgroundColor: z.string().optional(),
        horizontalAlignment: z.enum(["left", "center", "right"]).optional(),
        numberFormat: z.string().optional(),
      })
      .optional(),
    borderStyles: z
      .object({ top: borderSide, bottom: borderSide, left: borderSide, right: borderSide })
      .optional(),
  });
  const cellMatrix = z
    .array(z.array(z.union([cellInput, scalarCell])).min(1))
    .min(1)
    .superRefine((rows, ctx) => {
      const width = rows[0]?.length;
      rows.forEach((row, index) => {
        if (row.length !== width) {
          ctx.addIssue({
            code: "custom",
            path: [index],
            message: `Row ${index + 1} has ${row.length} cells; expected ${width}.`,
          });
        }
      });
    });
  const cellPayload = z
    .union([
      cellMatrix,
      z.array(z.union([cellInput, scalarCell])).min(1),
      cellInput,
      scalarCell,
    ])
    .describe(
      "Cell data as a JSON array (not a string): a rectangular 2D array of rows, a 1D row/column, or one cell.",
    );
  const rangesPayload = z.union([
    z.array(z.string().min(1)).min(1),
    z.string().min(1),
  ]);
  const size = z
    .object({ type: z.enum(["points", "standard"]), value: z.number().positive() })
    .optional();

  const excel_get_workbook_metadata = tool(
    "excel_get_workbook_metadata",
    "READ. Workbook blueprint: file name, every sheet with its stable sheetId, name, used size and frozen panes, plus the active sheet and current selection. Call this first to get the sheetId values the other mcp__office__excel_* tools require.",
    {},
    wrap("excel_get_workbook_metadata"),
  );

  const excel_get_cell_ranges = tool(
    "excel_get_cell_ranges",
    "READ. Read cell values and formulas (plus formatting with includeStyles) as a sparse A1-keyed object. Each call scans at most 20000 cells in bounded chunks. If hasMore is true, pass remainingRanges as ranges in the next call with the same sheetId and options; unread ranges may contain blanks. Use this to inspect data before modifying it.",
    {
      sheetId,
      ranges: rangesPayload.describe(
        "One range or an array of ranges in A1 notation, e.g. 'A1:C10' or ['A1:C10', 'E1:E100'].",
      ),
      includeStyles: z
        .boolean()
        .optional()
        .describe("Include font/fill styling info. Default: false."),
      cellLimit: z
        .number()
        .int()
        .positive()
        .max(20_000)
        .optional()
        .describe("Maximum populated cells to return. Default: 2000."),
      explanation,
    },
    (args) =>
      wrap("excel_get_cell_ranges")({
        ...args,
        ranges: normalizeRanges(args.ranges),
        includeStyles: args.includeStyles ?? false,
      }),
  );

  const excel_get_range_as_csv = tool(
    "excel_get_range_as_csv",
    "READ. Read a bounded page of cell data as CSV, at most 20000 cells. If hasMore is true, continue with nextRange as range and includeHeaders=true so the next data row is not skipped. Use for tabular analysis without styling info.",
    {
      sheetId,
      range: z.string().describe("Range in A1 notation, e.g. 'A1:Z100'."),
      includeHeaders: z
        .boolean()
        .optional()
        .describe("Include first row as headers. Default: true."),
      maxRows: z
        .number()
        .int()
        .positive()
        .max(20_000)
        .optional()
        .describe("Maximum rows to return, also bounded by 20000 cells. Default: 500."),
      explanation,
    },
    wrap("excel_get_range_as_csv"),
  );

  const excel_search_data = tool(
    "excel_search_data",
    "READ. Find text, values or formulas in bounded chunks, scanning at most 20000 cells per call. Supports regex and case-sensitive search. If hasMore is true, pass nextCursor unchanged as cursor with the same search arguments, even when matches is empty. hasMore means unscanned cells remain, not guaranteed hits. totalFound is cumulative and exact only when totalFoundIsExact is true. Search reads live data; restart after workbook structure changes.",
    {
      searchTerm: z.string().describe("The text or pattern to search for."),
      sheetId: sheetId.optional().describe("Limit to a specific sheet."),
      range: z.string().optional().describe("Limit search scope, e.g. 'A1:Z100'."),
      offset: z
        .number()
        .int()
        .nonnegative()
        .optional()
        .describe("Initial matching results to skip. Default: 0. Use cursor for continuation."),
      cursor: z
        .string()
        .optional()
        .describe(
          "Opaque nextCursor from the previous result. Keep all search arguments unchanged.",
        ),
      options: z
        .object({
          matchCase: z.boolean().optional(),
          matchEntireCell: z.boolean().optional(),
          matchFormulas: z.boolean().optional(),
          useRegex: z.boolean().optional(),
          maxResults: z.number().int().positive().max(5000).optional(),
        })
        .optional(),
      explanation,
    },
    wrap("excel_search_data"),
  );

  const excel_get_all_objects = tool(
    "excel_get_all_objects",
    "READ. List all charts, pivot tables, and other objects in the workbook. Use this to discover what exists before modifying it.",
    {
      sheetId: sheetId.optional().describe("Filter to a specific sheet."),
      id: z.string().optional().describe("Filter by object ID."),
      explanation,
    },
    wrap("excel_get_all_objects"),
  );

  // Ported from pi-for-excel's explain_formula / trace_dependencies (MIT); the
  // task pane runs Pi's implementations, these mirror their schemas.
  const formulaCell = z
    .string()
    .min(1)
    .describe('Single cell, e.g. "D10" or "Sheet2!F5" (sheet name, not sheetId). Not a range.');

  const excel_explain_formula = tool(
    "excel_explain_formula",
    "READ. Explain what a formula cell does in plain language, with its direct input references and their current values. Use before editing or debugging a formula you didn't write.",
    {
      cell: formulaCell,
      max_references: z
        .number()
        .int()
        .min(1)
        .max(20)
        .optional()
        .describe("Max direct references to preview. Default: 8."),
    },
    wrap("excel_explain_formula"),
  );

  const excel_trace_dependencies = tool(
    "excel_trace_dependencies",
    "READ. Trace formula lineage for a cell: precedents (what feeds it) or dependents (what it feeds, i.e. what breaks if it changes), recursively up to depth 5.",
    {
      cell: formulaCell,
      mode: z.enum(["precedents", "dependents"]).optional().describe("Default: precedents."),
      depth: z.number().int().min(1).max(5).optional().describe("Levels to trace. Default: 2."),
    },
    wrap("excel_trace_dependencies"),
  );

  const excel_set_cell_range = tool(
    "excel_set_cell_range",
    "WRITE. Write values, formulas, and formatting to cells. Accepts 2D matrices, a single cell, or a 1D list (a row, or a column when range is one column wide). A formula pattern smaller than range is filled across it with relative-reference translation; values are written once from the top-left cell. Computed formula values and errors come back for verification. OVERWRITE PROTECTION: use allow_overwrite=true immediately when the user's requested edit targets existing cells; ask only when the overwrite is outside the requested scope. Use copyToRange to expand larger patterns.",
    {
      sheetId,
      range: z.string().describe("Target range in A1 notation (auto-expands to match cells)."),
      cells: cellPayload,
      copyToRange: z
        .string()
        .optional()
        .describe("Expand the written pattern to this larger range."),
      resizeWidth: size,
      resizeHeight: size,
      allow_overwrite: z
        .boolean()
        .optional()
        .describe("Set true when the user's requested edit includes replacing data in the target range."),
      explanation,
    },
    (args) => {
      try {
        const cells = cellMatrix.parse(normalizeCellMatrix(args.cells, args.range));
        return wrap("excel_set_cell_range")(prepareCellWrite(args, cells));
      } catch (error) {
        return asMcpError(error);
      }
    },
  );

  const excel_fill_formula = tool(
    "excel_fill_formula",
    "WRITE. Fill one Excel formula through a target range using relative-reference translation. Give the target range and one formula; do not build a formula matrix. Returns the computed results and formula errors for verification. Set allow_overwrite=true when the user's requested edit includes replacing existing cells in that range.",
    {
      sheetId,
      range: z.string().min(1).describe("Full target range in A1 notation, e.g. 'F2:F5000'."),
      formula: z.string().min(2).startsWith("=").describe("Formula for the top-left cell, e.g. '=SUM(B2:E2)'."),
      allow_overwrite: z.boolean().optional(),
      explanation,
    },
    (args) => {
      try {
        if (!parseA1RangeSize(args.range)) {
          throw new Error("range must be a valid A1 cell or rectangular range");
        }
        const { formula, ...writeArgs } = args;
        return wrap("excel_set_cell_range")(
          prepareCellWrite(writeArgs, [[{ formula }]]),
        );
      } catch (error) {
        return asMcpError(error);
      }
    },
  );

  const excel_clear_cell_range = tool(
    "excel_clear_cell_range",
    "WRITE. Clear contents, formatting, or both from a range. 'contents' keeps formatting, 'formats' keeps values, 'all' clears everything.",
    {
      sheetId,
      range: z.string().describe("Range to clear in A1 notation."),
      clearType: z.enum(["contents", "all", "formats"]).optional().describe("Default: 'contents'."),
      explanation,
    },
    wrap("excel_clear_cell_range"),
  );

  const excel_copy_to = tool(
    "excel_copy_to",
    "WRITE. Copy a range to another location with formula translation. If the destination is larger, the source pattern repeats. Set allow_overwrite=true when the requested edit targets existing destination cells; ask only when replacement falls outside the user's requested scope.",
    {
      sheetId,
      sourceRange: z.string().describe("Source range in A1 notation."),
      destinationRange: z.string().describe("Destination range in A1 notation."),
      allow_overwrite: z
        .boolean()
        .optional()
        .describe("Set true when the user's requested edit includes replacing destination data."),
      explanation,
    },
    wrap("excel_copy_to"),
  );

  const excel_modify_sheet_structure = tool(
    "excel_modify_sheet_structure",
    "WRITE. Insert, delete, hide, unhide, freeze, or unfreeze rows and columns. Use a reference like '5' for row 5 or 'C' for column C.",
    {
      sheetId,
      operation: z.enum(["insert", "delete", "hide", "unhide", "freeze", "unfreeze"]),
      dimension: z.enum(["rows", "columns"]),
      reference: z.string().optional().describe("Row number or column letter, e.g. '5' or 'C'."),
      count: z.number().int().positive().optional().describe("Number of rows/columns. Default: 1."),
      position: z
        .enum(["before", "after"])
        .optional()
        .describe("Insert before or after reference. Default: 'before'."),
      explanation,
    },
    wrap("excel_modify_sheet_structure"),
  );

  const excel_modify_workbook_structure = tool(
    "excel_modify_workbook_structure",
    "WRITE. Create, delete, rename, or duplicate worksheets.",
    {
      operation: z.enum(["create", "delete", "rename", "duplicate"]),
      sheetId: sheetId.optional().describe("Sheet ID for delete/rename/duplicate."),
      sheetName: z.string().optional().describe("Name for a new sheet (create)."),
      newName: z
        .string()
        .optional()
        .describe("New name (rename) or name for the copy (duplicate)."),
      tabColor: z.string().optional().describe("Tab color as hex, e.g. '#ff0000'."),
      explanation,
    },
    wrap("excel_modify_workbook_structure"),
  );

  const excel_resize_range = tool(
    "excel_resize_range",
    "WRITE. Adjust column widths or row heights. Use 'A:D' for columns A-D, '1:5' for rows 1-5, or omit range for the entire sheet.",
    {
      sheetId,
      range: z
        .string()
        .optional()
        .describe("Column range (A:D) or row range (1:5). Omit for the entire sheet."),
      width: size,
      height: size,
      explanation,
    },
    wrap("excel_resize_range"),
  );

  const pivotField = z.object({
    field: z.string(),
    summarizeBy: z.enum(["sum", "count", "average", "max", "min"]).optional(),
  });
  const excel_modify_object = tool(
    "excel_modify_object",
    "WRITE. Create, update, or delete charts and pivot tables. For charts, specify chartType, source, and anchor. For pivot tables, specify source, range, rows, columns, and values.",
    {
      operation: z.enum(["create", "update", "delete"]),
      sheetId,
      objectType: z.enum(["pivotTable", "chart"]),
      id: z.string().optional().describe("Object ID (required for update/delete)."),
      properties: z
        .object({
          name: z.string().optional(),
          source: z.string().optional().describe("Data source range, e.g. 'Sheet1!A1:D100'."),
          range: z.string().optional().describe("Output location (pivot table top-left cell)."),
          anchor: z.string().optional().describe("Chart placement (top-left cell)."),
          rows: z.array(z.object({ field: z.string() })).optional(),
          columns: z.array(z.object({ field: z.string() })).optional(),
          values: z.array(pivotField).optional(),
          title: z.string().optional(),
          chartType: z
            .enum(["columnClustered", "barClustered", "line", "pie", "scatter", "area", "doughnut"])
            .optional(),
        })
        .optional(),
      explanation,
    },
    wrap("excel_modify_object"),
  );

  let computeShell = null;
  const excel_bash = tool(
    "excel_bash",
    COMPUTE_TOOL_DESCRIPTION,
    { command: z.string().min(1).describe("The bash command to run in the sandbox.") },
    async (args) => {
      try {
        computeShell ??= createComputeShell(call, { signal });
        const result = await computeShell(args);
        return asMcpResult(result, { isError: result.exitCode !== 0 });
      } catch (e) {
        return asMcpError(e);
      }
    },
  );

  const excelTools = [
    excel_get_workbook_metadata,
    excel_get_selected_range,
    excel_get_cell_ranges,
    excel_get_range_as_csv,
    excel_search_data,
    excel_get_all_objects,
    excel_explain_formula,
    excel_trace_dependencies,
    excel_set_cell_range,
    excel_fill_formula,
    excel_clear_cell_range,
    excel_copy_to,
    excel_modify_sheet_structure,
    excel_modify_workbook_structure,
    excel_resize_range,
    excel_modify_object,
    excel_select_range,
    excel_set_format,
    excel_sort_range,
    excel_autofilter,
    excel_create_table,
    excel_add_table_rows,
    excel_workbook_history,
    excel_bash,
  ];
  const tools = excelTools;

  diag(
    `createOfficeBridgeMcp host=${host ?? "both"} → ${tools.length} tools:`,
    tools.map((t) => t?.name).join(", "),
  );

  return createSdkMcpServer({
    name: "office",
    version: "0.1.0",
    // alwaysLoad: ensure these tools are in the initial prompt, not deferred
    // behind tool-search. They are the only sanctioned editing path for the
    // active Office document.
    alwaysLoad: true,
    tools,
  });
}
