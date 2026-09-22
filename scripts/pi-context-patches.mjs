// Local fixes for the vendored pi-for-excel selection context. Each
// replacement must match exactly once, so an upstream change stops the build.
import { replaceOnce } from "./office-agents-patches.mjs";

// Upstream loads the whole selection and ±5 rows around all of it, so selecting
// a full column reads a million cells. Cap the selection part at MAX_SEL_ROWS
// rows and let the column window follow a selection beyond the first 20 columns.
export function patchSelectionContext(input) {
  let source = input.replace(/\r\n/g, "\n");

  source = replaceOnce(
    source,
    "selection address parser",
    "export async function readSelectionContext(): Promise<SelectionContext | null> {",
    `function parseSelectionAddress(address: string) {
  const separator = address.lastIndexOf("!");
  if (separator <= 0 || separator === address.length - 1) return null;
  const qualifiedSheet = address.slice(0, separator);
  const sheetName = qualifiedSheet.startsWith("'") && qualifiedSheet.endsWith("'")
    ? qualifiedSheet.slice(1, -1).replace(/''/g, "'")
    : qualifiedSheet;
  return { sheetName, rangeAddress: address.slice(separator + 1) };
}

export async function readSelectionContext(expectedAddress?: string): Promise<SelectionContext | null> {`,
  );

  source = replaceOnce(
    source,
    "selection fixed to the submitted address",
    "      const sel = context.workbook.getSelectedRange();",
    `      const target = expectedAddress ? parseSelectionAddress(expectedAddress) : null;
      if (expectedAddress && !target) return null;
      const sel = target
        ? context.workbook.worksheets.getItem(target.sheetName).getRange(target.rangeAddress)
        : context.workbook.getSelectedRange();`,
  );

  source = replaceOnce(
    source,
    "selection row cap",
    "const MAX_CONTEXT_COLS = 20;",
    "const MAX_CONTEXT_COLS = 20;\n/** Maximum selected rows read into context */\nconst MAX_SEL_ROWS = 20;",
  );

  source = replaceOnce(
    source,
    "selection metadata only",
    `sel.load("address,values,formulas,rowIndex,columnIndex,rowCount,columnCount,worksheet/name");`,
    `sel.load("address,rowIndex,columnIndex,rowCount,columnCount,worksheet/name");`,
  );

  source = replaceOnce(
    source,
    "column window follows selection",
    `      const usedStartCol = used.columnIndex;
      const usedEndCol = Math.min(
        used.columnIndex + used.columnCount - 1,
        used.columnIndex + MAX_CONTEXT_COLS - 1,
      );`,
    `      const usedStartCol = sel.columnIndex >= used.columnIndex + MAX_CONTEXT_COLS
        ? sel.columnIndex
        : used.columnIndex;
      const usedEndCol = Math.min(
        used.columnIndex + used.columnCount - 1,
        usedStartCol + MAX_CONTEXT_COLS - 1,
      );`,
  );

  source = replaceOnce(
    source,
    "bounded row window",
    `      const contextEndRow = Math.min(usedEndRow, sel.rowIndex + sel.rowCount - 1 + CONTEXT_ROWS);`,
    `      const contextEndRow = Math.min(
        usedEndRow,
        sel.rowIndex + Math.min(sel.rowCount, MAX_SEL_ROWS) - 1 + CONTEXT_ROWS,
      );
      if (contextStartRow > contextEndRow || usedStartCol > usedEndCol) {
        return { address: selAddress, text: \`**Selection:** \${selAddress} (outside the sheet's data)\` };
      }`,
  );

  source = replaceOnce(
    source,
    "bounded selected formulas",
    `      const contextRange = sheet.getRange(contextRangeAddr);
      contextRange.load("values,formulas,address");
      await context.sync();`,
    `      const contextRange = sheet.getRange(contextRangeAddr);
      contextRange.load("values,formulas,address");
      const selHead = sel.getCell(0, 0).getResizedRange(
        Math.min(sel.rowCount, MAX_SEL_ROWS) - 1,
        Math.min(sel.columnCount, MAX_CONTEXT_COLS) - 1,
      );
      selHead.load("formulas");
      await context.sync();`,
  );

  source = replaceOnce(
    source,
    "selected formulas from bounded head",
    "const selFormulas = extractFormulas(sel.formulas, selStartAddress);",
    "const selFormulas = extractFormulas(selHead.formulas, selStartAddress);",
  );

  source = replaceOnce(
    source,
    "quoted sheet names containing exclamation marks",
    `sel.address.split("!")[1] ?? sel.address`,
    `sel.address.slice(sel.address.lastIndexOf("!") + 1)`,
  );

  source = replaceOnce(
    source,
    "bounded cell preview text",
    "lines.push(formatAsMarkdownTable(contextRange.values));",
    `lines.push(formatAsMarkdownTable(contextRange.values.map((row) =>
        row.map((value) => typeof value === "string" && value.length > 160
          ? value.slice(0, 157) + "..."
          : value),
      )));`,
  );

  source = replaceOnce(
    source,
    "bounded selected formula list",
    "**Selected formulas:** ${selFormulas.join(\", \")}",
    "**Selected formulas:** ${selFormulas.slice(0, 20).map((formula) => formula.length > 180 ? formula.slice(0, 177) + \"...\" : formula).join(\", \")}${selFormulas.length > 20 ? `, ... (+${selFormulas.length - 20} more)` : \"\"}",
  );

  source = replaceOnce(
    source,
    "bounded nearby error list",
    "errors.map((e) => `${e.address}=${e.error}`).join(\", \")",
    "errors.slice(0, 20).map((e) => `${e.address}=${e.error}`).join(\", \")}${errors.length > 20 ? `, ... (+${errors.length - 20} more)` : \"\"",
  );

  return source;
}

export function patchChangeTracker(input) {
  let source = replaceOnce(
    input.replace(/\r\n/g, "\n"),
    "truthful change attribution",
    "**User changes since last message:**",
    "**Workbook changes since last message (may include assistant writes):**",
  );
  // Excel tags edits made through this add-in's Office.js calls (ExcelApi 1.14),
  // i.e. the assistant's task-pane writes. Skip those; COM writes and hosts
  // without triggerSource still pass through, hence the label above.
  source = replaceOnce(
    source,
    "skip this add-in's own writes",
    `          sheet.onChanged.add(async (event) => {
            const sheetName`,
    `          sheet.onChanged.add(async (event) => {
            if (event.triggerSource === "ThisLocalAddin") return;
            const sheetName`,
  );
  return source;
}

export function patchWorkbookOverview(input) {
  let source = input.replace(/\r\n/g, "\n");
  source = replaceOnce(
    source,
    "bounded overview header read",
    `// Get header row (first populated row)
      const headerRange = sheet.getRange("1:1").getUsedRangeOrNullObject();`,
    `// Get a bounded preview of the header row
      const headerRange = sheet.getRange("A1:T1").getUsedRangeOrNullObject();`,
  );
  source = replaceOnce(
    source,
    "overview shape count without loading every shape",
    `let shapes: Excel.ShapeCollection | null = null;
      try {
        shapes = sheet.shapes;
        shapes.load("items");
      } catch {
        shapes = null;
      }`,
    `let shapeCountResult: OfficeExtension.ClientResult<number> | null = null;
      try {
        shapeCountResult = sheet.shapes.getCount();
      } catch {
        shapeCountResult = null;
      }`,
  );
  source = replaceOnce(
    source,
    "overview shape count property",
    `const shapeCount = shapes ? shapes.items.length : 0;`,
    `const shapeCount = shapeCountResult?.value || 0;`,
  );
  return source;
}

// Excel on Windows reports an unfilled cell's fill color as "" and rejects ""
// when it is written back (InvalidArgument at RangeFill.color), so restoring
// "no fill" failed. fill.clear() is how Office.js removes a fill.
export function patchRecoveryFormatState(input) {
  return replaceOnce(
    input.replace(/\r\n/g, "\n"),
    "restore no fill with clear()",
    `  if (state.fillColor !== undefined) {
    range.format.fill.color = state.fillColor;
  }`,
    `  if (state.fillColor !== undefined) {
    if (state.fillColor === "") range.format.fill.clear();
    else range.format.fill.color = state.fillColor;
  }`,
  );
}

// Excel returns range precedents/dependents as areas ("Sheet1!A1:A3"), and
// upstream kept only each area's top-left cell, so SUM(A1:A3) traced to A1
// alone. Expand areas into their cells (bounded by the per-node child cap).
export function patchTraceDependencies(input) {
  let source = input.replace(/\r\n/g, "\n");
  source = replaceOnce(
    source,
    "area expansion helper",
    "const MAX_DEPENDENT_SCAN_FORMULA_CELLS = 50_000;",
    `const MAX_DEPENDENT_SCAN_FORMULA_CELLS = 50_000;

function expandTraversalAddresses(address: string, sheetName: string): string[] {
  const bang = address.lastIndexOf("!");
  const prefix = bang >= 0 ? address.slice(0, bang + 1) : "";
  const [start, end] = address.slice(bang + 1).split(":");
  const isCell = (part: string | undefined) => /^\\$?[A-Z]+\\$?\\d+$/i.test(part ?? "");
  if (!isCell(start) || !isCell(end)) {
    const single = normalizeTraversalAddress(address, sheetName);
    return single ? [single] : [];
  }
  const from = parseCell(start as string);
  const to = parseCell(end as string);
  const cells: string[] = [];
  for (let row = from.row; row <= to.row; row++) {
    for (let col = from.col; col <= to.col; col++) {
      const cell = normalizeTraversalAddress(prefix + cellAddress(col, row), sheetName);
      if (cell) cells.push(cell);
      if (cells.length > MAX_CHILDREN_PER_NODE) return cells;
    }
  }
  return cells;
}`,
  );
  for (const indent of ["          ", "        "]) {
    source = replaceOnce(
      source,
      `expand areas (${indent.length}-space loop)`,
      `\n${indent}addChild(normalizeTraversalAddress(address, sheetName));`,
      `\n${indent}for (const cell of expandTraversalAddresses(address, sheetName)) addChild(cell);`,
    );
  }
  return source;
}

// Upstream names every inverse snapshot restore:<its own snapshot id>, so a
// grouped restore (values + format + structure from one call) leaves inverses
// that no longer group, and undoing that restore brings back only part of it.
// Let the caller give one id to every inverse of one restore; without it the
// behavior is upstream's.
const unixNewlines = (text) => text.split(String.fromCharCode(13)).join("");

export function patchRecoveryRestore(input) {
  let source = unixNewlines(input);
  source = replaceOnce(
    source,
    "restore args take the inverse call id",
    "  dependencies: RestoreWorkbookRecoverySnapshotDependencies;\n}",
    "  dependencies: RestoreWorkbookRecoverySnapshotDependencies;\n  /** Shared group identity and inverse application order. */\n  toolCallId?: string;\n  restoreOrder?: number;\n}",
  );
  source = replaceOnce(
    source,
    "resolve the inverse call id",
    "  const { snapshot, scope, dependencies } = args;",
    "  const { snapshot, scope, dependencies } = args;\n  const inverseCallId = args.toolCallId ?? `restore:${snapshot.id}`;",
  );
  const perSnapshot = "toolCallId: `restore:${snapshot.id}`,";
  const count = source.split(perSnapshot).length - 1;
  if (count !== 6) throw new Error(`pi restore patch expected 6 inverse call ids, found ${count}`);
  return source.split(perSnapshot).join("toolCallId: inverseCallId,\n        restoreOrder: args.restoreOrder,");
}

export function patchRecoveryLogRestore(input) {
  let source = replaceOnce(
    unixNewlines(input),
    "restore accepts the inverse call id",
    "  async restore(snapshotId: string): Promise<RestoreWorkbookRecoverySnapshotResult> {",
    "  async restore(snapshotId: string, options: { toolCallId?: string; restoreOrder?: number } = {}): Promise<RestoreWorkbookRecoverySnapshotResult> {",
  );
  source = replaceOnce(
    source,
    "pass the inverse call id on",
    "    return restoreWorkbookRecoverySnapshot({\n      snapshot,\n      scope,",
    "    return restoreWorkbookRecoverySnapshot({\n      snapshot,\n      scope,\n      toolCallId: options.toolCallId,\n      restoreOrder: options.restoreOrder,",
  );
  // Snapshot plus the six append argument interfaces share this metadata.
  const field = "  restoredFromSnapshotId?: string;";
  if (source.split(field).length - 1 !== 7) throw new Error("Expected seven recovery metadata interfaces");
  source = source.split(field).join(`${field}\n  restoreOrder?: number;`);
  const append = "...(args.restoredFromSnapshotId !== undefined ? { restoredFromSnapshotId: args.restoredFromSnapshotId } : {}),";
  if (source.split(append).length - 1 !== 6) throw new Error("Expected six recovery append paths");
  return source.split(append).join(`${append}\n      ...(args.restoreOrder !== undefined ? { restoreOrder: args.restoreOrder } : {}),`);
}

export function patchRecoveryOrderCodec(input) {
  let source = replaceOnce(
    unixNewlines(input),
    "persist inverse application order",
    "  restoredFromSnapshotId: Type.Optional(Type.String()),",
    "  restoredFromSnapshotId: Type.Optional(Type.String()),\n  restoreOrder: Type.Optional(Type.Integer({ minimum: 0 })),",
  );
  return replaceOnce(
    source,
    "reload inverse application order",
    "  if (persisted.restoredFromSnapshotId !== undefined) {",
    "  if (persisted.restoreOrder !== undefined) snapshot.restoreOrder = persisted.restoreOrder;\n  if (persisted.restoredFromSnapshotId !== undefined) {",
  );
}
