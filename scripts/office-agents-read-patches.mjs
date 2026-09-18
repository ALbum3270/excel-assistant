// Bound host reads before values/formulas cross the Office.js bridge.
export function patchBoundedReads(source, replaceOnce) {
  const patch = (name, before, after) => {
    source = replaceOnce(source, name, before, after);
  };

  patch(
    "sparse continuation type",
    `export interface GetCellRangesResult {
  success: boolean;`,
    `export interface GetCellRangesResult {
  remainingRanges: string[];
  success: boolean;`,
  );
  patch(
    "CSV continuation type",
    `export interface GetRangeAsCsvResult {
  success: boolean;`,
    `export interface GetRangeAsCsvResult {
  nextRange: string | null;
  success: boolean;`,
  );
  patch(
    "bounded read planner",
    "export async function getCellRanges(",
    `// Return the unvisited part of a rectangle in row-major order.
function remainingReadRanges(range: Excel.Range, row: number, column: number): string[] {
  const { startRow, startCol } = parseRangeAddress(range.address);
  const endRow = startRow + range.rowCount - 1;
  const endCol = startCol + range.columnCount - 1;
  if (row >= range.rowCount) return [];
  const tail: string[] = [];
  if (column > 0) {
    tail.push(cellAddress(startRow + row, startCol + column) + ":" + cellAddress(startRow + row, endCol));
    row++;
  }
  if (row < range.rowCount) {
    tail.push(cellAddress(startRow + row, startCol) + ":" + cellAddress(endRow, endCol));
  }
  return tail;
}

async function planReadRanges(context: Excel.RequestContext, sheet: Excel.Worksheet, ranges: string[]) {
  const targets = ranges.map((address) => sheet.getRange(address));
  for (const target of targets) target.load("address,rowCount,columnCount");
  await context.sync();
  const pages: string[] = [];
  let budget = 20000;
  for (let i = 0; i < targets.length; i++) {
    const target = targets[i];
    const { startRow, startCol } = parseRangeAddress(target.address);
    let row = 0, column = 0;
    while (row < target.rowCount) {
      if (budget === 0) {
        return { pages, remaining: [...remainingReadRanges(target, row, column), ...ranges.slice(i + 1)] };
      }
      const width = Math.min(target.columnCount - column, 2000, budget);
      const height = column === 0 && width === target.columnCount
        ? Math.min(target.rowCount - row, Math.floor(Math.min(2000, budget) / width)) : 1;
      const count = width * height;
      pages.push(row === 0 && column === 0 && height === target.rowCount && width === target.columnCount
        ? ranges[i]
        : cellAddress(startRow + row, startCol + column) + ":" + cellAddress(startRow + row + height - 1, startCol + column + width - 1));
      budget -= count;
      column += width;
      if (column === target.columnCount) { column = 0; row += height; }
    }
  }
  return { pages, remaining: [] as string[] };
}

export async function getCellRanges(`,
  );
  patch(
    "plan sparse reads before loading values",
    `    for (const rangeAddr of ranges) {
      if (hasMore) break;`,
    `    const readPlan = await planReadRanges(context, sheet, ranges);
    let remainingRanges = readPlan.remaining;
    for (let pageIndex = 0; pageIndex < readPlan.pages.length; pageIndex++) {
      if (hasMore) break;
      const rangeAddr = readPlan.pages[pageIndex];`,
  );
  patch(
    "sparse read continuation",
    `          if (totalCells >= cellLimit) {
            hasMore = true;
            break scanRange;
          }`,
    `          if (totalCells >= cellLimit) {
            hasMore = true;
            remainingRanges = [
              ...remainingReadRanges(range, r, c),
              ...readPlan.pages.slice(pageIndex + 1),
              ...readPlan.remaining,
            ];
            break scanRange;
          }`,
  );
  patch(
    "sparse read budget status",
    `      hasMore,
      worksheet: {`,
    `      hasMore: hasMore || remainingRanges.length > 0,
      remainingRanges,
      worksheet: {`,
  );
  patch(
    "CSV positive row limit",
    `  const { includeHeaders = true, maxRows = 500 } = options;`,
    `  const { includeHeaders = true, maxRows = 500 } = options;
  if (!Number.isInteger(maxRows) || maxRows <= 0 || maxRows > 20000) {
    throw new Error("maxRows must be an integer from 1 to 20000");
  }`,
  );
  patch(
    "CSV bounded host read",
    `    range.load("values,rowCount,columnCount");
    await context.sync();

    const startRow = includeHeaders ? 0 : 1;
    const availableRows = range.rowCount - startRow;
    const actualRows = Math.min(availableRows, maxRows);
    const hasMore = availableRows > maxRows;

    const rows: string[] = [];
    for (let r = startRow; r < startRow + actualRows; r++) {
      const row = range.values[r].map((v) => {`,
    `    range.load("address,rowCount,columnCount");
    await context.sync();

    const startRow = includeHeaders ? 0 : 1;
    const availableRows = Math.max(0, range.rowCount - startRow);
    const actualRows = Math.min(availableRows, maxRows, Math.max(1, Math.floor(20000 / range.columnCount)));
    const hasMore = availableRows > actualRows;
    const { startRow: firstRow, startCol } = parseRangeAddress(range.address);
    const endColumn = startCol + range.columnCount - 1;
    const preview = actualRows === 0 ? null : sheet.getRange(
      cellAddress(firstRow + startRow, startCol) + ":" + cellAddress(firstRow + startRow + actualRows - 1, endColumn),
    );
    if (preview) { preview.load("values"); await context.sync(); }
    const nextRange = hasMore
      ? cellAddress(firstRow + startRow + actualRows, startCol) + ":" + cellAddress(firstRow + range.rowCount - 1, endColumn)
      : null;

    const rows: string[] = [];
    for (let r = 0; r < actualRows; r++) {
      const row = preview!.values[r].map((v) => {`,
  );
  patch(
    "CSV continuation result",
    `      columnCount: range.columnCount,
      hasMore,
      sheetName: sheet.name,`,
    `      columnCount: range.columnCount,
      hasMore,
      nextRange,
      sheetName: sheet.name,`,
  );
  return source;
}
