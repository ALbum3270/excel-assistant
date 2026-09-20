// Reproducible local fixes for the vendored office-agents Excel API.
// Each replacement must match exactly once. An upstream change therefore
// stops the build instead of silently dropping a safety fix.
import { patchBoundedReads } from "./office-agents-read-patches.mjs";
import { readFileSync } from "node:fs";

export function replaceOnce(source, name, before, after) {
  const first = source.indexOf(before);
  const second = first === -1 ? -1 : source.indexOf(before, first + before.length);
  if (first === -1 || second !== -1) {
    throw new Error(
      `office-agents patch "${name}" expected exactly one match (found ${first === -1 ? 0 : 2})`,
    );
  }
  return source.slice(0, first) + after + source.slice(first + before.length);
}

export function patchOfficeAgentsExcelApi(input) {
  let source = input.replace(/\r\n/g, "\n");

  source = replaceOnce(
    source,
    "overwrite helpers",
    `function excelColorToHex(
  color: Excel.RangeFont | Excel.RangeFill,
): string | undefined {`,
    `function nonEmptyCellAddresses(
  range: Excel.Range,
  shouldCheck: (row: number, column: number, address: string) => boolean = () => true,
): string[] {
  const nonEmpty: string[] = [];
  const { startCol, startRow } = parseRangeAddress(range.address);
  for (let r = 0; r < range.rowCount; r++) {
    for (let c = 0; c < range.columnCount; c++) {
      const address = cellAddress(startRow + r, startCol + c);
      if (!shouldCheck(r, c, address)) continue;
      const value = range.values[r][c];
      const formula = range.formulas[r][c];
      if (
        (value !== null && value !== "" && value !== undefined) ||
        (typeof formula === "string" && formula.startsWith("="))
      ) {
        nonEmpty.push(address);
      }
    }
  }
  return nonEmpty;
}

function rangeCellAddresses(range: Excel.Range): Set<string> {
  const addresses = new Set<string>();
  const { startCol, startRow } = parseRangeAddress(range.address);
  for (let r = 0; r < range.rowCount; r++) {
    for (let c = 0; c < range.columnCount; c++) {
      addresses.add(cellAddress(startRow + r, startCol + c));
    }
  }
  return addresses;
}

function throwOverwriteError(addresses: string[]): void {
  if (addresses.length === 0) return;
  const unique = Array.from(new Set(addresses));
  const cellList =
    unique.length <= 10 ? unique.join(", ") : \
      \`\${unique.slice(0, 10).join(", ")}...\`;
  throw new Error(
    \`Would overwrite \${unique.length} non-empty cell(s): \${cellList}. \` +
      "If the requested edit targets these cells, retry with allow_overwrite set to true; otherwise ask before overwriting.",
  );
}

function excelColorToHex(
  color: Excel.RangeFont | Excel.RangeFill,
): string | undefined {`,
  );

  source = replaceOnce(
    source,
    "range input validation",
    `  const { includeStyles = true, cellLimit = 2000 } = options;

  return Excel.run(async (context) => {`,
    `  const { includeStyles = true, cellLimit = 2000 } = options;
  if (!Array.isArray(ranges) || ranges.length === 0) {
    throw new Error("ranges must be a non-empty array");
  }
  if (!Number.isInteger(cellLimit) || cellLimit <= 0) {
    throw new Error("cellLimit must be a positive integer");
  }

  return Excel.run(async (context) => {`,
  );

  source = replaceOnce(
    source,
    "honest range truncation",
    `      for (let r = 0; r < range.rowCount && totalCells < cellLimit; r++) {
        for (let c = 0; c < range.columnCount && totalCells < cellLimit; c++) {
          const addr = cellAddress(startRow + r, startCol + c);
          const value = range.values[r][c];
          const formula = range.formulas[r][c];

          if (value !== null && value !== "" && value !== undefined) {
            cells[addr] = value as string | number | boolean;
            totalCells++;
            if (includeStyles) {
              styleTargetsMap.set(addr, range.getCell(r, c));
            }
          }

          if (typeof formula === "string" && formula.startsWith("=")) {
            formulas[addr] = formula;
            if (includeStyles) {
              styleTargetsMap.set(addr, range.getCell(r, c));
            }
          }
        }
      }`,
    `      scanRange: for (let r = 0; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const addr = cellAddress(startRow + r, startCol + c);
          const value = range.values[r][c];
          const formula = range.formulas[r][c];
          const hasValue = value !== null && value !== "" && value !== undefined;
          const hasFormula = typeof formula === "string" && formula.startsWith("=");
          if (!hasValue && !hasFormula) continue;
          if (totalCells >= cellLimit) {
            hasMore = true;
            break scanRange;
          }

          if (hasValue) cells[addr] = value as string | number | boolean;
          if (hasFormula) formulas[addr] = formula;
          totalCells++;
          if (includeStyles) styleTargetsMap.set(addr, range.getCell(r, c));
        }
      }`,
  );

  source = replaceOnce(
    source,
    "verify later ranges before reporting truncation",
    `    for (const rangeAddr of ranges) {
      if (totalCells >= cellLimit) {
        hasMore = true;
        break;
      }`,
    `    for (const rangeAddr of ranges) {
      if (hasMore) break;`,
  );

  source = replaceOnce(
    source,
    "cell matrix validation",
    `  const { copyToRange, resizeWidth, resizeHeight, allowOverwrite } = options;

  return Excel.run(async (context) => {`,
    `  const { copyToRange, resizeWidth, resizeHeight, allowOverwrite } = options;
  if (!Array.isArray(cells) || cells.length === 0 || !Array.isArray(cells[0]) || cells[0].length === 0) {
    throw new Error("cells must be a non-empty rectangular 2D array");
  }
  const width = cells[0].length;
  if (cells.some((row) => !Array.isArray(row) || row.length !== width)) {
    throw new Error("cells must be rectangular; every row must have the same length");
  }
  for (const row of cells) {
    for (const cell of row) {
      if (cell.formula !== undefined && !cell.formula.startsWith("=")) {
        throw new Error("cell formulas must start with '='");
      }
    }
  }

  return Excel.run(async (context) => {`,
  );

  source = replaceOnce(
    source,
    "complete overwrite protection",
    `    if (!allowOverwrite) {
      const nonEmptyCells: string[] = [];
      const { startCol, startRow } = parseRangeAddress(range.address);

      for (let r = 0; r < range.rowCount; r++) {
        for (let c = 0; c < range.columnCount; c++) {
          const value = range.values[r][c];
          const formula = range.formulas[r][c];
          const hasValue =
            value !== null && value !== "" && value !== undefined;
          const hasFormula =
            typeof formula === "string" && formula.startsWith("=");

          if (hasValue || hasFormula) {
            nonEmptyCells.push(cellAddress(startRow + r, startCol + c));
          }
        }
      }

      if (nonEmptyCells.length > 0) {
        const cellList =
          nonEmptyCells.length <= 10
            ? nonEmptyCells.join(", ")
            : \`\${nonEmptyCells.slice(0, 10).join(", ")}...\`;
        throw new Error(
          \`Would overwrite \${nonEmptyCells.length} non-empty cell(s): \${cellList}. \` +
            \`To proceed with overwriting existing data, retry with allow_overwrite set to true.\`,
        );
      }
    }`,
    `    let copyDestination: Excel.Range | null = null;
    if (copyToRange && !allowOverwrite) {
      copyDestination = sheet.getRange(copyToRange);
      copyDestination.load("rowCount,columnCount,values,formulas,address");
      await context.sync();
    }

    if (!allowOverwrite) {
      const overwritten = nonEmptyCellAddresses(range, (r, c) => {
        const cell = cells[r]?.[c];
        return Boolean(
          cell?.formula ||
          (cell && Object.prototype.hasOwnProperty.call(cell, "value")),
        );
      });
      if (copyDestination) {
        const sourceAddresses = rangeCellAddresses(range);
        overwritten.push(
          ...nonEmptyCellAddresses(copyDestination, (_r, _c, address) =>
            !sourceAddresses.has(address),
          ),
        );
      }
      throwOverwriteError(overwritten);
    }`,
  );

  source = replaceOnce(
    source,
    "preserve sparse cell contents",
    `    const values: unknown[][] = [];
    const formulas: (string | null)[][] = [];
    let hasFormulas = false;

    for (let r = 0; r < cells.length; r++) {
      values[r] = [];
      formulas[r] = [];
      for (let c = 0; c < cells[r].length; c++) {
        const cell = cells[r][c];
        if (cell.formula) {
          formulas[r][c] = cell.formula;
          values[r][c] = null;
          hasFormulas = true;
        } else {
          values[r][c] = cell.value ?? null;
          formulas[r][c] = null;
        }
      }
    }

    if (hasFormulas) {
      range.formulas = formulas.map((row, r) =>
        row.map((f, c) => f ?? values[r][c]),
      );
    } else {
      range.values = values;
    }`,
    `    const writeMatrix: unknown[][] = [];
    const formulas: (string | null)[][] = [];
    let hasDataWrites = false;
    let hasFormulas = false;

    for (let r = 0; r < cells.length; r++) {
      writeMatrix[r] = [];
      formulas[r] = [];
      for (let c = 0; c < cells[r].length; c++) {
        const cell = cells[r][c];
        if (cell.formula !== undefined) {
          writeMatrix[r][c] = cell.formula;
          formulas[r][c] = cell.formula;
          hasDataWrites = true;
          hasFormulas = true;
        } else if (Object.prototype.hasOwnProperty.call(cell, "value")) {
          writeMatrix[r][c] = cell.value ?? null;
          formulas[r][c] = null;
          hasDataWrites = true;
        } else {
          const existingFormula = range.formulas[r][c];
          writeMatrix[r][c] =
            typeof existingFormula === "string" && existingFormula.startsWith("=")
              ? existingFormula
              : range.values[r][c];
          formulas[r][c] = null;
        }
      }
    }

    if (hasDataWrites) range.formulas = writeMatrix;`,
  );

  source = replaceOnce(
    source,
    "reuse checked copy destination",
    `    if (copyToRange) {
      const destRange = sheet.getRange(copyToRange);
      destRange.copyFrom(range, Excel.RangeCopyType.all);
      await context.sync();
    }`,
    `    let verificationRange = range;
    if (copyToRange) {
      const destRange = copyDestination ?? sheet.getRange(copyToRange);
      destRange.copyFrom(range, Excel.RangeCopyType.all);
      await context.sync();
      verificationRange = destRange;
    }`,
  );

  source = replaceOnce(
    source,
    "write verification result",
    `    const formulaResults: Record<string, unknown> = {};
    if (hasFormulas) {
      range.load("values,address");
      await context.sync();
      const { startCol, startRow } = parseRangeAddress(range.address);
      for (let r = 0; r < range.values.length; r++) {
        for (let c = 0; c < range.values[r].length; c++) {
          if (formulas[r]?.[c]) {
            formulaResults[cellAddress(startRow + r, startCol + c)] =
              range.values[r][c];
          }
        }
      }
    }

    return {
      success: true,
      cellsWritten: cells.flat().length,
      ...(Object.keys(formulaResults).length > 0 && { formulaResults }),`,
    `    const formulaResults: Record<string, unknown> = {};
    const formulaErrors: Array<{ address: string; value: string }> = [];
    verificationRange.load("values,formulas,address");
    await context.sync();
    const { startCol, startRow } = parseRangeAddress(verificationRange.address);
    for (let r = 0; r < verificationRange.values.length; r++) {
      for (let c = 0; c < verificationRange.values[r].length; c++) {
        if (typeof verificationRange.formulas[r]?.[c] !== "string" || !verificationRange.formulas[r][c].startsWith("=")) continue;
        const address = cellAddress(startRow + r, startCol + c);
        const value = verificationRange.values[r][c];
        formulaResults[address] = value;
        if (typeof value === "string" && /^#(?:REF!|VALUE!|NAME\\?|DIV\\/0!|N\\/A|NUM!|NULL!|SPILL!|CALC!)/i.test(value)) {
          formulaErrors.push({ address, value });
        }
      }
    }

    return {
      success: true,
      commitStatus: "committed",
      writtenRange: verificationRange.address.split("!")[1] || verificationRange.address,
      cellsWritten: cells.flat().length,
      cellsCommitted: verificationRange.values.length * (verificationRange.values[0]?.length ?? 0),
      ...(Object.keys(formulaResults).length > 0 && { formulaResults }),
      ...(formulaErrors.length > 0 && { formulaErrors }),`,
  );

  source = replaceOnce(
    source,
    "copy overwrite protection",
    `export async function copyTo(
  sheetId: number,
  sourceRange: string,
  destinationRange: string,
): Promise<CopyToResult> {
  return Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) throw new Error(\`Worksheet with ID \${sheetId} not found\`);

    const source = sheet.getRange(sourceRange);
    const dest = sheet.getRange(destinationRange);
    dest.copyFrom(source, Excel.RangeCopyType.all);`,
    `export async function copyTo(
  sheetId: number,
  sourceRange: string,
  destinationRange: string,
  allowOverwrite = false,
): Promise<CopyToResult> {
  return Excel.run(async (context) => {
    const sheet = await getWorksheetById(context, sheetId);
    if (!sheet) throw new Error(\`Worksheet with ID \${sheetId} not found\`);

    const source = sheet.getRange(sourceRange);
    const dest = sheet.getRange(destinationRange);
    if (!allowOverwrite) {
      source.load("rowCount,columnCount,address");
      dest.load("rowCount,columnCount,values,formulas,address");
      await context.sync();
      const sourceAddresses = rangeCellAddresses(source);
      throwOverwriteError(
        nonEmptyCellAddresses(dest, (_r, _c, address) => !sourceAddresses.has(address)),
      );
    }
    dest.copyFrom(source, Excel.RangeCopyType.all);`,
  );

  source = replaceOnce(
    source,
    "copy commit result",
    `    return {
      success: true,
      source: sourceRange,
      destination: destinationRange,
    };`,
    `    dest.load("values,formulas,address");
    await context.sync();
    const formulaErrors: Array<{ address: string; value: string }> = [];
    const { startCol, startRow } = parseRangeAddress(dest.address);
    for (let r = 0; r < dest.values.length; r++) {
      for (let c = 0; c < dest.values[r].length; c++) {
        const formula = dest.formulas[r]?.[c];
        const value = dest.values[r][c];
        if (
          typeof formula === "string" &&
          formula.startsWith("=") &&
          typeof value === "string" &&
          /^#(?:REF!|VALUE!|NAME\\?|DIV\\/0!|N\\/A|NUM!|NULL!|SPILL!|CALC!)/i.test(value)
        ) {
          formulaErrors.push({ address: cellAddress(startRow + r, startCol + c), value });
        }
      }
    }
    return {
      success: true,
      commitStatus: "committed",
      source: sourceRange,
      destination: destinationRange,
      ...(formulaErrors.length > 0 && { formulaErrors }),
    };`,
  );

  source = replaceOnce(
    source,
    "sheet structure input validation",
    `  const {
    operation,
    dimension,
    reference,
    count = 1,
    position = "before",
  } = params;

  return Excel.run(async (context) => {`,
    `  const {
    operation,
    dimension,
    reference,
    count = 1,
    position = "before",
  } = params;
  if (!Number.isInteger(count) || count <= 0) {
    throw new Error("count must be a positive integer");
  }
  if (!["freeze", "unfreeze"].includes(operation)) {
    if (!reference?.trim()) {
      throw new Error(\`reference is required for \${operation}\`);
    }
    if (dimension === "rows" && !/^[1-9]\\d*$/.test(reference)) {
      throw new Error("row reference must be a positive row number");
    }
    if (dimension === "columns" && !/^[A-Za-z]{1,3}$/.test(reference)) {
      throw new Error("column reference must contain only column letters");
    }
    if (dimension === "columns" && letterToColumnIndex(reference) > 16383) {
      throw new Error("column reference must be between A and XFD");
    }
  }

  return Excel.run(async (context) => {`,
  );

  const searchStart = source.indexOf("export async function searchData(");
  const searchEnd = source.indexOf("export interface ExcelObject {", searchStart);
  if (searchStart < 0 || searchEnd < 0) throw new Error("Missing upstream search function");
  source = replaceOnce(
    source,
    "bounded search with scan cursor",
    source.slice(searchStart, searchEnd),
    readFileSync(new URL("./office-agents-search.ts", import.meta.url), "utf8") + "\n\n",
  );
  source = replaceOnce(
    source,
    "search continuation result type",
    "export interface SearchDataResult {",
    `export interface SearchDataResult {
  nextCursor: string | null;
  scannedCells: number;
  totalFoundIsExact: boolean;`,
  );
  // Excel on Windows loads excel-win32-16.01.js, which has CommentCollection
  // but no NoteCollection, so sheet.notes.add() throws there. Threaded
  // comments are the equivalent the host does expose, and pi-for-excel's
  // recovery snapshots cover them.
  source = replaceOnce(
    source,
    "cell notes as threaded comments",
    `        if (cell.note) {
          cellRange.load("address");
          await context.sync();
          const cellAddr = cellRange.address.split("!")[1] || cellRange.address;
          sheet.notes.add(cellAddr, cell.note);
        }`,
    `        if (cell.note) {
          cellRange.load("address");
          const comments = sheet.comments;
          comments.load("items");
          await context.sync();
          const located = comments.items.map((comment) => {
            const location = comment.getLocation();
            location.load("address");
            return { comment, location };
          });
          if (located.length > 0) await context.sync();
          const target = (cellRange.address.split("!").pop() ?? "").toUpperCase();
          let replaced = false;
          for (const entry of located) {
            const at = (entry.location.address.split("!").pop() ?? "").toUpperCase();
            if (at !== target) continue;
            entry.comment.delete();
            replaced = true;
          }
          if (replaced) await context.sync();
          sheet.comments.add(cellRange, cell.note);
        }`,
  );
  // A chart or shape can be the current selection, and workbook.getSelectedRange()
  // then throws InvalidSelection at sync — which failed the whole metadata read
  // right after the agent created a chart. Read the selection in its own sync so
  // a non-range selection only costs the selection field.
  source = replaceOnce(
    source,
    "metadata tolerates a non-range selection",
    `    const selectedRange = workbook.getSelectedRange();
    selectedRange.load("address");

    await context.sync();`,
    `    await context.sync();

    let selectedAddress: string | null = null;
    try {
      const selectedRange = workbook.getSelectedRange();
      selectedRange.load("address");
      await context.sync();
      selectedAddress = selectedRange.address;
    } catch {
      // A chart, shape or nothing at all is selected; not a range.
      selectedAddress = null;
    }`,
  );
  source = replaceOnce(
    source,
    "metadata selection address may be absent",
    `    const rangeAddress = selectedRange.address.includes("!")
      ? selectedRange.address.split("!")[1]
      : selectedRange.address;`,
    `    const rangeAddress = selectedAddress
      ? selectedAddress.includes("!")
        ? selectedAddress.split("!")[1]
        : selectedAddress
      : null;`,
  );
  source = replaceOnce(
    source,
    "metadata selection logging",
    `    console.log(
      "[getWorkbookMetadata] selectedRange.address:",
      selectedRange.address,
    );`,
    `    console.log("[getWorkbookMetadata] selectedRange.address:", selectedAddress);`,
  );
  return patchBoundedReads(source, replaceOnce);
}
