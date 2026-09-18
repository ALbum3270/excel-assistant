// Reproducible local fixes for the vendored office-agents Excel API.
// Each replacement must match exactly once. An upstream change therefore
// stops the build instead of silently dropping a safety fix.

function replaceOnce(source, name, before, after) {
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
      "To proceed with overwriting existing data, retry with allow_overwrite set to true.",
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
    `      const destRange = sheet.getRange(copyToRange);
      destRange.copyFrom(range, Excel.RangeCopyType.all);`,
    `      const destRange = copyDestination ?? sheet.getRange(copyToRange);
      destRange.copyFrom(range, Excel.RangeCopyType.all);`,
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

  return source;
}
