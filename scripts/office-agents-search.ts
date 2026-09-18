// Inserted into the pinned upstream api.ts by the vendor build.
// Reuses its sheet-id and A1 helpers; this is not a standalone module.
export async function searchData(
  searchTerm: string,
  options: {
    sheetId?: number;
    range?: string;
    offset?: number;
    cursor?: string;
    matchCase?: boolean;
    matchEntireCell?: boolean;
    matchFormulas?: boolean;
    useRegex?: boolean;
    maxResults?: number;
  } = {},
): Promise<SearchDataResult> {
  const {
    sheetId,
    range,
    offset = 0,
    cursor,
    matchCase = false,
    matchEntireCell = false,
    matchFormulas = false,
    useRegex = false,
    maxResults = 500,
  } = options;
  if (
    !Number.isSafeInteger(offset) ||
    offset < 0 ||
    !Number.isInteger(maxResults) ||
    maxResults < 1 ||
    maxResults > 5000
  ) {
    throw new Error("offset must be nonnegative; maxResults must be between 1 and 5000");
  }
  const scope = JSON.stringify([
    searchTerm,
    sheetId ?? null,
    range ?? null,
    offset,
    matchCase,
    matchEntireCell,
    matchFormulas,
    useRegex,
  ]);
  const resume = cursor ? JSON.parse(cursor) : null;
  if (
    resume &&
    (resume.scope !== scope ||
      !Number.isInteger(resume.sheetId) ||
      !Number.isSafeInteger(resume.cellOffset) ||
      resume.cellOffset < 0 ||
      !Number.isSafeInteger(resume.matchedCount) ||
      resume.matchedCount < 0)
  ) {
    throw new Error("Search cursor does not match this search; restart without cursor");
  }
  const pattern = useRegex ? new RegExp(searchTerm, matchCase ? "" : "i") : null;
  const compareTerm = matchCase ? searchTerm : searchTerm.toLowerCase();

  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items");
    await context.sync();
    for (const sheet of sheets.items) sheet.load("id,name");
    await context.sync();
    const ids = await preloadSheetIds(sheets.items);
    const targets =
      sheetId === undefined
        ? sheets.items
        : sheets.items.filter((sheet) => ids.get(sheet.id) === sheetId);
    if (sheetId !== undefined && targets.length === 0) {
      throw new Error(`Worksheet with ID ${sheetId} not found`);
    }
    const firstSheet = resume
      ? targets.findIndex((sheet) => ids.get(sheet.id) === resume.sheetId)
      : 0;
    if (firstSheet < 0)
      throw new Error("Search worksheet no longer exists; restart without cursor");
    const matches: SearchMatch[] = [];
    let matchedCount = resume?.matchedCount ?? 0;
    let scannedCells = 0;
    let loadedCells = 0;

    const result = (nextCursor: string | null) => ({
      success: true,
      matches,
      totalFound: matchedCount,
      totalFoundIsExact: nextCursor === null,
      returned: matches.length,
      offset: Math.max(offset, resume?.matchedCount ?? 0),
      hasMore: nextCursor !== null,
      scannedCells,
      searchTerm,
      searchScope: sheetId === undefined ? "All sheets" : `Sheet ${sheetId}`,
      // Match offsets alone cannot advance through a page with no hits.
      nextOffset: null,
      nextCursor,
    });
    const continuation = (index: number, cellOffset: number) =>
      JSON.stringify({
        sheetId: ids.get(targets[index].id),
        cellOffset,
        matchedCount,
        scope,
      });

    for (let index = firstSheet; index < targets.length; index++) {
      const sheet = targets[index];
      const target = range ? sheet.getRange(range) : sheet.getUsedRangeOrNullObject();
      target.load("address,rowCount,columnCount");
      await context.sync();
      if (target.isNullObject) continue;
      const { startRow, startCol } = parseRangeAddress(target.address);
      const totalCells = target.rowCount * target.columnCount;
      let position = index === firstSheet && resume ? resume.cellOffset : 0;

      while (position < totalCells) {
        if (loadedCells >= 20000 || matches.length >= maxResults) {
          return result(continuation(index, position));
        }
        const row = Math.floor(position / target.columnCount);
        const column = position % target.columnCount;
        const capacity = Math.min(2000, 20000 - loadedCells);
        const width = Math.min(target.columnCount - column, capacity);
        const height =
          column === 0 && width === target.columnCount
            ? Math.min(target.rowCount - row, Math.floor(capacity / width))
            : 1;
        const block = sheet.getRange(
          cellAddress(startRow + row, startCol + column) +
            ":" +
            cellAddress(startRow + row + height - 1, startCol + column + width - 1),
        );
        block.load("values,formulas");
        await context.sync();
        loadedCells += width * height;

        for (let r = 0; r < height; r++) {
          for (let c = 0; c < width; c++) {
            const value = block.values[r][c];
            const formula = block.formulas[r][c];
            const text = matchFormulas && formula ? String(formula) : String(value ?? "");
            const compared = matchCase ? text : text.toLowerCase();
            const isMatch = pattern
              ? pattern.test(text)
              : matchEntireCell
                ? compared === compareTerm
                : compared.includes(compareTerm);
            position++;
            scannedCells++;
            if (!isMatch) continue;
            matchedCount++;
            if (matchedCount <= offset) continue;
            matches.push({
              sheetName: sheet.name,
              sheetId: ids.get(sheet.id)!,
              a1: cellAddress(startRow + row + r, startCol + column + c),
              value: value as string | number | boolean,
              formula: typeof formula === "string" && formula.startsWith("=") ? formula : null,
              row: startRow + row + r + 1,
              column: startCol + column + c + 1,
            });
            if (matches.length >= maxResults) {
              if (position < totalCells) return result(continuation(index, position));
              return result(index + 1 < targets.length ? continuation(index + 1, 0) : null);
            }
          }
        }
      }
    }
    return result(null);
  });
}
