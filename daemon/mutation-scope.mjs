const MAX_EXCEL_ROW = 1_048_576;
const MAX_EXCEL_COLUMN = 16_384;

function columnNumber(label) {
  return [...label.toUpperCase()].reduce((value, char) => value * 26 + char.charCodeAt(0) - 64, 0);
}

function columnLetters(number) {
  let result = "";
  for (let value = number; value > 0; value = Math.floor((value - 1) / 26)) {
    result = String.fromCharCode(65 + ((value - 1) % 26)) + result;
  }
  return result;
}

function unquoteSheetName(value) {
  const text = String(value ?? "").trim();
  return text.startsWith("'") && text.endsWith("'")
    ? text.slice(1, -1).replaceAll("''", "'")
    : text;
}

export function parseScopedRange(value) {
  const text = String(value ?? "").trim();
  if (!text) return null;
  const bang = text.lastIndexOf("!");
  const sheetName = bang >= 0 ? unquoteSheetName(text.slice(0, bang)) : null;
  const local = (bang >= 0 ? text.slice(bang + 1) : text).replaceAll("$", "").trim();
  let match = /^([A-Z]{1,3})([1-9]\d*)(?::([A-Z]{1,3})([1-9]\d*))?$/i.exec(local);
  if (match) {
    const startColumn = columnNumber(match[1]);
    const endColumn = columnNumber(match[3] ?? match[1]);
    const startRow = Number(match[2]);
    const endRow = Number(match[4] ?? match[2]);
    if (
      startColumn > MAX_EXCEL_COLUMN || endColumn > MAX_EXCEL_COLUMN ||
      startRow > MAX_EXCEL_ROW || endRow > MAX_EXCEL_ROW ||
      endColumn < startColumn || endRow < startRow
    ) return null;
    return { sheetName, address: local.toUpperCase(), startColumn, endColumn, startRow, endRow };
  }
  match = /^([A-Z]{1,3}):([A-Z]{1,3})$/i.exec(local);
  if (match) {
    const startColumn = columnNumber(match[1]);
    const endColumn = columnNumber(match[2]);
    if (startColumn > MAX_EXCEL_COLUMN || endColumn > MAX_EXCEL_COLUMN || endColumn < startColumn) return null;
    return { sheetName, address: local.toUpperCase(), startColumn, endColumn, startRow: 1, endRow: MAX_EXCEL_ROW };
  }
  match = /^([1-9]\d*):([1-9]\d*)$/.exec(local);
  if (match) {
    const startRow = Number(match[1]);
    const endRow = Number(match[2]);
    if (startRow > MAX_EXCEL_ROW || endRow > MAX_EXCEL_ROW || endRow < startRow) return null;
    return { sheetName, address: local, startColumn: 1, endColumn: MAX_EXCEL_COLUMN, startRow, endRow };
  }
  return null;
}

function rangeTokens(text) {
  const pattern = /(?<![\w.])(?:(?:'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_.]*)!)?\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?|(?<![\w.])(?:\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|[1-9]\d*:[1-9]\d*)(?![\w.])/giu;
  return [...String(text ?? "").matchAll(pattern)].map((match) => ({ value: match[0].trim(), index: match.index ?? 0 }));
}

function sameRange(a, b) {
  return (
    a.startColumn === b.startColumn && a.endColumn === b.endColumn &&
    a.startRow === b.startRow && a.endRow === b.endRow &&
    String(a.sheetName ?? "").toLowerCase() === String(b.sheetName ?? "").toLowerCase()
  );
}

function addScopeRange(result, value, reason, defaultSheet = null) {
  const parsed = parseScopedRange(value);
  if (!parsed) return;
  if (!parsed.sheetName && defaultSheet) parsed.sheetName = defaultSheet;
  if (result.ranges.some((range) => sameRange(range, parsed))) return;
  result.ranges.push({ ...parsed, reason });
}

const TARGET_WORDS = /(?:修改|覆盖|填充|填写|写入|写到|清空|清除|删除|排序|格式化|设置|替换|更新|放到|生成到|输出到|edit|update|write|put|fill|clear|delete|remove|sort|format|replace|overwrite|target|output)/iu;
const SOURCE_END = /(?:根据|依据|来自|读取|参照|使用|from|using|based\s+on|source)\s*$/iu;

export function buildMutationScope(text, context = {}, sheets = []) {
  const result = { ranges: [] };
  const selection = parseScopedRange(context?.selection?.address);
  if (selection) result.ranges.push({ ...selection, reason: "submitted_selection" });
  const defaultSheet = selection?.sheetName ?? sheets.find?.((sheet) => sheet.active)?.name ?? null;
  const source = String(text ?? "");

  const answerBlock = /###\s*answer_position\s*\r?\n\s*([^\r\n]+)/iu.exec(source);
  if (answerBlock) {
    for (const token of rangeTokens(answerBlock[1])) {
      addScopeRange(result, token.value, "answer_position", defaultSheet);
    }
  }

  for (const token of rangeTokens(source)) {
    const preceding = source[token.index - 1] ?? "";
    if ("=+*/(-,".includes(preceding)) continue;
    const before = source.slice(Math.max(0, token.index - 28), token.index);
    const after = source.slice(token.index + token.value.length, token.index + token.value.length + 28);
    if (SOURCE_END.test(before)) continue;
    if (TARGET_WORDS.test(before) || TARGET_WORDS.test(after)) {
      addScopeRange(result, token.value, "explicit_target", defaultSheet);
    }
  }

  const idByName = new Map(
    (Array.isArray(sheets) ? sheets : []).map((sheet) => [String(sheet.name).toLowerCase(), sheet.id]),
  );
  for (const range of result.ranges) {
    if (range.sheetName) range.sheetId = idByName.get(range.sheetName.toLowerCase()) ?? null;
  }
  return result;
}

export function bindScopeSheetIds(scope, sheets = []) {
  const idByName = new Map(
    (Array.isArray(sheets) ? sheets : []).map((sheet) => [String(sheet.name).toLowerCase(), sheet.id]),
  );
  for (const range of scope?.ranges ?? []) {
    if (range.sheetName) range.sheetId = idByName.get(range.sheetName.toLowerCase()) ?? null;
  }
  return scope;
}

function expandedCellTarget(args) {
  const base = parseScopedRange(args.copyToRange ?? args.range);
  if (!base || args.copyToRange || !Array.isArray(args.cells) || args.cells.length === 0) return base;
  const height = args.cells.length;
  const width = args.cells[0]?.length ?? 1;
  if (height <= 1 && width <= 1) return base;
  return {
    ...base,
    address: `${columnLetters(base.startColumn)}${base.startRow}:${columnLetters(base.startColumn + width - 1)}${base.startRow + height - 1}`,
    endColumn: base.startColumn + width - 1,
    endRow: base.startRow + height - 1,
  };
}

function destructiveTarget(toolName, args) {
  switch (toolName) {
    case "excel_set_cell_range": {
      const styleOnly = args.cells?.flat?.().some((cell) => cell?.cellStyles || cell?.borderStyles || cell?.note);
      return args.allow_overwrite || styleOnly ? expandedCellTarget(args) : null;
    }
    case "excel_copy_to":
      return args.allow_overwrite ? parseScopedRange(args.destinationRange) : null;
    case "excel_clear_cell_range":
      return parseScopedRange(args.range);
    case "excel_set_format":
    case "excel_sort_range":
      return parseScopedRange(args.address);
    case "excel_resize_range":
      return args.range ? parseScopedRange(args.range) : null;
    default:
      return null;
  }
}

function sheetMatches(allowed, target, args) {
  if (allowed.sheetId != null && args.sheetId != null) return allowed.sheetId === args.sheetId;
  const targetSheet = target.sheetName ?? args.sheet ?? null;
  if (allowed.sheetName && targetSheet) {
    return allowed.sheetName.toLowerCase() === String(targetSheet).toLowerCase();
  }
  return allowed.sheetId == null && !allowed.sheetName;
}

function contains(outer, inner) {
  return (
    outer.startColumn <= inner.startColumn && outer.endColumn >= inner.endColumn &&
    outer.startRow <= inner.startRow && outer.endRow >= inner.endRow
  );
}

export function assertMutationAuthorized(toolName, args, scope) {
  const target = destructiveTarget(toolName, args ?? {});
  if (!target) return;
  const allowed = scope?.ranges ?? [];
  if (allowed.some((range) => sheetMatches(range, target, args ?? {}) && contains(range, target))) return;
  const ranges = allowed.map((range) => `${range.sheetName ? `${range.sheetName}!` : ""}${range.address}`).join(", ");
  const error = new Error(
    `Mutation target ${target.address} is outside this turn's authorized ranges${ranges ? ` (${ranges})` : ""}. ` +
      "Read the intended area and ask the user to name or select the target before changing existing content.",
  );
  error.code = "MUTATION_SCOPE_REQUIRED";
  error.commitStatus = "not_committed";
  throw error;
}

export function bindMutationSheet(toolName, args, scope) {
  if (args?.sheet || args?.sheetId != null) return args;
  if (!["excel_set_format", "excel_sort_range", "excel_autofilter", "excel_create_table"].includes(toolName)) {
    return args;
  }
  const selected = scope?.ranges?.find((range) => range.reason === "submitted_selection" && range.sheetName);
  return selected ? { ...args, sheet: selected.sheetName } : args;
}
