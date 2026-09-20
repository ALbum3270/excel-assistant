import { z } from "zod";

const MAX_CELLS = 200_000;
const MAX_EXAMPLES = 5;
const scalar = z.union([z.string().max(32767), z.number().finite(), z.boolean(), z.null()]);
const target = z.object({
  sheetId: z.number().int(),
  range: z
    .string()
    .regex(/^\$?[A-Za-z]+\$?[1-9]\d*(?::\$?[A-Za-z]+\$?[1-9]\d*)?$/)
    .describe(
      "Explicit local A1 range, excluding headers. Include the old output tail when checking for leftover rows.",
    ),
});
const common = { label: z.string().min(1).max(200), target };
const tolerance = z
  .number()
  .finite()
  .nonnegative()
  .optional()
  .describe("Absolute numeric tolerance; default 0. Text and numbers remain distinct.");
const check = z.discriminatedUnion("type", [
  z.object({ ...common, type: z.literal("row_count"), expected: z.number().int().nonnegative() }),
  z.object({ ...common, type: z.literal("unique") }),
  z.object({ ...common, type: z.literal("not_blank") }),
  z.object({ ...common, type: z.literal("formulas") }),
  z.object({
    ...common,
    type: z.literal("matches"),
    expected: z.array(z.array(scalar).min(1)).min(1),
    tolerance,
  }),
  z.object({ ...common, type: z.literal("sum"), expected: z.number().finite(), tolerance }),
  z.object({ ...common, type: z.literal("rows_in_source"), source: target }),
  z.object({ ...common, type: z.literal("same_rows"), source: target }),
]);

export const TASK_CHECK_SCHEMA = {
  action: z.enum(["define", "run"]),
  checks: z
    .array(check)
    .min(1)
    .max(20)
    .optional()
    .describe("Required for define; run reuses the declared plan."),
};

export const TASK_CHECK_DESCRIPTION =
  "Declare and run read-only task-result checks against live Excel data. Before changing data, " +
  "call action=define with checks derived from the user's requirements; source ranges are snapshotted then. " +
  "After writing, call action=run to see failures and fix the result before reporting. Checks are per user turn; " +
  "the daemon also reruns the declared plan at normal turn completion. row_count counts rows with any nonblank value; " +
  "unique checks whole-row tuples (choose the key columns as target); not_blank checks every value; formulas requires " +
  "a formula in every target cell; matches compares the full target to an independently computed expected matrix; " +
  "sum checks a numeric total and fails on nonnumeric nonblank cells; rows_in_source checks whole-row membership " +
  "in the saved source; same_rows checks the same row multiset, preserving duplicate counts but ignoring order. " +
  "Row comparisons ignore entirely blank rows. Comparisons are case-sensitive and type-preserving. " +
  "Exclude headers; include old output tails in row_count checks to catch leftover data. " +
  "Use independent source-based expectations, not a copy of the actual output. " +
  "Passing checks proves only the declared constraints. Total declared ranges are limited to 200000 cells; " +
  "use excel_bash for larger or custom assertions. This tool never changes the workbook.";

function columnName(index) {
  let text = "";
  for (let n = index; n > 0; n = Math.floor((n - 1) / 26))
    text = String.fromCharCode(65 + ((n - 1) % 26)) + text;
  return text;
}

function rangeShape(range) {
  const match = /^([A-Z]+)([1-9]\d*)(?::([A-Z]+)([1-9]\d*))?$/.exec(
    range.replaceAll("$", "").toUpperCase(),
  );
  if (!match) throw new Error(`Use a bounded local A1 range: ${range}`);
  const column = (s) => [...s].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const col = column(match[1]),
    row = Number(match[2]);
  const lastCol = column(match[3] ?? match[1]),
    lastRow = Number(match[4] ?? match[2]);
  if (lastCol < col || lastRow < row || lastCol > 16384 || lastRow > 1048576)
    throw new Error(`Invalid Excel range: ${range}`);
  return { row, col, rows: lastRow - row + 1, columns: lastCol - col + 1 };
}

const keyFor = ({ sheetId, range }) => `${sheetId}:${range.replaceAll("$", "").toUpperCase()}`;
const blank = (value) => value === "" || value === null || value === undefined;
const nonemptyRows = (data) =>
  data.values
    .map((values, i) => ({ values, row: data.row + i }))
    .filter(({ values }) => values.some((value) => !blank(value)));
const rowKey = (values) => JSON.stringify(values.map((value) => (blank(value) ? "" : value)));
const preview = (value) =>
  typeof value === "string" && value.length > 160 ? value.slice(0, 160) + "…" : value;
const equal = (a, b, tolerance = 0) =>
  (blank(a) && blank(b)) ||
  (typeof a === "number" && typeof b === "number" ? Math.abs(a - b) <= tolerance : a === b);

async function readRange(call, spec) {
  const shape = rangeShape(spec.range);
  const cells = {},
    formulas = {};
  let remaining = [spec.range];
  const cursors = new Set();
  while (remaining.length) {
    const cursor = JSON.stringify(remaining);
    if (cursors.has(cursor)) throw new Error("Range read did not advance; checks are incomplete.");
    cursors.add(cursor);
    const page = await call("excel_get_cell_ranges", {
      sheetId: spec.sheetId,
      ranges: remaining,
      includeStyles: false,
      cellLimit: 5000,
    });
    if (page?.success !== true || !page.worksheet?.cells)
      throw new Error("Range read failed; checks are incomplete.");
    Object.assign(cells, page.worksheet.cells);
    Object.assign(formulas, page.worksheet.formulas);
    if (page.hasMore && !page.remainingRanges?.length)
      throw new Error("Range read was truncated without a continuation.");
    remaining = page.hasMore ? page.remainingRanges : [];
  }
  const values = Array.from({ length: shape.rows }, (_, r) =>
    Array.from(
      { length: shape.columns },
      (_, c) => cells[`${columnName(shape.col + c)}${shape.row + r}`] ?? "",
    ),
  );
  return { ...shape, values, formulas };
}

function evaluate(check, data, source) {
  const result = {
    label: check.label,
    type: check.type,
    target: check.target,
    status: "passed",
    failedCount: 0,
    examples: [],
  };
  const fail = (example) => {
    result.status = "failed";
    result.failedCount++;
    if (result.examples.length < MAX_EXAMPLES) result.examples.push(example);
  };
  const address = (r, c) => `${columnName(data.col + c)}${data.row + r}`;
  const rows = nonemptyRows(data);
  if (check.type === "row_count") {
    result.expected = check.expected;
    result.actual = rows.length;
    if (rows.length !== check.expected) fail({ expected: check.expected, actual: rows.length });
  } else if (check.type === "unique") {
    const seen = new Map();
    for (const { values, row } of rows) {
      const key = rowKey(values);
      if (seen.has(key)) fail({ row, duplicateOfRow: seen.get(key) });
      else seen.set(key, row);
    }
  } else if (check.type === "rows_in_source" || check.type === "same_rows") {
    const counts = new Map();
    for (const { values } of nonemptyRows(source)) {
      const key = rowKey(values);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    for (const { values, row } of rows) {
      const key = rowKey(values),
        count = counts.get(key) ?? 0;
      if (!count)
        fail({ row, reason: "Row absent from source snapshot or repeated too many times" });
      else if (check.type === "same_rows") counts.set(key, count - 1);
    }
    if (check.type === "same_rows") {
      const missingRows = [...counts.values()].reduce((sum, n) => sum + n, 0);
      if (missingRows) {
        fail({ missingSourceRows: missingRows });
        result.failedCount += missingRows - 1;
      }
    }
  } else {
    let sum = 0;
    for (let r = 0; r < data.rows; r++)
      for (let c = 0; c < data.columns; c++) {
        const value = data.values[r][c],
          cell = address(r, c);
        if (check.type === "not_blank" && blank(value))
          fail({ address: cell, reason: "Blank value" });
        if (check.type === "formulas" && !data.formulas[cell])
          fail({ address: cell, reason: "No formula" });
        if (check.type === "matches" && !equal(value, check.expected[r][c], check.tolerance)) {
          fail({ address: cell, expected: preview(check.expected[r][c]), actual: preview(value) });
        }
        if (check.type === "sum" && !blank(value)) {
          if (typeof value !== "number")
            fail({ address: cell, reason: "Non-numeric value", actual: preview(value) });
          else sum += value;
        }
      }
    if (check.type === "sum") {
      result.expected = check.expected;
      result.actual = sum;
      if (!equal(sum, check.expected, check.tolerance))
        fail({ expected: check.expected, actual: sum });
    }
  }
  return result;
}

// Per live SDK session, reset for each user turn. Only declared constraints are
// evaluated: never infer requirements from text or pretend they cover all intent.
export function createTaskVerification(call) {
  let state;
  const reset = () => {
    state = { plan: null, sources: new Map(), mutations: 0 };
  };
  reset();
  return {
    reset,
    markMutation() {
      state.mutations++;
    },
    async define(rawChecks) {
      const checks = z.array(check).min(1).max(20).parse(rawChecks);
      const ranges = new Map();
      for (const item of checks) {
        const size = rangeShape(item.target.range);
        ranges.set(keyFor(item.target), size.rows * size.columns);
        if (
          item.type === "matches" &&
          (item.expected.length !== size.rows ||
            item.expected.some((row) => row.length !== size.columns))
        ) {
          throw new Error(
            `Expected matrix for "${item.label}" must match ${size.rows} x ${size.columns}.`,
          );
        }
        if (item.source) {
          const sourceSize = rangeShape(item.source.range);
          if (sourceSize.columns !== size.columns)
            throw new Error(
              `Source and target for "${item.label}" must have the same number of columns.`,
            );
          ranges.set(keyFor(item.source), sourceSize.rows * sourceSize.columns);
        }
      }
      if ([...ranges.values()].reduce((sum, cells) => sum + cells, 0) > MAX_CELLS) {
        throw new Error(
          `Task checks exceed ${MAX_CELLS} cells. Use smaller ranges or independent assertions in excel_bash.`,
        );
      }
      const current = state;
      const sources = new Map();
      for (const item of checks)
        if (item.source && !sources.has(keyFor(item.source))) {
          sources.set(keyFor(item.source), await readRange(call, item.source));
        }
      current.plan = checks;
      current.sources = sources;
      current.definedBeforeChanges = current.mutations === 0;
      return {
        status: "defined",
        checkCount: checks.length,
        definedBeforeChanges: current.definedBeforeChanges,
      };
    },
    async run() {
      const current = state;
      if (!current.plan)
        return {
          status: "not_checked",
          reason: "No task-result checks were declared for this turn.",
          checks: [],
        };
      const reads = new Map(),
        results = [];
      for (const item of current.plan) {
        const key = keyFor(item.target);
        if (!reads.has(key)) reads.set(key, readRange(call, item.target));
        try {
          results.push(
            evaluate(
              item,
              await reads.get(key),
              item.source && current.sources.get(keyFor(item.source)),
            ),
          );
        } catch (error) {
          results.push({
            label: item.label,
            type: item.type,
            target: item.target,
            status: "incomplete",
            reason: error.message,
          });
        }
      }
      const passed = results.filter((item) => item.status === "passed").length;
      return {
        status: results.some((item) => item.status === "failed")
          ? "failed"
          : passed === results.length
            ? "passed"
            : "incomplete",
        definedBeforeChanges: current.definedBeforeChanges,
        scope: "Declared constraints only; this is not proof of the complete task intent.",
        checkedAt: new Date().toISOString(),
        passed,
        total: results.length,
        checks: results,
      };
    },
    async finish() {
      return state.plan || state.mutations ? this.run() : null;
    },
  };
}
