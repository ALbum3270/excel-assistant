import { test } from "node:test";
import assert from "node:assert/strict";
import { createTurnProgress } from "../daemon/turn-progress.mjs";

test("a write needs a subsequent complete target read before ending", () => {
  const progress = createTurnProgress();
  progress.record(
    "excel_set_cell_range",
    { sheetId: 1, range: "B2:B4" },
    { success: true, commitStatus: "committed", writtenRange: "B2:B4" },
    true,
  );
  assert.equal(progress.summary().status, "write_unchecked");
  assert.equal(progress.stopCheck({}).decision, "block");
  assert.deepEqual(progress.stopCheck({}), {}, "only one automatic reminder per turn");
  progress.record(
    "excel_get_cell_ranges",
    { sheetId: 1, ranges: ["A1:C5"] },
    { success: true, hasMore: false },
    false,
  );
  assert.equal(progress.summary().status, "read_after_write");
  assert.equal(progress.summary().unreadTargets, 0);
  progress.recordAssertion({ sheetId: 1, range: "A1:C5" });
  assert.equal(progress.summary().status, "expected_values_matched");
  progress.record(
    "excel_set_cell_range",
    { sheetId: 1, range: "D2:D4" },
    { success: true, commitStatus: "committed", writtenRange: "D2:D4" },
    true,
  );
  assert.equal(progress.summary().status, "write_unchecked");
  assert.equal(progress.summary().unreadTargets, 1);
  progress.reset();
  assert.equal(progress.summary().status, "no_write");
});

const committed = (writtenRange) => ({ success: true, commitStatus: "committed", writtenRange });
const readResult = (name = "Sheet1", sheetId = 1) => ({
  success: true,
  hasMore: false,
  remainingRanges: [],
  worksheet: { name, sheetId },
});

test("a write made by sheet name is checked by a read made by sheetId (eval task 38074)", () => {
  // Formula filled by sheetId, then formatted by sheet name, then read back twice.
  // Before the fix the format target had no sheetId and could never count as read.
  const progress = createTurnProgress();
  progress.record("excel_set_cell_range", { sheetId: 1, range: "K2:K5" }, committed("K2:K5"), true);
  progress.record(
    "excel_set_format",
    { address: "K2:K5", sheet: "Sheet1", number_format: "0" },
    { success: true, commitStatus: "committed", sheet: "Sheet1", address: "Sheet1!K2:K5" },
    true,
  );
  progress.record("excel_get_cell_ranges", { sheetId: 1, ranges: "K2:K5" }, readResult(), false);
  assert.equal(progress.summary().status, "read_after_write");
  assert.deepEqual(progress.stopCheck({}), {});
});

test("sort, filter and table writes by name are checked the same way", () => {
  const progress = createTurnProgress();
  progress.record(
    "excel_sort_range",
    { address: "A1:D40", sheet: "Data" },
    { success: true, commitStatus: "committed", sheet: "Data", address: "Data!A1:D40" },
    true,
  );
  progress.record(
    "excel_get_cell_ranges",
    { sheetId: 2, ranges: ["A1:D40"] },
    readResult("Other", 2),
    false,
  );
  assert.equal(
    progress.summary().status,
    "write_unchecked",
    "a read of another sheet does not count",
  );
  progress.record(
    "excel_get_cell_ranges",
    { sheetId: 3, ranges: ["A1:D40"] },
    readResult("Data", 3),
    false,
  );
  assert.equal(progress.summary().status, "read_after_write");
});

test("first and last cells read as spot checks count, as the reminder asks (eval task 37462)", () => {
  const progress = createTurnProgress();
  progress.record(
    "excel_set_cell_range",
    { sheetId: 1, range: "F4:F244" },
    committed("F4:F244"),
    true,
  );
  progress.record(
    "excel_get_cell_ranges",
    { sheetId: 1, ranges: ["F4:F10", "F80:F90", "F240:F244"] },
    readResult(),
    false,
  );
  assert.equal(progress.summary().status, "read_after_write");
});

test("a few cells from the top of a large write do not count (eval task 455-35)", () => {
  const progress = createTurnProgress();
  progress.record(
    "excel_set_cell_range",
    { sheetId: 1, range: "A2:J10411" },
    committed("A2:J10411"),
    true,
  );
  progress.record("excel_get_cell_ranges", { sheetId: 1, ranges: "A2:A5" }, readResult(), false);
  assert.equal(progress.summary().status, "write_unchecked");
  assert.match(progress.stopCheck({}).reason, /A2:J10411/);
});

test("a bulk write sent in row chunks is one target (eval task 118-50)", () => {
  const progress = createTurnProgress();
  for (const range of ["A2:A2001", "A2002:A4001", "A4002:A6001"]) {
    progress.record("excel_set_cell_range", { sheetId: 1, range }, committed(range), true);
  }
  assert.equal(progress.summary().unreadTargets, 1);
  progress.record("excel_get_cell_ranges", { sheetId: 1, ranges: "A2:A2001" }, readResult(), false);
  assert.equal(
    progress.summary().status,
    "write_unchecked",
    "the block's last cell is still unread",
  );
  assert.match(progress.stopCheck({}).reason, /A2:A6001/);
  progress.record("excel_get_cell_ranges", { sheetId: 1, ranges: "A6001" }, readResult(), false);
  assert.equal(progress.summary().status, "read_after_write");
});

test("comma-separated read ranges in one string are split (eval task 15387)", () => {
  const progress = createTurnProgress();
  progress.record(
    "excel_set_cell_range",
    { sheetId: 1, range: "B13:F13" },
    committed("B13:F13"),
    true,
  );
  progress.record(
    "excel_get_cell_ranges",
    { sheetId: 1, ranges: "A13:A14, B13:F13" },
    readResult(),
    false,
  );
  assert.equal(progress.summary().status, "read_after_write");
});

test("rewriting cells after reading them makes them unchecked again, without a duplicate", () => {
  const progress = createTurnProgress();
  progress.record("excel_set_cell_range", { sheetId: 1, range: "K2:K5" }, committed("K2:K5"), true);
  progress.record("excel_get_cell_ranges", { sheetId: 1, ranges: ["K2:K5"] }, readResult(), false);
  progress.record("excel_set_cell_range", { sheetId: 1, range: "K2:K5" }, committed("K2:K5"), true);
  assert.equal(progress.summary().status, "write_unchecked");
  assert.equal(progress.summary().unreadTargets, 1);
});

test("the reminder names every unread target, not only the first", () => {
  const progress = createTurnProgress();
  progress.record("excel_set_cell_range", { sheetId: 1, range: "B2" }, committed("B2"), true);
  progress.record(
    "excel_set_format",
    { address: "D1:D9", sheet: "Sheet1" },
    { success: true, commitStatus: "committed", sheet: "Sheet1", address: "Sheet1!D1:D9" },
    true,
  );
  const { reason } = progress.stopCheck({});
  assert.match(reason, /2 workbook change/);
  assert.match(reason, /B2/);
  assert.match(reason, /Sheet1!D1:D9/);
});

test("a copy is checked at its destination, not treated as a write without a range (eval task 57989)", () => {
  const progress = createTurnProgress();
  progress.record(
    "excel_copy_to",
    { sheetId: 1, sourceRange: "B25:B43", destinationRange: "C25:H43", allow_overwrite: true },
    { success: true, commitStatus: "committed" },
    true,
  );
  assert.match(progress.stopCheck({}).reason, /C25:H43/);
  progress.record(
    "excel_get_cell_ranges",
    { sheetId: 1, ranges: ["C25:H43"] },
    readResult(),
    false,
  );
  assert.equal(progress.summary().status, "read_after_write");
});
