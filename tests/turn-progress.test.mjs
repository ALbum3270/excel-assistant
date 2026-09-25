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
