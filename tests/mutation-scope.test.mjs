import { test } from "node:test";
import assert from "node:assert/strict";
import {
  assertMutationAuthorized,
  bindMutationSheet,
  buildMutationScope,
  parseScopedRange,
} from "../daemon/mutation-scope.mjs";

test("mutation scope binds answer_position and submitted selection to sheet IDs", () => {
  const scope = buildMutationScope(
    "### answer_position\n'Sales 2026'!B2:D10",
    { selection: { address: "'Sales 2026'!C3" } },
    [{ id: 7, name: "Sales 2026" }],
  );
  assert.equal(scope.ranges.length, 2);
  assert.ok(scope.ranges.every((range) => range.sheetId === 7));
  assert.doesNotThrow(() =>
    assertMutationAuthorized(
      "excel_set_cell_range",
      { sheetId: 7, range: "C4:D5", cells: [[1, 2], [3, 4]], allow_overwrite: true },
      scope,
    ),
  );
  assert.throws(
    () => assertMutationAuthorized("excel_clear_cell_range", { sheetId: 7, range: "E2:E3" }, scope),
    (error) => error.code === "MUTATION_SCOPE_REQUIRED" && error.commitStatus === "not_committed",
  );
});

test("formula sources and source-marked ranges do not grant edit authority", () => {
  const scope = buildMutationScope(
    "根据 A1:A10 填充 B1，公式 =SUM(A1:A10)",
    {},
    [{ id: 1, name: "Sheet1" }],
  );
  assert.deepEqual(scope.ranges.map((range) => range.address), ["B1"]);
});

test("a matrix starting at one cell is checked at its actual expanded size", () => {
  const scope = buildMutationScope("修改 B2", { selection: { address: "Sheet1!B2" } }, [{ id: 1, name: "Sheet1" }]);
  assert.throws(
    () =>
      assertMutationAuthorized(
        "excel_set_cell_range",
        { sheetId: 1, range: "B2", cells: [[1, 2], [3, 4]], allow_overwrite: true },
        scope,
      ),
    /target B2:C3 is outside/,
  );
});

test("legacy range tools are pinned to the submitted selection sheet", () => {
  const scope = buildMutationScope("格式化 B2", { selection: { address: "'My Sheet'!B2" } }, []);
  assert.deepEqual(bindMutationSheet("excel_set_format", { address: "B2", bold: true }, scope), {
    address: "B2",
    bold: true,
    sheet: "My Sheet",
  });
  assert.equal(parseScopedRange("A:D").endRow, 1_048_576);
  assert.equal(parseScopedRange("2:5").endColumn, 16_384);
});
