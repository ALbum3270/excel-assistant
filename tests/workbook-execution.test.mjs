import { test } from "node:test";
import assert from "node:assert/strict";
import { createWorkbookExecution } from "../daemon/workbook-execution.mjs";

const BOOK = "C:\\evals\\57989.xlsx";

// A formula write that Excel rejected; the pane reports how sure it is that
// nothing was written.
async function rejectedWrite(execution, commitStatus) {
  const error = await execution
    .run(BOOK, { write: true, toolName: "excel_fill_formula" }, async () => {
      throw Object.assign(new Error("invalid argument [at Range.formulas]"), {
        code: "InvalidArgument",
        commitStatus,
        executionSettled: true,
      });
    })
    .catch((e) => e);
  assert.equal(error.code, "InvalidArgument");
  return error.workbookRevision;
}

const nextWrite = (execution, expectedRevision) =>
  execution.run(
    BOOK,
    { write: true, expectedRevision, toolName: "excel_fill_formula" },
    async () => ({ success: true, commitStatus: "committed" }),
  );

test("a formula Excel rejected before writing leaves the next write free to go (eval task 57989)", async () => {
  const execution = createWorkbookExecution();
  const known = await rejectedWrite(execution, "not_committed");
  assert.equal(
    await rejectedWrite(execution, "not_committed"),
    known,
    "the workbook did not change",
  );
  const { result, revision } = await nextWrite(execution, known);
  assert.equal(result.commitStatus, "committed");
  assert.equal(revision, known + 1);
});

test("an outcome that may have changed the workbook still requires a reread", async () => {
  // What every rejected formula caused before: the next writes were refused
  // until the model read the sheet again.
  const execution = createWorkbookExecution();
  const known = await rejectedWrite(execution, "not_committed");
  assert.equal(await rejectedWrite(execution, "unknown"), known + 1);
  const refused = await nextWrite(execution, known).catch((e) => e);
  assert.equal(refused.code, "STALE_WORKBOOK_REVISION");
});
