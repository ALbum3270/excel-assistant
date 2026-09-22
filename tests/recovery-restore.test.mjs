import { test } from "node:test";
import assert from "node:assert/strict";
import { WorkbookRecoveryLog } from "../taskpane/shared/vendor/pi-recovery.js";

function memoryLog(store = new Map()) {
  return new WorkbookRecoveryLog({
    settings: {
      get: async (key) => store.get(key) ?? null,
      set: async (key, value) => void store.set(key, value),
    },
    getWorkbookContext: async () => ({ workbookId: "url_sha256:book", workbookName: "Book.xlsx", source: "document.url" }),
    getDocumentInstance: () => ({ read: () => "doc-1", ensure: async () => "doc-1" }),
    // Restore writes the saved values and reports what was there, which becomes
    // the inverse snapshot.
    applySnapshot: async () => ({ values: [["now"]], formulas: [["now"]] }),
  });
}

test("grouped inverses keep their identity and application order after reloading storage", async () => {
  const store = new Map();
  const log = memoryLog(store);
  for (const address of ["Sheet1!A1", "Sheet1!B1"]) {
    await log.append({
      toolName: "write_cells",
      toolCallId: "call-1",
      address,
      changedCount: 1,
      beforeValues: [["before"]],
      beforeFormulas: [["before"]],
    });
  }
  const originals = await log.listForCurrentWorkbook(10);
  assert.equal(originals.length, 2);

  const restoreCallId = "restore:call-1:x";
  for (const [index, snapshot] of originals.entries()) {
    await log.restore(snapshot.id, { toolCallId: restoreCallId, restoreOrder: originals.length - index - 1 });
  }

  const reloaded = memoryLog(store);
  const inverses = (await reloaded.listForCurrentWorkbook(10)).filter((item) => item.restoredFromSnapshotId);
  assert.equal(inverses.length, 2);
  assert.deepEqual([...new Set(inverses.map((item) => item.toolCallId))], [restoreCallId]);
  for (const [index, original] of originals.entries()) {
    assert.equal(inverses.find((item) => item.restoredFromSnapshotId === original.id).restoreOrder, originals.length - index - 1);
  }
});

test("without an id, restore keeps upstream's per-snapshot naming", async () => {
  const log = memoryLog();
  await log.append({
    toolName: "write_cells",
    toolCallId: "call-2",
    address: "Sheet1!C1",
    changedCount: 1,
    beforeValues: [["before"]],
    beforeFormulas: [["before"]],
  });
  const [snapshot] = await log.listForCurrentWorkbook(10);
  await log.restore(snapshot.id);
  const [inverse] = (await log.listForCurrentWorkbook(10)).filter((item) => item.restoredFromSnapshotId);
  assert.equal(inverse.toolCallId, `restore:${snapshot.id}`);
});
