import { test } from "node:test";
import assert from "node:assert/strict";
import { WorkbookRecoveryLog } from "../taskpane/shared/vendor/pi-recovery.js";
import { readFileSync } from "node:fs";
import vm from "node:vm";

function memoryLog(store = new Map()) {
  return new WorkbookRecoveryLog({
    settings: {
      get: async (key) => store.get(key) ?? null,
      set: async (key, value) => void store.set(key, value),
    },
    getWorkbookContext: async () => ({
      workbookId: "url_sha256:book",
      workbookName: "Book.xlsx",
      source: "document.url",
    }),
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
    await log.restore(snapshot.id, {
      toolCallId: restoreCallId,
      restoreOrder: originals.length - index - 1,
    });
  }

  const reloaded = memoryLog(store);
  const inverses = (await reloaded.listForCurrentWorkbook(10)).filter(
    (item) => item.restoredFromSnapshotId,
  );
  assert.equal(inverses.length, 2);
  assert.deepEqual([...new Set(inverses.map((item) => item.toolCallId))], [restoreCallId]);
  for (const [index, original] of originals.entries()) {
    assert.equal(
      inverses.find((item) => item.restoredFromSnapshotId === original.id).restoreOrder,
      originals.length - index - 1,
    );
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
  const [inverse] = (await log.listForCurrentWorkbook(10)).filter(
    (item) => item.restoredFromSnapshotId,
  );
  assert.equal(inverse.toolCallId, `restore:${snapshot.id}`);
});

test("retention preserves whole operations across both stores and during restore", async () => {
  const store = new Map();
  let log = memoryLog(store);
  for (const address of ["Sheet1!A1", "Sheet1!B1"]) {
    await log.append({
      toolName: "write_cells",
      toolCallId: "old",
      address,
      beforeValues: [["before"]],
      beforeFormulas: [["before"]],
    });
  }
  const original = await log.listForCurrentWorkbook(10);
  let custom = [
    {
      id: "custom-old",
      toolCallId: "old",
      at: 0,
      snapshotKind: "custom_state",
      address: "Sheet1!A1",
    },
  ];
  for (let i = 0; i < 119; i++) {
    await log.append({
      toolName: "write_cells",
      toolCallId: `new-${i}`,
      address: "Sheet1!C1",
      beforeValues: [[i]],
      beforeFormulas: [[i]],
    });
  }
  // Reload through the real codec too: storage must not silently drop member 121.
  log = memoryLog(store);
  assert.equal((await log.listForCurrentWorkbook(1000)).length, 121);
  const source = readFileSync(new URL("../taskpane/shared/recovery.js", import.meta.url), "utf8");
  const applied = [];
  const ctx = vm.createContext({
    recoveryLog: log,
    readCustomSnapshots: async () => custom,
    writeCustomSnapshots: async (next) => {
      custom = next;
    },
    restoreCustomSnapshot: async (snapshot, toolCallId, restoreOrder, restoreDepth) => {
      applied.push(snapshot.id);
      custom.unshift({
        ...snapshot,
        id: "custom-inverse",
        at: Date.now(),
        toolCallId,
        restoreOrder,
        restoreDepth,
        restoredFromSnapshotId: snapshot.id,
      });
      return {
        restoredSnapshotId: snapshot.id,
        inverseSnapshotId: "custom-inverse",
        address: snapshot.address,
        changedCount: 1,
      };
    },
  });
  vm.runInContext(
    source.slice(source.indexOf("function compactSnapshotGroup(")).replace("export async", "async"),
    ctx,
  );
  const result = await ctx.workbookHistory({ action: "restore", snapshot_id: original[0].id });
  assert.equal(result.restoredSnapshotIds.length, 3);
  assert.deepEqual(applied, ["custom-old"]);
  assert.equal(result.inverseSnapshotIds.length, 3);
  const listed = await ctx.workbookHistory({ action: "list", limit: 1000 });
  assert.equal(listed.snapshots.length, 120);
  assert.equal(
    (await log.listForCurrentWorkbook(1000)).some((s) => s.toolCallId === "old"),
    false,
  );
  assert.equal(
    custom.some((s) => s.toolCallId === "old"),
    false,
  );
  const reloaded = memoryLog(store);
  const inverses = (await reloaded.listForCurrentWorkbook(1000)).filter(
    (s) => s.restoredFromSnapshotId,
  );
  assert.equal(inverses.length, 2);
  assert.ok(inverses.every((s) => s.restoreDepth === 1));
});

test("redoing an undo made before restoreDepth existed records depth 2, not 1", async () => {
  const store = new Map();
  const log = memoryLog(store);
  await log.append({
    toolName: "write_cells",
    toolCallId: "change",
    address: "Sheet1!A1",
    beforeValues: [["before"]],
    beforeFormulas: [["before"]],
  });
  const [change] = await log.listForCurrentWorkbook(10);
  // An undo written by the previous release: it links to the change but
  // carries no depth of its own.
  await log.append({
    toolName: "restore_snapshot",
    toolCallId: "legacy-undo",
    address: "Sheet1!A1",
    beforeValues: [["after"]],
    beforeFormulas: [["after"]],
    restoredFromSnapshotId: change.id,
  });
  const legacy = (await log.listForCurrentWorkbook(10)).find((s) => s.toolCallId === "legacy-undo");
  assert.equal(legacy.restoreDepth, undefined);

  const source = readFileSync(new URL("../taskpane/shared/recovery.js", import.meta.url), "utf8");
  const ctx = vm.createContext({
    recoveryLog: log,
    readCustomSnapshots: async () => [],
    writeCustomSnapshots: async () => {},
  });
  vm.runInContext(
    source.slice(source.indexOf("function compactSnapshotGroup(")).replace("export async", "async"),
    ctx,
  );
  await ctx.workbookHistory({ action: "restore", snapshot_id: legacy.id });

  const redo = (await memoryLog(store).listForCurrentWorkbook(10)).find(
    (s) => s.restoredFromSnapshotId === legacy.id,
  );
  // Counted as a first undo, the redo would label the change "undone" while
  // it is back in the workbook, and the next click would undo it again.
  assert.equal(redo.restoreDepth, 2);
});
