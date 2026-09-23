import { test } from "node:test";
import assert from "node:assert/strict";
import { collapseRestoreChains } from "../taskpane/app/core/restore-chains.js";

// The shape the restore-points page gets back from workbook_history "list":
// compact snapshot groups, newest first.
function group(id, at, extra = {}) {
  return {
    id,
    snapshotIds: [id],
    createdAt: new Date(at).toISOString(),
    operation: "excel_set_cell_range",
    addresses: ["Sales!J10:J11"],
    changedCount: 4,
    kinds: ["range_values"],
    ...extra,
  };
}

const restore = (id, at, from) =>
  group(id, at, { operation: "restore", restoredFromSnapshotId: from });

test("a change that was never restored is one row, ready to undo", () => {
  const [row, ...rest] = collapseRestoreChains([group("s1", 1000)]);
  assert.equal(rest.length, 0);
  assert.equal(row.tipId, "s1");
  assert.equal(row.undone, false);
  assert.equal(row.restoreCount, 0);
  assert.equal(row.lastRestoredAt, null);
});

test("undo, redo and undo again stay one row and never add a second", () => {
  // What four clicks on the page used to leave behind: the change plus three
  // reverse checkpoints, each made from the previous one.
  const snapshots = [
    restore("s4", 4000, "s3"),
    restore("s3", 3000, "s2"),
    restore("s2", 2000, "s1"),
    group("s1", 1000),
  ];
  const rows = collapseRestoreChains(snapshots);
  assert.equal(rows.length, 1, "one row per change");
  const [row] = rows;
  assert.deepEqual(row.chain, ["s1", "s2", "s3", "s4"]);
  assert.equal(row.tipId, "s4", "the next click restores the newest checkpoint");
  assert.equal(row.restoreCount, 3);
  assert.equal(row.undone, true, "an odd number of restores leaves the change rolled back");
  assert.equal(row.lastRestoredAt, new Date(4000).toISOString());
  assert.equal(row.operation, "excel_set_cell_range", "the row still describes the change");
});

test("an even number of restores puts the change back in effect", () => {
  const rows = collapseRestoreChains([restore("s2", 2000, "s1"), group("s1", 1000)]);
  assert.equal(rows[0].undone, true);
  const redone = collapseRestoreChains([
    restore("s3", 3000, "s2"),
    restore("s2", 2000, "s1"),
    group("s1", 1000),
  ]);
  assert.equal(redone[0].undone, false);
  assert.equal(redone[0].tipId, "s3");
});

test("a restore is folded into its change even when it names another group member", () => {
  // A multi-cell change is one group of several snapshots; the reverse
  // checkpoint names the individual snapshot it came from.
  const change = group("a1", 1000, { snapshotIds: ["a1", "a2", "a3"] });
  const rows = collapseRestoreChains([restore("r1", 2000, "a3"), change]);
  assert.equal(rows.length, 1);
  assert.deepEqual(rows[0].chain, ["a1", "r1"]);
  assert.equal(rows[0].tipId, "r1");
});

test("separate changes keep separate rows, in the order they were listed", () => {
  const rows = collapseRestoreChains([
    restore("s4", 4000, "s3"),
    group("s3", 3000, { addresses: ["Sales!H1"] }),
    restore("s2", 2000, "s1"),
    group("s1", 1000),
  ]);
  assert.deepEqual(
    rows.map((row) => row.id),
    ["s3", "s1"],
  );
  assert.deepEqual(
    rows.map((row) => row.undone),
    [true, true],
  );
});

test("a restore whose change is no longer listed keeps its own row", () => {
  // History pruning or the page's limit can cut the change off; the checkpoint
  // is still the only way back, so it must stay visible.
  const rows = collapseRestoreChains([restore("r9", 9000, "gone")]);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].id, "r9");
  assert.equal(rows[0].tipId, "r9");
  assert.equal(rows[0].undone, false);
});

test("restoring the same change twice follows the newer checkpoint", () => {
  // Undoing from the tool card and from this page both restore the same
  // snapshot, which leaves two reverse checkpoints pointing at it.
  const rows = collapseRestoreChains([
    restore("late", 5000, "s1"),
    restore("early", 2000, "s1"),
    group("s1", 1000),
  ]);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].tipId, "late");
  assert.deepEqual(rows[0].chain, ["s1", "early", "late"]);
  assert.equal(rows[0].restoreCount, 2);
});

test("an inverse retains its redo state after its parent is pruned", () => {
  const [row] = collapseRestoreChains([{ ...restore("r", 2, "gone"), restoreDepth: 1 }]);
  assert.equal(row.undone, true);
});

test("a checkpoint that points at itself does not loop", () => {
  const rows = collapseRestoreChains([group("s1", 1000, { restoredFromSnapshotId: "s1" })]);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].tipId, "s1");
});

test("an empty or missing list is an empty page", () => {
  assert.deepEqual(collapseRestoreChains([]), []);
  assert.deepEqual(collapseRestoreChains(), []);
  assert.deepEqual(collapseRestoreChains([null, undefined]), []);
});

test("a chain started before restoreDepth existed is counted from its links", () => {
  // Undo from the previous release (no depth), then a redo whose stored depth
  // an older build miscounted as 1: the change is back in effect either way.
  const [row] = collapseRestoreChains([
    { ...restore("s3", 3000, "s2"), restoreDepth: 1 },
    restore("s2", 2000, "s1"),
    group("s1", 1000),
  ]);
  assert.equal(row.tipId, "s3");
  assert.equal(row.undone, false);
});

test("a pruned change's row adds the hops below its first stored depth", () => {
  const [row] = collapseRestoreChains([
    restore("r2", 3000, "r1"),
    { ...restore("r1", 2000, "gone"), restoreDepth: 1 },
  ]);
  assert.equal(row.tipId, "r2");
  assert.equal(row.undone, false, "undo (depth 1) then redo (depth 2) leaves the change in effect");
});
