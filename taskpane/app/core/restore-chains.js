// The restore-points page shows one row per change, not one row per click.
//
// Restoring a change writes an inverse checkpoint, so a change that has been
// undone and redone a few times leaves a trail of near-identical "restore"
// entries — which is what the log needs, but not what anyone reading the page
// wants to see. Every inverse records the snapshot it was made from
// (restoredFromSnapshotId), so the trail is folded back into the change it
// belongs to: one row, with Undo and Redo on it, exactly like the tool card.
//
// The links come from the log itself, so the page and the card agree even
// after a pane reload, and undoing from one is visible in the other.

const createdTime = (group) => Date.parse(group?.createdAt ?? "") || 0;

/**
 * Fold restore checkpoints into the change they undo.
 *
 * @param {Array<object>} snapshots compact snapshot groups, newest first
 * @returns {Array<object>} one row per change, each carrying:
 *   `chain` — every group's id, the change first and the newest restore last;
 *   `tipId` — the snapshot the next click must restore;
 *   `restoreCount` — how many times this change has been restored;
 *   `undone` — whether the change is currently rolled back;
 *   `lastRestoredAt` — when the last restore happened, or null.
 */
export function collapseRestoreChains(snapshots = []) {
  const groups = snapshots.filter(Boolean);

  // A restore names the individual snapshot it came from, which may be any
  // member of a multi-cell group, so look up members rather than group ids.
  const byMember = new Map();
  for (const group of groups) {
    const members = group.snapshotIds?.length ? group.snapshotIds : [group.id];
    for (const id of members) if (!byMember.has(id)) byMember.set(id, group);
  }

  const parents = new Map();
  const children = new Map();
  for (const group of groups) {
    const parent = group.restoredFromSnapshotId ? byMember.get(group.restoredFromSnapshotId) : null;
    // A restore whose change is no longer listed (pruned, deleted, or past the
    // page's limit) keeps its own row: there is nothing left to fold it into.
    if (!parent || parent === group) continue;
    parents.set(group, parent);
    if (!children.has(parent)) children.set(parent, []);
    children.get(parent).push(group);
  }

  const rows = [];
  for (const group of groups) {
    if (parents.has(group)) continue;
    const chain = [group];
    const seen = new Set(chain);
    let current = group;
    for (;;) {
      const next = (children.get(current) ?? [])
        // Restoring the same snapshot twice leaves two inverses; the newest is
        // the one the next click has to restore.
        .reduce(
          (newest, candidate) =>
            createdTime(candidate) >= createdTime(newest) ? candidate : newest,
          null,
        );
      if (!next || seen.has(next)) break;
      seen.add(next);
      chain.push(next);
      current = next;
    }
    const tip = chain[chain.length - 1];
    const restoreCount = chain.length - 1;
    rows.push({
      ...group,
      chain: chain.map((item) => item.id),
      tipId: tip.id,
      restoreCount,
      undone: restoreCount % 2 === 1,
      lastRestoredAt: restoreCount ? tip.createdAt : null,
    });
  }
  return rows;
}
