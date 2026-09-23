// Collapse consecutive calls to the same tool into one expandable block.
//
// The rules are pi-for-excel's src/ui/tool-grouping.ts (MIT): runs of two or more
// completed calls to the same tool group together, runs of three or more start
// collapsed, and a running call never groups so it always stays visible. pi does
// this on the DOM with a MutationObserver; here it is a pure pass over the
// timeline the view renders, and the view keeps what the user toggled by hand.

export const COLLAPSE_THRESHOLD = 3;

/**
 * @param {Array<{kind: string, id: string, local?: string, state?: string}>} items
 * @returns {Array<{kind: "item", item} | {kind: "group", id, toolName, items, collapsed}>}
 */
export function groupTimeline(items) {
  const out = [];
  let run = [];

  function flush() {
    if (run.length >= 2) {
      out.push({
        kind: "group",
        id: `group:${run[0].id}`,
        toolName: run[0].local,
        items: run,
        collapsed: run.length >= COLLAPSE_THRESHOLD,
      });
    } else {
      for (const item of run) out.push({ kind: "item", item });
    }
    run = [];
  }

  for (const item of items) {
    const groupable = item.kind === "tool" && item.local && item.state === "success";
    if (groupable && run.length > 0 && run[0].local === item.local) {
      run.push(item);
      continue;
    }
    flush();
    if (groupable) run.push(item);
    else out.push({ kind: "item", item });
  }
  flush();
  return out;
}
