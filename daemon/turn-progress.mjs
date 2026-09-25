// Evidence from executed Office calls, scoped to one user turn. It does not
// infer the user's intended answer or treat a commit receipt as semantic proof.
function bounds(address) {
  const local = String(address ?? "")
    .split("!")
    .at(-1)
    .replaceAll("$", "");
  const match = /^([A-Z]+)(\d+)(?::([A-Z]+)(\d+))?$/i.exec(local);
  if (!match) return null;
  const col = (label) =>
    [...label.toUpperCase()].reduce((n, char) => n * 26 + char.charCodeAt(0) - 64, 0);
  return {
    left: col(match[1]),
    top: Number(match[2]),
    right: col(match[3] ?? match[1]),
    bottom: Number(match[4] ?? match[2]),
  };
}

function covers(read, written) {
  const a = bounds(read);
  const b = bounds(written);
  return a && b && a.left <= b.left && a.top <= b.top && a.right >= b.right && a.bottom >= b.bottom;
}

export function createTurnProgress() {
  let writes = 0;
  let targets = [];
  let stopNudged = false;

  return {
    reset() {
      writes = 0;
      targets = [];
      stopNudged = false;
    },
    record(name, args, result, isWrite) {
      if (isWrite && result?.success !== false && result?.commitStatus !== "not_committed") {
        writes += 1;
        targets.push({
          sheetId: args?.sheetId,
          range:
            result?.writtenRange ??
            args?.copyToRange ??
            args?.range ??
            result?.address ??
            args?.address,
          read: false,
          asserted: false,
        });
      } else if (
        writes > 0 &&
        result?.success !== false &&
        result?.hasMore !== true &&
        !result?.remainingRanges?.length
      ) {
        for (const target of targets) {
          if (target.read) continue;
          if (name === "excel_get_workbook_metadata" && !target.range) target.read = true;
          if (target.sheetId == null || args?.sheetId !== target.sheetId || !target.range) continue;
          if (
            name === "excel_get_cell_ranges" &&
            (Array.isArray(args.ranges) ? args.ranges : [args.ranges]).some((range) =>
              covers(range, target.range),
            )
          )
            target.read = true;
          if (name === "excel_get_range_as_csv" && covers(args.range, target.range))
            target.read = true;
        }
      }
    },
    recordAssertion({ sheetId, range }) {
      for (const target of targets) {
        if (target.sheetId === sheetId && target.range && covers(range, target.range)) {
          target.read = true;
          target.asserted = true;
        }
      }
    },
    recordRead({ sheetId, range }) {
      for (const target of targets) {
        if (target.sheetId === sheetId && target.range && covers(range, target.range)) {
          target.read = true;
        }
      }
    },
    stopCheck(input) {
      const unread = targets.filter((target) => !target.read);
      if (input?.stop_hook_active || stopNudged || unread.length === 0) return {};
      stopNudged = true;
      return {
        decision: "block",
        reason:
          `${unread.length} workbook change(s) have no subsequent target read. ` +
          `Start with ${unread[0].range ?? "workbook metadata"}. ` +
          "Read the target from Excel, check its first and last cells and any required counts or constraints against the user's request. If you cannot verify correctness, say what remains unverified.",
      };
    },
    summary() {
      const unreadTargets = targets.filter((target) => !target.read).length;
      const assertedTargets = targets.filter((target) => target.asserted).length;
      return {
        writes,
        readTargets: targets.length - unreadTargets,
        unreadTargets,
        assertedTargets,
        status:
          writes === 0
            ? "no_write"
            : unreadTargets > 0
              ? "write_unchecked"
              : assertedTargets === targets.length
                ? "expected_values_matched"
                : "read_after_write",
      };
    },
  };
}
