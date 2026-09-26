// Evidence from executed Office calls, scoped to one user turn. It does not
// infer the user's intended answer or treat a commit receipt as semantic proof.

// The A1 part of "A1", "B2:C9" or "'My Sheet'!B2:C9" as column/row bounds.
function bounds(address) {
  const local = String(address ?? "")
    .split("!")
    .at(-1)
    .replaceAll("$", "")
    .trim();
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

function columnLabel(n) {
  let label = "";
  for (let rest = n; rest > 0; rest = Math.floor((rest - 1) / 26)) {
    label = String.fromCharCode(65 + ((rest - 1) % 26)) + label;
  }
  return label;
}

const a1 = (box) =>
  box.left === box.right && box.top === box.bottom
    ? `${columnLabel(box.left)}${box.top}`
    : `${columnLabel(box.left)}${box.top}:${columnLabel(box.right)}${box.bottom}`;

// The sheet named in "Sheet1!A1" or "'It''s'!A1", if any.
function sheetOf(address) {
  const text = String(address ?? "");
  const bang = text.lastIndexOf("!");
  if (bang < 0) return undefined;
  const sheet = text.slice(0, bang).trim();
  return /^'.*'$/.test(sheet) ? sheet.slice(1, -1).replaceAll("''", "'") : sheet;
}

// "A1:B6, C1:C6" as separate ranges; a comma inside a quoted sheet name stays.
function splitRanges(ranges) {
  if (Array.isArray(ranges)) return ranges.map(String);
  const parts = [];
  let current = "";
  let quoted = false;
  for (const char of String(ranges ?? "")) {
    if (char === "'") quoted = !quoted;
    if (char === "," && !quoted) {
      parts.push(current);
      current = "";
    } else current += char;
  }
  return [...parts, current].map((part) => part.trim()).filter(Boolean);
}

const hasCell = (box, col, row) =>
  box.left <= col && col <= box.right && box.top <= row && row <= box.bottom;
const contains = (outer, inner) =>
  outer.left <= inner.left &&
  outer.top <= inner.top &&
  outer.right >= inner.right &&
  outer.bottom >= inner.bottom;

export function createTurnProgress() {
  let writes = 0;
  let targets = [];
  let stopNudged = false;
  // Cell tools address a sheet by numeric sheetId; format, sort, filter and
  // table tools by name. Read results carry both, which lets a read by id
  // check a write made by name.
  const sheetIds = new Map();

  function learn(name, id) {
    if (typeof name === "string" && Number.isInteger(id)) sheetIds.set(name, id);
  }

  function sameSheet(target, id, name) {
    if (target.sheetId !== undefined && target.sheetId === id) return true;
    if (target.sheetName === undefined) return false;
    return target.sheetName === name || (id !== undefined && sheetIds.get(target.sheetName) === id);
  }

  function addTarget(sheetId, sheetName, range) {
    const box = bounds(range);
    const same = (target) =>
      (sheetId !== undefined && target.sheetId === sheetId) ||
      (sheetName !== undefined && target.sheetName === sheetName);
    // Writing the same cells again makes them unchecked again; it is still one target.
    const existing =
      box && targets.find((target) => target.box && same(target) && a1(target.box) === a1(box));
    if (existing) {
      Object.assign(existing, { first: false, last: false, read: false, asserted: false });
      return;
    }
    // A bulk write arrives as consecutive row chunks of one block: keep them as
    // one target, so the block's own first and last cells stand for all of it.
    const previous = targets.at(-1);
    if (
      box &&
      previous?.box &&
      !previous.read &&
      same(previous) &&
      previous.box.left === box.left &&
      previous.box.right === box.right &&
      box.top === previous.box.bottom + 1
    ) {
      previous.box = { ...previous.box, bottom: box.bottom };
      previous.last = false;
      return;
    }
    targets.push({
      sheetId,
      sheetName,
      box,
      first: false,
      last: false,
      read: false,
      asserted: false,
    });
  }

  // A target counts as read once reads have shown its first and last cells —
  // what the reminder asks for — in one range or as separate spot checks.
  function markRead(id, name, ranges) {
    const boxes = splitRanges(ranges).map(bounds).filter(Boolean);
    for (const target of targets) {
      if (target.read || !target.box || !sameSheet(target, id, name)) continue;
      const { left, top, right, bottom } = target.box;
      if (boxes.some((box) => hasCell(box, left, top))) target.first = true;
      if (boxes.some((box) => hasCell(box, right, bottom))) target.last = true;
      if (target.first && target.last) target.read = true;
    }
  }

  const label = (target) =>
    target.box
      ? `${target.sheetName ? `${target.sheetName}!` : ""}${a1(target.box)}`
      : "workbook metadata";

  return {
    reset() {
      writes = 0;
      targets = [];
      stopNudged = false;
    },
    record(name, args, result, isWrite) {
      learn(result?.worksheet?.name, result?.worksheet?.sheetId);
      learn(result?.sheetName, args?.sheetId);
      if (isWrite && result?.success !== false && result?.commitStatus !== "not_committed") {
        writes += 1;
        const range =
          result?.writtenRange ??
          args?.copyToRange ??
          args?.destinationRange ??
          args?.range ??
          result?.address ??
          args?.address;
        addTarget(
          Number.isInteger(args?.sheetId) ? args.sheetId : undefined,
          result?.sheet ?? args?.sheet ?? sheetOf(range),
          range,
        );
        return;
      }
      if (
        writes === 0 ||
        result?.success === false ||
        result?.hasMore === true ||
        result?.remainingRanges?.length
      ) {
        return;
      }
      if (name === "excel_get_workbook_metadata") {
        for (const target of targets) if (!target.box) target.read = true;
      }
      if (name === "excel_get_cell_ranges") {
        markRead(args?.sheetId, result?.worksheet?.name, args?.ranges);
      }
      if (name === "excel_get_range_as_csv") {
        markRead(args?.sheetId, result?.sheetName, [args?.range]);
      }
    },
    recordAssertion({ sheetId, range }) {
      const box = bounds(range);
      for (const target of targets) {
        if (box && target.box && sameSheet(target, sheetId) && contains(box, target.box)) {
          target.read = true;
          target.asserted = true;
        }
      }
    },
    recordRead({ sheetId, range }) {
      markRead(sheetId, undefined, [range]);
    },
    stopCheck(input) {
      const unread = targets.filter((target) => !target.read);
      if (input?.stop_hook_active || stopNudged || unread.length === 0) return {};
      stopNudged = true;
      const listed = unread.slice(0, 5).map(label).join(", ");
      return {
        decision: "block",
        reason:
          `${unread.length} workbook change(s) have not been read back since they were made: ` +
          `${listed}${unread.length > 5 ? ", ..." : ""}. ` +
          "Read each target from Excel, at least its first and last cells, and check any required counts or constraints against the user's request. If you cannot verify correctness, say what remains unverified.",
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
