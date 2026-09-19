/* global Excel, Office */

import {
  MAX_RECOVERY_CELLS,
  WorkbookRecoveryLog,
  captureFormatCellsState,
} from "./vendor/pi-recovery.js";

const RECOVERY_DB = "draftspect-recovery";
const RECOVERY_STORE = "settings";
const DOCUMENT_TOKEN_KEY = "draftspect-recovery-document-token-v1";

const CELL_FORMAT_PROPERTIES = Object.freeze({
  numberFormat: true,
  fillColor: true,
  fontColor: true,
  bold: true,
  italic: true,
  underlineStyle: true,
  fontName: true,
  fontSize: true,
  horizontalAlignment: true,
  verticalAlignment: true,
  wrapText: true,
  borderTop: true,
  borderBottom: true,
  borderLeft: true,
  borderRight: true,
  borderInsideHorizontal: true,
  borderInsideVertical: true,
});

let recoveryDatabasePromise;
function recoveryDatabase() {
  recoveryDatabasePromise ??= new Promise((resolve, reject) => {
    const request = indexedDB.open(RECOVERY_DB, 1);
    request.onupgradeneeded = () => request.result.createObjectStore(RECOVERY_STORE);
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
  return recoveryDatabasePromise;
}

const recoverySettings = {
  async get(key) {
    const db = await recoveryDatabase();
    return new Promise((resolve, reject) => {
      const request = db.transaction(RECOVERY_STORE).objectStore(RECOVERY_STORE).get(key);
      request.onsuccess = () => resolve(request.result ?? null);
      request.onerror = () => reject(request.error);
    });
  },
  async set(key, value) {
    const db = await recoveryDatabase();
    return new Promise((resolve, reject) => {
      const transaction = db.transaction(RECOVERY_STORE, "readwrite");
      transaction.objectStore(RECOVERY_STORE).put(value, key);
      transaction.oncomplete = () => resolve();
      transaction.onerror = () => reject(transaction.error);
    });
  },
};

function documentTokenIdentity() {
  const settings = Office?.context?.document?.settings;
  return {
    read() {
      return settings?.get(DOCUMENT_TOKEN_KEY) || null;
    },
    async ensure() {
      const existing = settings?.get(DOCUMENT_TOKEN_KEY);
      if (existing) return existing;
      if (!settings) return null;
      const token = globalThis.crypto?.randomUUID?.() ?? `doc_${Date.now().toString(36)}_${Math.random().toString(36).slice(2)}`;
      settings.set(DOCUMENT_TOKEN_KEY, token);
      await new Promise((resolve) => settings.saveAsync(() => resolve()));
      return token;
    },
  };
}

function workbookName(rawUrl) {
  if (!rawUrl) return null;
  const clean = String(rawUrl).split(/[?#]/, 1)[0].replaceAll("\\", "/");
  const tail = clean.split("/").at(-1);
  if (!tail) return null;
  try {
    return decodeURIComponent(tail);
  } catch {
    return tail;
  }
}

async function sha256Hex(value) {
  const bytes = new TextEncoder().encode(value);
  if (globalThis.crypto?.subtle?.digest) {
    const digest = await globalThis.crypto.subtle.digest("SHA-256", bytes);
    return [...new Uint8Array(digest)].map((byte) => byte.toString(16).padStart(2, "0")).join("");
  }
  let hash = 0x811c9dc5;
  for (const byte of bytes) {
    hash ^= byte;
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

async function currentWorkbookContext() {
  const rawUrl = Office?.context?.document?.url?.trim?.() || null;
  if (!rawUrl) return { workbookId: null, workbookName: null, source: "unknown" };
  const normalized = rawUrl.split(/[?#]/, 1)[0].replaceAll("\\", "/");
  return {
    workbookId: `url_sha256:${await sha256Hex(normalized)}`,
    workbookName: workbookName(rawUrl),
    source: "document.url",
  };
}

const recoveryLog = new WorkbookRecoveryLog({
  settings: recoverySettings,
  getWorkbookContext: currentWorkbookContext,
  getDocumentInstance: documentTokenIdentity,
});

function splitSheetAddress(address) {
  const raw = String(address ?? "").trim();
  if (!raw) return { sheetName: null, a1: null };
  if (raw.startsWith("'")) {
    for (let index = 1; index < raw.length; index += 1) {
      if (raw[index] !== "'") continue;
      if (raw[index + 1] === "'") {
        index += 1;
        continue;
      }
      if (raw[index + 1] === "!") {
        return { sheetName: raw.slice(1, index).replaceAll("''", "'"), a1: raw.slice(index + 2) };
      }
      break;
    }
  }
  const bang = raw.indexOf("!");
  return bang > 0
    ? { sheetName: raw.slice(0, bang), a1: raw.slice(bang + 1) }
    : { sheetName: null, a1: raw };
}

function refreshDocumentSettings() {
  const settings = Office?.context?.document?.settings;
  if (!settings?.refreshAsync) return Promise.resolve();
  return new Promise((resolve) => settings.refreshAsync(() => resolve()));
}

async function resolveWorksheet(context, { sheetId, sheet, address }) {
  const parsed = splitSheetAddress(address);
  const requestedName = parsed.sheetName ?? sheet ?? null;
  const worksheets = context.workbook.worksheets;
  worksheets.load("items");
  await context.sync();
  for (const item of worksheets.items) item.load("id,name");
  await context.sync();

  if (requestedName) {
    const match = worksheets.items.find((item) => item.name.toLocaleLowerCase() === requestedName.toLocaleLowerCase());
    if (!match) throw new Error(`Worksheet '${requestedName}' not found.`);
    return { worksheet: match, a1: parsed.a1 };
  }

  if (sheetId !== undefined && sheetId !== null) {
    await refreshDocumentSettings();
    const stableMap = Office?.context?.document?.settings?.get("openexcel-sheet-id-map") ?? {};
    const match = worksheets.items.find(
      (item) => item.id === sheetId || Number(stableMap[item.id]) === Number(sheetId),
    );
    if (!match) throw new Error(`Worksheet with ID ${sheetId} not found.`);
    return { worksheet: match, a1: parsed.a1 };
  }

  return { worksheet: context.workbook.worksheets.getActiveWorksheet(), a1: parsed.a1 };
}

async function captureRangeSnapshot(target) {
  return Excel.run(async (context) => {
    const { worksheet, a1 } = await resolveWorksheet(context, target);
    let range = target.useUsedRange
      ? worksheet.getUsedRangeOrNullObject()
      : worksheet.getRange(a1);
    let rows = target.rows;
    let columns = target.columns;
    if (target.sizeFromAddress && !String(target.address).includes(":")) {
      const source = worksheet.getRange(target.sizeFromAddress);
      source.load("rowCount,columnCount");
      await context.sync();
      rows = source.rowCount;
      columns = source.columnCount;
    }
    range.load("address,rowCount,columnCount,isNullObject");
    await context.sync();
    if (range.isNullObject) range = worksheet.getRange("A1");
    if (rows && columns) {
      range = range.getCell(0, 0).getResizedRange(rows - 1, columns - 1);
    }
    range.load("address,rowCount,columnCount,values,formulas");
    await context.sync();
    const cellCount = range.rowCount * range.columnCount;
    if (cellCount > MAX_RECOVERY_CELLS) {
      throw new Error(`Recovery snapshot exceeds the ${MAX_RECOVERY_CELLS}-cell limit.`);
    }
    return {
      kind: "range",
      address: range.address,
      beforeValues: range.values,
      beforeFormulas: range.formulas,
      cellCount,
    };
  });
}

async function qualifiedAddress(target) {
  return Excel.run(async (context) => {
    const { worksheet, a1 } = await resolveWorksheet(context, target);
    let range = target.useUsedRange
      ? worksheet.getUsedRangeOrNullObject()
      : worksheet.getRange(a1);
    let rows = target.rows;
    let columns = target.columns;
    if (target.sizeFromAddress && !String(target.address).includes(":")) {
      const source = worksheet.getRange(target.sizeFromAddress);
      source.load("rowCount,columnCount");
      await context.sync();
      rows = source.rowCount;
      columns = source.columnCount;
    }
    range.load("address,isNullObject");
    await context.sync();
    if (!range.isNullObject) {
      if (rows && columns) {
        range = range.getCell(0, 0).getResizedRange(rows - 1, columns - 1);
        range.load("address");
        await context.sync();
      }
      return range.address;
    }
    const firstCell = worksheet.getRange("A1");
    firstCell.load("address");
    await context.sync();
    return firstCell.address;
  });
}

async function captureFormatSnapshot(target, selection = CELL_FORMAT_PROPERTIES) {
  const address = await qualifiedAddress(target);
  const captured = await captureFormatCellsState(address, selection, {
    maxCellCount: MAX_RECOVERY_CELLS,
  });
  if (!captured.supported || !captured.state) {
    throw new Error(captured.reason ?? "Format recovery is unavailable for this range.");
  }
  return { kind: "format", address, state: captured.state, cellCount: captured.state.cellCount };
}

function matrixShape(cells) {
  if (!Array.isArray(cells) || !Array.isArray(cells[0])) return {};
  return { rows: cells.length, columns: cells[0].length };
}

function formatSelectionFromArgs(args) {
  return {
    ...(args.number_format !== undefined ? { numberFormat: true } : {}),
    ...(args.bold !== undefined ? { bold: true } : {}),
    ...(args.italic !== undefined ? { italic: true } : {}),
    ...(args.font_size !== undefined ? { fontSize: true } : {}),
    ...(args.font_name !== undefined ? { fontName: true } : {}),
    ...(args.font_color !== undefined ? { fontColor: true } : {}),
    ...(args.fill_color !== undefined ? { fillColor: true } : {}),
    ...(args.border
      ? {
          borderTop: true,
          borderBottom: true,
          borderLeft: true,
          borderRight: true,
          borderInsideHorizontal: true,
          borderInsideVertical: true,
        }
      : {}),
  };
}

function recoveryPlan(name, args) {
  const bySheetId = (address, extra = {}) => ({ sheetId: args.sheetId, address, ...extra });
  const bySheetName = (address, extra = {}) => ({ sheet: args.sheet, address, ...extra });
  switch (name) {
    case "excel_set_cell_range": {
      const shape = matrixShape(args.cells);
      const plans = [
        { capture: "range", target: bySheetId(args.range, shape) },
        ...(args.copyToRange
          ? [
              { capture: "range", target: bySheetId(args.copyToRange) },
              { capture: "format", target: bySheetId(args.copyToRange, shape), selection: CELL_FORMAT_PROPERTIES },
            ]
          : []),
      ];
      if (args.cells?.some((row) => row.some((cell) => cell?.cellStyles || cell?.borderStyles))) {
        plans.push({ capture: "format", target: bySheetId(args.range, shape), selection: CELL_FORMAT_PROPERTIES });
      }
      if (args.resizeWidth || args.resizeHeight) {
        plans.push({
          capture: "format",
          target: bySheetId(args.range, shape),
          selection: {
            ...(args.resizeWidth ? { columnWidth: true } : {}),
            ...(args.resizeHeight ? { rowHeight: true } : {}),
          },
        });
      }
      if (args.cells?.some((row) => row.some((cell) => cell?.note))) {
        plans.push({ unsupported: "Cell notes are not included in automatic recovery yet." });
      }
      return plans;
    }
    case "excel_clear_cell_range": {
      const clearType = args.clearType ?? "contents";
      return [
        ...(clearType !== "formats" ? [{ capture: "range", target: bySheetId(args.range) }] : []),
        ...(clearType !== "contents"
          ? [{ capture: "format", target: bySheetId(args.range), selection: CELL_FORMAT_PROPERTIES }]
          : []),
      ];
    }
    case "excel_copy_to":
      return [
        {
          capture: "range",
          target: bySheetId(args.destinationRange, { sizeFromAddress: args.sourceRange }),
        },
        {
          capture: "format",
          target: bySheetId(args.destinationRange, { sizeFromAddress: args.sourceRange }),
          selection: CELL_FORMAT_PROPERTIES,
        },
      ];
    case "excel_set_format":
      return [{ capture: "format", target: bySheetName(args.address), selection: formatSelectionFromArgs(args) }];
    case "excel_sort_range":
      return [
        { capture: "range", target: bySheetName(args.address) },
        { capture: "format", target: bySheetName(args.address), selection: CELL_FORMAT_PROPERTIES },
      ];
    case "excel_resize_range":
      return [{
        capture: "format",
        target: bySheetId(args.range, { useUsedRange: !args.range }),
        selection: {
          ...(args.width !== undefined ? { columnWidth: true } : {}),
          ...(args.height !== undefined ? { rowHeight: true } : {}),
        },
      }];
    default:
      return [];
  }
}

export async function prepareMutationRecovery(name, args, toolCallId) {
  const plan = recoveryPlan(name, args ?? {});
  const captures = [];
  const failures = [];
  for (const item of plan) {
    if (item.unsupported) {
      failures.push(item.unsupported);
      continue;
    }
    try {
      captures.push(
        item.capture === "format"
          ? await captureFormatSnapshot(item.target, item.selection)
          : await captureRangeSnapshot(item.target),
      );
    } catch (error) {
      failures.push(error?.message ?? String(error));
    }
  }

  return async function commitRecovery(result) {
    const snapshots = [];
    for (const capture of captures) {
      try {
        const snapshot = capture.kind === "format"
          ? await recoveryLog.appendFormatCells({
              toolName: "format_cells",
              toolCallId,
              address: capture.address,
              changedCount: capture.cellCount,
              formatRangeState: capture.state,
            })
          : await recoveryLog.append({
              toolName: "write_cells",
              toolCallId,
              address: capture.address,
              changedCount: result?.cellsCommitted ?? capture.cellCount,
              beforeValues: capture.beforeValues,
              beforeFormulas: capture.beforeFormulas,
            });
        if (snapshot) snapshots.push(snapshot);
      } catch (error) {
        failures.push(error?.message ?? String(error));
      }
    }

    if (snapshots.length === 0) {
      return {
        status: "not_available",
        reason: failures[0] ?? "This operation does not yet support an automatic recovery checkpoint.",
      };
    }
    return {
      status: "checkpoint_created",
      snapshotIds: snapshots.map((snapshot) => snapshot.id),
      targets: snapshots.map((snapshot) => snapshot.address),
      ...(failures.length ? { partial: true, unavailableReasons: failures } : {}),
    };
  };
}

function compactSnapshotGroup(snapshots) {
  const snapshot = snapshots[0];
  return {
    id: snapshot.id,
    snapshotIds: snapshots.map((item) => item.id),
    createdAt: new Date(snapshot.at).toISOString(),
    operation: snapshot.toolName === "restore_snapshot" ? "restore" : snapshot.toolName,
    addresses: [...new Set(snapshots.map((item) => item.address))],
    changedCount: Math.max(...snapshots.map((item) => item.changedCount)),
    kinds: [...new Set(snapshots.map((item) => item.snapshotKind ?? "range_values"))],
    restoredFromSnapshotId: snapshot.restoredFromSnapshotId,
  };
}

function groupSnapshots(snapshots) {
  const groups = new Map();
  for (const snapshot of snapshots) {
    const key = snapshot.toolCallId || snapshot.id;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(snapshot);
  }
  return [...groups.values()];
}

async function resolveSnapshotGroup(snapshotId) {
  const snapshots = await recoveryLog.listForCurrentWorkbook(120);
  const anchor = snapshotId
    ? snapshots.find((snapshot) => snapshot.id === snapshotId)
    : snapshots[0];
  if (!anchor) throw new Error("No recovery checkpoint is available for this workbook.");
  return snapshots.filter((snapshot) => snapshot.toolCallId === anchor.toolCallId);
}

export async function workbookHistory({ action = "list", snapshot_id: snapshotId, limit = 20 } = {}) {
  switch (action) {
    case "list":
      return {
        success: true,
        action,
        snapshots: groupSnapshots(await recoveryLog.listForCurrentWorkbook(120))
          .slice(0, limit)
          .map(compactSnapshotGroup),
      };
    case "restore": {
      const group = await resolveSnapshotGroup(snapshotId);
      const restored = [];
      for (const snapshot of group) restored.push(await recoveryLog.restore(snapshot.id));
      return {
        success: true,
        action,
        commitStatus: "committed",
        restoredSnapshotIds: restored.map((item) => item.restoredSnapshotId),
        inverseSnapshotIds: restored.map((item) => item.inverseSnapshotId).filter(Boolean),
        addresses: [...new Set(restored.map((item) => item.address))],
        changedCount: Math.max(...restored.map((item) => item.changedCount)),
        recovery: {
          status: "checkpoint_created",
          snapshotIds: restored.map((item) => item.inverseSnapshotId).filter(Boolean),
        },
      };
    }
    case "delete": {
      if (!snapshotId) throw new Error("snapshot_id is required for delete.");
      const group = await resolveSnapshotGroup(snapshotId);
      const deleted = await Promise.all(group.map((snapshot) => recoveryLog.delete(snapshot.id)));
      return { success: deleted.every(Boolean), action, snapshotIds: group.map((snapshot) => snapshot.id) };
    }
    case "clear":
      return { success: true, action, removed: await recoveryLog.clearForCurrentWorkbook() };
    default:
      throw new Error(`Unsupported workbook history action: ${action}`);
  }
}
