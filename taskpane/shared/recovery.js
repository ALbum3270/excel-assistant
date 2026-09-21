/* global Excel, Office */

import {
  MAX_RECOVERY_CELLS,
  WorkbookRecoveryLog,
  captureChartPresentState,
  captureCommentThreadState,
  captureFormatCellsState,
  captureModifyStructureState,
  captureSheetValueDataRange,
  captureValueDataRange,
  isRecoverySheetVisibility,
} from "./vendor/pi-recovery.js";

const RECOVERY_DB = "excel-assistant-recovery";
const RECOVERY_STORE = "settings";
const DOCUMENT_TOKEN_KEY = "excel-assistant-recovery-document-token-v1";

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

const CUSTOM_RECOVERY_PREFIX = "excel-assistant-custom-recovery-v1:";

async function customWorkbookId() {
  const workbook = await currentWorkbookContext();
  return workbook.workbookId ?? (await documentTokenIdentity().ensure());
}

async function readCustomSnapshots() {
  const workbookId = await customWorkbookId();
  if (!workbookId) return [];
  return (await recoverySettings.get(`${CUSTOM_RECOVERY_PREFIX}${workbookId}`)) ?? [];
}

async function writeCustomSnapshots(snapshots) {
  const workbookId = await customWorkbookId();
  if (!workbookId) return;
  await recoverySettings.set(`${CUSTOM_RECOVERY_PREFIX}${workbookId}`, snapshots.slice(0, 120));
}

async function appendCustomSnapshot({ toolName, toolCallId, address, state, restoredFromSnapshotId }) {
  const workbookId = await customWorkbookId();
  if (!workbookId) return null;
  const snapshot = {
    id: `checkpoint_custom_${Date.now().toString(36)}_${Math.random().toString(36).slice(2, 9)}`,
    at: Date.now(),
    toolName,
    toolCallId,
    address,
    changedCount: Number(state.count ?? state.values?.length ?? 1),
    workbookId,
    snapshotKind: "custom_state",
    customState: state,
    ...(restoredFromSnapshotId ? { restoredFromSnapshotId } : {}),
  };
  const snapshots = await readCustomSnapshots();
  snapshots.unshift(snapshot);
  await writeCustomSnapshots(snapshots);
  return snapshot;
}

async function deleteCustomSnapshot(id) {
  const snapshots = await readCustomSnapshots();
  const next = snapshots.filter((item) => item.id !== id);
  if (next.length === snapshots.length) return false;
  await writeCustomSnapshots(next);
  return true;
}

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

// The range a write will touch, resolved before anything is read from it.
// A copy expands a destination smaller than its source (Range.copyFrom), so
// with a source the extent is the larger of the two in each dimension — the
// same rule the copy itself uses. Only dimensions are loaded here.
async function captureRange(context, target) {
  const { worksheet, a1 } = await resolveWorksheet(context, target);
  let range = target.useUsedRange ? worksheet.getUsedRangeOrNullObject() : worksheet.getRange(a1);
  range.load("address,rowCount,columnCount,isNullObject");
  const source = target.sizeFromAddress ? worksheet.getRange(target.sizeFromAddress) : null;
  source?.load("rowCount,columnCount");
  await context.sync();
  let baseRows = 1;
  let baseColumns = 1;
  if (range.isNullObject) range = worksheet.getRange("A1");
  else {
    baseRows = range.rowCount;
    baseColumns = range.columnCount;
  }
  let rows = target.rows;
  let columns = target.columns;
  if (source) {
    rows = Math.max(source.rowCount, baseRows);
    columns = Math.max(source.columnCount, baseColumns);
  }
  if (rows && columns) range = range.getCell(0, 0).getResizedRange(rows - 1, columns - 1);
  range.load("address,rowCount,columnCount");
  await context.sync();
  return range;
}

async function captureRangeSnapshot(target) {
  return Excel.run(async (context) => {
    const range = await captureRange(context, target);
    // Refuse on size before loading a single value: a million-cell target used
    // to be read in full and only then turned away.
    const cellCount = range.rowCount * range.columnCount;
    if (cellCount > MAX_RECOVERY_CELLS) {
      throw new Error(`Recovery snapshot exceeds the ${MAX_RECOVERY_CELLS}-cell limit.`);
    }
    range.load("values,formulas");
    await context.sync();
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
  return Excel.run(async (context) => (await captureRange(context, target)).address);
}

// Pi records one format state per area and gives up when an area mixes formats
// (a header row with one bold cell, banded fills). A single cell is never mixed,
// so ranges small enough to list are captured cell by cell. The snapshot keeps
// that address, so the inverse checkpoint Pi takes when restoring does too.
const MAX_CELLWISE_FORMAT_CELLS = 500;

function cellwiseAddress(address) {
  const { sheetName, a1 } = splitSheetAddress(address);
  const match = /^\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/i.exec(a1 ?? "");
  if (!match) return null;
  const column = (label) => [...label.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
  const letters = (n) => (n > 0 ? letters(Math.floor((n - 1) / 26)) + String.fromCharCode(65 + ((n - 1) % 26)) : "");
  const [c1, r1, c2, r2] = [column(match[1]), Number(match[2]), column(match[3] ?? match[1]), Number(match[4] ?? match[2])];
  if ((c2 - c1 + 1) * (r2 - r1 + 1) > MAX_CELLWISE_FORMAT_CELLS) return null;
  const cells = [];
  for (let r = r1; r <= r2; r += 1) for (let c = c1; c <= c2; c += 1) cells.push(`${letters(c)}${r}`);
  return `${sheetName ? `'${sheetName.replaceAll("'", "''")}'!` : ""}${cells.join(",")}`;
}

async function captureFormatSnapshot(target, selection = CELL_FORMAT_PROPERTIES) {
  const rangeAddress = await qualifiedAddress(target);
  const address = cellwiseAddress(rangeAddress) ?? rangeAddress;
  const captured = await captureFormatCellsState(address, selection, {
    maxCellCount: MAX_RECOVERY_CELLS,
  });
  if (!captured.supported || !captured.state) {
    throw new Error(captured.reason ?? "Format recovery is unavailable for this range.");
  }
  return { kind: "format", address, state: captured.state, cellCount: captured.state.cellCount };
}

// A cell note is written as a threaded comment (Excel on Windows has no Notes
// API), so pi-for-excel's comment_thread snapshots restore it: capture the
// thread that is at each target cell before the write replaces it.
const MAX_COMMENT_CAPTURES = 20;

function notedCells(cells) {
  const noted = [];
  if (!Array.isArray(cells)) return noted;
  for (let row = 0; row < cells.length; row += 1) {
    const line = cells[row];
    if (!Array.isArray(line)) continue;
    for (let column = 0; column < line.length; column += 1) {
      if (line[column]?.note) noted.push({ row, column });
    }
  }
  return noted;
}

async function captureCommentSnapshot(target, offset) {
  const address = await Excel.run(async (context) => {
    const { worksheet, a1 } = await resolveWorksheet(context, target);
    const cell = worksheet.getRange(a1).getCell(offset.row, offset.column);
    cell.load("address");
    await context.sync();
    return cell.address;
  });
  return { kind: "comment", address, state: await captureCommentThreadState(address), cellCount: 1 };
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
              // The fill covers copyToRange, expanded to the pattern if smaller;
              // sizing it by the pattern's own shape captured only its first block.
              { capture: "range", target: bySheetId(args.copyToRange, { sizeFromAddress: args.range }) },
              {
                capture: "format",
                target: bySheetId(args.copyToRange, { sizeFromAddress: args.range }),
                selection: CELL_FORMAT_PROPERTIES,
              },
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
      const noted = notedCells(args.cells);
      if (noted.length > MAX_COMMENT_CAPTURES) {
        plans.push({
          unsupported: `Only ${MAX_COMMENT_CAPTURES} cell comments can be checkpointed in one write; this one sets ${noted.length}.`,
        });
      } else {
        for (const offset of noted) {
          plans.push({ capture: "comment", target: bySheetId(args.range), offset });
        }
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

async function captureAutoFilterState(target) {
  return Excel.run(async (context) => {
    const { worksheet } = await resolveWorksheet(context, target);
    worksheet.load("name");
    const filter = worksheet.autoFilter;
    filter.load("enabled,criteria");
    const range = filter.getRangeOrNullObject();
    range.load("isNullObject,address");
    await context.sync();
    return {
      address: `${worksheet.name}!autofilter`,
      state: {
        kind: "autofilter",
        sheetName: worksheet.name,
        range: filter.enabled && !range.isNullObject ? range.address : null,
        criteria: filter.enabled ? JSON.parse(JSON.stringify(filter.criteria ?? [])) : [],
        count: 1,
      },
    };
  });
}

async function captureDimensionState(args) {
  return Excel.run(async (context) => {
    const { worksheet } = await resolveWorksheet(context, { sheetId: args.sheetId });
    worksheet.load("name");
    const span = structureSpan(args);
    const hidden = [];
    for (let offset = 0; offset < span.count; offset += 1) {
      const position = span.position + offset;
      const address = span.kind === "rows" ? `${position}:${position}` : `${columnLetters(position)}:${columnLetters(position)}`;
      const range = worksheet.getRange(address);
      range.load(span.kind === "rows" ? "rowHidden" : "columnHidden");
      hidden.push(range);
    }
    await context.sync();
    return {
      address: `${worksheet.name}!${span.address}`,
      state: {
        kind: "dimension_visibility",
        sheetId: args.sheetId,
        dimension: args.dimension,
        position: span.position,
        hidden: hidden.map((range) =>
          span.kind === "rows" ? Boolean(range.rowHidden) : Boolean(range.columnHidden),
        ),
        count: span.count,
      },
    };
  });
}

async function captureFreezeState(args) {
  return Excel.run(async (context) => {
    const { worksheet } = await resolveWorksheet(context, { sheetId: args.sheetId });
    worksheet.load("name");
    const location = worksheet.freezePanes.getLocationOrNullObject();
    location.load("isNullObject,rowCount,columnCount");
    await context.sync();
    return {
      address: `${worksheet.name}!frozen panes`,
      state: {
        kind: "freeze_panes",
        sheetId: args.sheetId,
        rows: location.isNullObject ? 0 : location.rowCount,
        columns: location.isNullObject ? 0 : location.columnCount,
        count: 1,
      },
    };
  });
}

async function captureTableState(tableName) {
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(tableName);
    const range = table.getRange();
    table.load("name,showHeaders");
    range.load("address");
    await context.sync();
    return {
      kind: "table_present",
      name: table.name,
      address: range.address,
      hasHeaders: Boolean(table.showHeaders),
      count: 1,
    };
  });
}

async function captureTableRowCount(tableName) {
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(tableName);
    table.rows.load("count");
    await context.sync();
    return table.rows.count;
  });
}

async function prepareCustomRecovery(name, args, toolCallId) {
  let captured;
  let tableValues = null;
  let tableFormat = null;
  const limitations = [];
  if (name === "excel_modify_sheet_structure" && ["hide", "unhide"].includes(args.operation)) {
    captured = await captureDimensionState(args);
  } else if (
    name === "excel_modify_sheet_structure" &&
    ["freeze", "unfreeze"].includes(args.operation)
  ) {
    captured = await captureFreezeState(args);
  } else if (name === "excel_add_table_rows") {
    const beforeCount = await captureTableRowCount(args.table);
    captured = {
      address: args.table,
      state: {
        kind: "table_rows_added",
        table: args.table,
        index: Number.isInteger(args.index) ? args.index : beforeCount,
        count: args.values.length,
        expectedRowCount: beforeCount + args.values.length,
      },
    };
  } else if (name === "excel_create_table") {
    const target = { address: args.address, sheet: args.sheet };
    try {
      tableValues = await captureRangeSnapshot(target);
    } catch (error) {
      limitations.push(`Original cell values were not captured: ${error?.message ?? String(error)}`);
    }
    try {
      tableFormat = await captureFormatSnapshot(target);
    } catch (error) {
      limitations.push(`Original cell formatting was not captured: ${error?.message ?? String(error)}`);
    }
  } else if (name === "excel_autofilter") {
    captured = await captureAutoFilterState({ address: args.address, sheet: args.sheet });
  }

  return async function commitCustomRecovery(result) {
    let checkpoint = captured;
    if (name === "excel_create_table") {
      const tableName = result?.table;
      if (!tableName) return { status: "not_available", reason: "The created table name was unavailable." };
      checkpoint = {
        address: result?.range ?? args.address,
        state: { kind: "table_absent", table: tableName, count: 1 },
      };
    }
    if (!checkpoint) return { status: "not_available", reason: "The previous workbook state was unavailable." };
    const snapshot = await appendCustomSnapshot({
      toolName: name,
      toolCallId,
      address: checkpoint.address,
      state: checkpoint.state,
    });
    if (!snapshot) return { status: "not_available", reason: "The recovery log did not accept the checkpoint." };
    const snapshots = [snapshot];
    if (tableValues) {
      try {
        snapshots.push(await recoveryLog.append({
          toolName: "write_cells",
          toolCallId,
          address: tableValues.address,
          changedCount: tableValues.cellCount,
          beforeValues: tableValues.beforeValues,
          beforeFormulas: tableValues.beforeFormulas,
        }));
      } catch (error) {
        limitations.push(`Original cell values could not be saved: ${error?.message ?? String(error)}`);
      }
    }
    if (tableFormat) {
      try {
        snapshots.push(await recoveryLog.appendFormatCells({
          toolName: "format_cells",
          toolCallId,
          address: tableFormat.address,
          changedCount: tableFormat.cellCount,
          formatRangeState: tableFormat.state,
        }));
      } catch (error) {
        limitations.push(`Original cell formatting could not be saved: ${error?.message ?? String(error)}`);
      }
    }
    const saved = snapshots.filter(Boolean);
    return {
      status: "checkpoint_created",
      snapshotIds: saved.map((item) => item.id),
      targets: saved.map((item) => item.address),
      ...(limitations.length ? { partial: true, unavailableReasons: limitations } : {}),
    };
  };
}

async function restoreCustomState(state) {
  switch (state.kind) {
    case "dimension_visibility":
      return Excel.run(async (context) => {
        const { worksheet } = await resolveWorksheet(context, { sheetId: state.sheetId });
        const inverse = [];
        const ranges = [];
        for (let offset = 0; offset < state.hidden.length; offset += 1) {
          const position = state.position + offset;
          const address = state.dimension === "rows"
            ? `${position}:${position}`
            : `${columnLetters(position)}:${columnLetters(position)}`;
          const range = worksheet.getRange(address);
          range.load(state.dimension === "rows" ? "rowHidden" : "columnHidden");
          ranges.push(range);
        }
        await context.sync();
        for (let offset = 0; offset < ranges.length; offset += 1) {
          const range = ranges[offset];
          inverse.push(
            state.dimension === "rows" ? Boolean(range.rowHidden) : Boolean(range.columnHidden),
          );
          if (state.dimension === "rows") range.rowHidden = state.hidden[offset];
          else range.columnHidden = state.hidden[offset];
        }
        await context.sync();
        return { ...state, hidden: inverse };
      });
    case "freeze_panes":
      return Excel.run(async (context) => {
        const { worksheet } = await resolveWorksheet(context, { sheetId: state.sheetId });
        const location = worksheet.freezePanes.getLocationOrNullObject();
        location.load("isNullObject,rowCount,columnCount");
        await context.sync();
        const inverse = {
          ...state,
          rows: location.isNullObject ? 0 : location.rowCount,
          columns: location.isNullObject ? 0 : location.columnCount,
        };
        if (state.rows && state.columns) {
          worksheet.freezePanes.freezeAt(
            worksheet.getRange(`A1:${columnLetters(state.columns)}${state.rows}`),
          );
        } else if (state.rows) {
          worksheet.freezePanes.freezeRows(state.rows);
        } else if (state.columns) {
          worksheet.freezePanes.freezeColumns(state.columns);
        } else {
          worksheet.freezePanes.unfreeze();
        }
        await context.sync();
        return inverse;
      });
    case "autofilter": {
      const inverse = await captureAutoFilterState({ sheet: state.sheetName });
      await Excel.run(async (context) => {
        const worksheet = context.workbook.worksheets.getItem(state.sheetName);
        worksheet.autoFilter.remove();
        if (state.range) {
          worksheet.autoFilter.apply(worksheet.getRange(splitSheetAddress(state.range).a1));
          for (let index = 0; index < state.criteria.length; index += 1) {
            if (state.criteria[index]?.filterOn) {
              worksheet.autoFilter.apply(
                worksheet.getRange(splitSheetAddress(state.range).a1),
                index,
                state.criteria[index],
              );
            }
          }
        }
        await context.sync();
      });
      return inverse.state;
    }
    case "table_absent": {
      const inverse = await captureTableState(state.table);
      await Excel.run(async (context) => {
        context.workbook.tables.getItem(state.table).convertToRange();
        await context.sync();
      });
      return inverse;
    }
    case "table_present":
      await Excel.run(async (context) => {
        const { worksheet, a1 } = await resolveWorksheet(context, { address: state.address });
        const table = worksheet.tables.add(a1, state.hasHeaders);
        table.name = state.name;
        await context.sync();
      });
      return { kind: "table_absent", table: state.name, count: 1 };
    case "table_rows_added":
      return Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(state.table);
        const body = table.getDataBodyRange();
        body.load("rowCount,columnCount");
        await context.sync();
        if (body.rowCount !== state.expectedRowCount || state.index + state.count > body.rowCount) {
          throw new Error("The table row layout changed, so the appended rows cannot be removed safely.");
        }
        const removed = body
          .getCell(state.index, 0)
          .getResizedRange(state.count - 1, body.columnCount - 1);
        removed.load("values,formulas");
        await context.sync();
        const values = removed.values;
        const formulas = removed.formulas;
        for (let offset = state.count - 1; offset >= 0; offset -= 1) {
          table.rows.getItemAt(state.index + offset).delete();
        }
        await context.sync();
        return { kind: "table_rows_missing", table: state.table, index: state.index, values, formulas, count: state.count };
      });
    case "table_rows_missing":
      await Excel.run(async (context) => {
        const table = context.workbook.tables.getItem(state.table);
        table.rows.add(state.index, state.values);
        await context.sync();
        const body = table.getDataBodyRange();
        body.load("columnCount");
        await context.sync();
        const restored = body
          .getCell(state.index, 0)
          .getResizedRange(state.count - 1, body.columnCount - 1);
        restored.formulas = state.formulas;
        await context.sync();
      });
      return {
        kind: "table_rows_added",
        table: state.table,
        index: state.index,
        count: state.values.length,
        expectedRowCount: await captureTableRowCount(state.table),
      };
    default:
      throw new Error("This recovery checkpoint type is no longer supported.");
  }
}

async function restoreCustomSnapshot(snapshot) {
  const inverseState = await restoreCustomState(snapshot.customState);
  const inverse = await appendCustomSnapshot({
    toolName: "restore_snapshot",
    toolCallId: `restore_${snapshot.id}`,
    address: snapshot.address,
    state: inverseState,
    restoredFromSnapshotId: snapshot.id,
  });
  return {
    restoredSnapshotId: snapshot.id,
    inverseSnapshotId: inverse?.id,
    address: snapshot.address,
    changedCount: snapshot.changedCount ?? 1,
  };
}

// Charts use pi-for-excel's chart_state snapshots: an update stores the chart's
// previous properties, and a create stores a chart_absent state whose restore
// deletes the chart again by its captured id. A delete cannot be checkpointed —
// the present state can only be read while the chart still exists.
function chartAddress(sheetName, chartName) {
  const escaped = String(sheetName).replace(/'/gu, "''");
  const quoted = /[\s'!]/u.test(sheetName) ? `'${escaped}'` : sheetName;
  return `${quoted}!${chartName}`;
}

async function resolveChartIdentity(sheetId, id) {
  return Excel.run(async (context) => {
    const { worksheet } = await resolveWorksheet(context, { sheetId });
    worksheet.load("name");
    const charts = worksheet.charts;
    charts.load("items/id,items/name");
    await context.sync();
    const chart =
      charts.items.find((item) => item.id === id) ?? charts.items.find((item) => item.name === id);
    if (!chart) throw new Error(`Chart "${id}" was not found on ${worksheet.name}.`);
    return { sheetName: worksheet.name, name: chart.name, chartId: chart.id };
  });
}

async function prepareChartRecovery(args, toolCallId) {
  const limitations = [];
  let before = null;
  let failure = null;
  if (args.objectType !== "chart") {
    failure = "PivotTables are not included in automatic recovery yet; delete the PivotTable to undo it.";
  } else if (args.operation === "delete") {
    failure = "A deleted chart cannot be restored automatically; re-create it with excel_modify_object.";
  } else if (args.operation === "update") {
    try {
      const identity = await resolveChartIdentity(args.sheetId, args.id);
      before = {
        address: chartAddress(identity.sheetName, identity.name),
        state: await captureChartPresentState(identity.name, identity.sheetName),
      };
      if (args.properties?.source) {
        limitations.push("The chart's previous data source is not part of the checkpoint; only its type, title, legend, name and position are restored.");
      }
    } catch (error) {
      failure = `Chart backup capture failed: ${error?.message ?? String(error)}`;
    }
  }

  return async function commitChartRecovery(result) {
    try {
      if (failure) return { status: "not_available", reason: failure };
      let checkpoint = before;
      if (args.operation === "create") {
        const created = result?.id;
        if (!created) {
          return { status: "not_available", reason: "The created chart's ID was unavailable." };
        }
        const identity = await resolveChartIdentity(args.sheetId, created);
        checkpoint = {
          address: chartAddress(identity.sheetName, identity.name),
          state: {
            kind: "chart_absent",
            sheetName: identity.sheetName,
            name: identity.name,
            ...(identity.chartId ? { chartId: identity.chartId } : {}),
          },
        };
      }
      if (!checkpoint) {
        return { status: "not_available", reason: "The chart state could not be captured." };
      }
      const snapshot = await recoveryLog.appendChart({
        toolName: "charts",
        toolCallId,
        address: checkpoint.address,
        chartState: checkpoint.state,
      });
      return snapshot
        ? {
            status: "checkpoint_created",
            snapshotIds: [snapshot.id],
            targets: [snapshot.address],
            ...(limitations.length ? { partial: true, unavailableReasons: limitations } : {}),
          }
        : { status: "not_available", reason: "The recovery log did not accept the checkpoint." };
    } catch (error) {
      return { status: "not_available", reason: error?.message ?? String(error) };
    }
  };
}

// Structure edits map onto pi-for-excel's structure states, the same ones its
// modify_structure tool records. Removals capture the removed data first
// (restore re-inserts it); insertions only need their position (restore deletes
// them again), and a new sheet's id is known only after the edit.
function structureSpan(args) {
  const count = args.count ?? 1;
  const after = args.operation === "insert" && args.position === "after" ? 1 : 0;
  if (args.dimension === "rows") {
    const start = Number.parseInt(args.reference, 10) + after;
    return { kind: "rows", position: start, count, address: `${start}:${start + count - 1}` };
  }
  const letters = (n) => (n > 0 ? letters(Math.floor((n - 1) / 26)) + String.fromCharCode(65 + ((n - 1) % 26)) : "");
  const start = [...String(args.reference).toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0) + after;
  return { kind: "columns", position: start, count, address: `${letters(start)}:${letters(start + count - 1)}` };
}

async function prepareStructureCheckpoint(name, args) {
  const workbookEdit = name === "excel_modify_workbook_structure";
  if (!workbookEdit && !["insert", "delete"].includes(args.operation)) {
    throw new Error(`Row/column ${args.operation} is not included in automatic recovery yet.`);
  }
  if (workbookEdit && args.operation === "duplicate") {
    // Pi refuses to delete a sheet that holds data, and a copy always does.
    throw new Error("A duplicated sheet has no checkpoint; undo it by deleting the copy, which is checkpointed.");
  }
  if (workbookEdit && !["create", "delete", "rename"].includes(args.operation)) {
    throw new Error(`Sheet ${args.operation} is not included in automatic recovery yet.`);
  }
  if (workbookEdit && args.operation === "create") {
    return { after: (result) => ({ newSheetId: result?.sheetId }) };
  }
  return Excel.run(async (context) => {
    const { worksheet } = await resolveWorksheet(context, { sheetId: args.sheetId });
    worksheet.load("id,name,position,visibility");
    await context.sync();
    const sheet = { sheetId: worksheet.id, sheetName: worksheet.name };

    if (!workbookEdit) {
      const span = structureSpan(args);
      if (args.operation === "insert") {
        return { address: `${worksheet.name}!${span.address}`, state: { kind: `${span.kind}_absent`, ...sheet, position: span.position, count: span.count } };
      }
      const data = await captureValueDataRange(context, worksheet.getRange(span.address), MAX_RECOVERY_CELLS);
      if (data.status === "too_large") throw new Error("The deleted data exceeds the recovery size limit.");
      return {
        address: `${worksheet.name}!${span.address}`,
        state: { kind: `${span.kind}_present`, ...sheet, position: span.position, count: span.count, ...(data.dataRange ? { dataRange: data.dataRange } : {}) },
      };
    }
    if (args.operation === "delete") {
      if (!isRecoverySheetVisibility(worksheet.visibility)) throw new Error("This sheet's visibility can't be restored.");
      const data = await captureSheetValueDataRange(context, worksheet, MAX_RECOVERY_CELLS);
      if (data.status === "too_large") throw new Error("The deleted sheet exceeds the recovery size limit.");
      return {
        address: worksheet.name,
        state: { kind: "sheet_present", ...sheet, position: worksheet.position, visibility: worksheet.visibility, ...(data.dataRange ? { dataRange: data.dataRange } : {}) },
      };
    }
    return { renameOf: worksheet.name };
  });
}

async function newSheetAbsentState(newSheetId) {
  return Excel.run(async (context) => {
    const { worksheet } = await resolveWorksheet(context, { sheetId: newSheetId });
    worksheet.load("id,name");
    await context.sync();
    return { address: worksheet.name, state: { kind: "sheet_absent", sheetId: worksheet.id, sheetName: worksheet.name } };
  });
}

async function prepareStructureRecovery(name, args, toolCallId) {
  let prepared = null;
  let failure = null;
  const limitations = [];
  // Deleting rows or columns turns same-sheet formulas that referenced them
  // into #REF!, and re-inserting the data does not repair them. Keep the
  // sheet's values and formulas too; restore puts them back after re-inserting.
  let sheetValues = null;
  try {
    prepared = await prepareStructureCheckpoint(name, args);
    if (prepared.renameOf) {
      const state = await captureModifyStructureState({ kind: "sheet_name", sheetRef: prepared.renameOf });
      prepared = state ? { address: prepared.renameOf, state } : null;
    }
    if (name === "excel_modify_sheet_structure" && args.operation === "delete") {
      try {
        sheetValues = await captureRangeSnapshot({ sheetId: args.sheetId, useUsedRange: true });
      } catch (error) {
        limitations.push(`The sheet-wide formula snapshot failed: ${error?.message ?? String(error)}`);
      }
      limitations.push("Formulas on other worksheets that referenced the deleted rows or columns are not restored.");
    }
    if (name === "excel_modify_workbook_structure" && args.operation === "delete") {
      limitations.push("Formulas on other worksheets that referenced the deleted sheet are not restored.");
    }
  } catch (error) {
    failure = error?.message ?? String(error);
  }
  return async function commitStructureRecovery(result, failure) {
    // A structure checkpoint is an inverse: rows_absent deletes the rows this
    // call inserted, sheet_absent deletes the sheet it created. Published for a
    // call that failed, it can delete rows or a sheet that were already there.
    // The arguments cannot show whether a failed change happened, so none is
    // recorded; the user is told to look instead.
    if (failure) {
      return {
        status: "not_available",
        reason:
          failure.commitStatus === "not_committed"
            ? "The structure change did not happen, so there is nothing to undo."
            : "The structure change failed part-way. No undo was recorded, because whether it happened cannot be told from the request; check the rows, columns or sheets involved.",
      };
    }
    try {
      const createdSheetId = prepared?.after ? prepared.after(result).newSheetId : null;
      const checkpoint = prepared?.after
        ? createdSheetId === undefined || createdSheetId === null
          ? null
          : await newSheetAbsentState(createdSheetId)
        : prepared;
      if (!checkpoint?.state) {
        return {
          status: "not_available",
          reason: failure ?? (prepared?.after
            ? "The new sheet ID was unavailable after the operation failed."
            : "The sheet structure state could not be captured."),
        };
      }
      const valuesSnapshot = sheetValues
        ? await recoveryLog.append({
            toolName: "modify_structure",
            toolCallId,
            address: sheetValues.address,
            changedCount: sheetValues.cellCount,
            beforeValues: sheetValues.beforeValues,
            beforeFormulas: sheetValues.beforeFormulas,
          })
        : null;
      const snapshot = await recoveryLog.appendModifyStructure({
        toolName: "modify_structure",
        toolCallId,
        address: checkpoint.address,
        modifyStructureState: checkpoint.state,
      });
      const created = [snapshot, valuesSnapshot].filter(Boolean);
      return created.length
        ? {
            status: "checkpoint_created",
            snapshotIds: created.map((item) => item.id),
            targets: created.map((item) => item.address),
            ...(limitations.length ? { partial: true, unavailableReasons: limitations } : {}),
          }
        : { status: "not_available", reason: "The recovery log did not accept the checkpoint." };
    } catch (error) {
      return { status: "not_available", reason: error?.message ?? String(error) };
    }
  };
}

// Before/after cell diffs for the tool card. The "before" side is the recovery
// capture that already exists, so a diff only costs one extra read. It stays in
// the pane: the model gets the receipt, the user gets the detail.
const MAX_DIFF_CELLS = 500;
const MAX_DIFF_ENTRIES = 20;
const mutationDiffs = new Map();

export function takeMutationDiff(toolCallId) {
  const diff = mutationDiffs.get(toolCallId);
  mutationDiffs.delete(toolCallId);
  return diff;
}

function columnLetters(number) {
  return number > 0
    ? columnLetters(Math.floor((number - 1) / 26)) + String.fromCharCode(65 + ((number - 1) % 26))
    : "";
}

function diffCellText(value, formula) {
  if (typeof formula === "string" && formula.startsWith("=")) return formula;
  if (value === null || value === undefined || value === "") return "(blank)";
  return String(value).slice(0, 40);
}

async function recordMutationDiff(toolCallId, captures) {
  const changes = [];
  let changed = 0;
  for (const capture of captures) {
    if (capture.kind !== "range" || capture.cellCount > MAX_DIFF_CELLS) continue;
    const after = await captureRangeSnapshot({ address: capture.address }).catch(() => null);
    if (!after) continue;
    const { sheetName, a1 } = splitSheetAddress(capture.address);
    const start = /^\$?([A-Z]+)\$?(\d+)/i.exec(a1 ?? "");
    if (!start) continue;
    const firstColumn = [...start[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
    const firstRow = Number(start[2]);
    for (let r = 0; r < capture.beforeValues.length; r += 1) {
      for (let c = 0; c < (capture.beforeValues[r]?.length ?? 0); c += 1) {
        const before = diffCellText(capture.beforeValues[r][c], capture.beforeFormulas?.[r]?.[c]);
        const now = diffCellText(after.beforeValues?.[r]?.[c], after.beforeFormulas?.[r]?.[c]);
        if (before === now) continue;
        changed += 1;
        if (changes.length >= MAX_DIFF_ENTRIES) continue;
        const cell = `${columnLetters(firstColumn + c)}${firstRow + r}`;
        changes.push({ cell: sheetName ? `${sheetName}!${cell}` : cell, before, after: now });
      }
    }
  }
  if (changed > 0) mutationDiffs.set(toolCallId, { changed, changes });
}

export async function prepareMutationRecovery(name, args, toolCallId) {
  if (
    (name === "excel_modify_sheet_structure" &&
      ["hide", "unhide", "freeze", "unfreeze"].includes(args?.operation)) ||
    name === "excel_create_table" ||
    name === "excel_add_table_rows" ||
    name === "excel_autofilter"
  ) {
    try {
      return await prepareCustomRecovery(name, args ?? {}, toolCallId);
    } catch (error) {
      return async () => ({ status: "not_available", reason: error?.message ?? String(error) });
    }
  }
  if (name === "excel_modify_sheet_structure" || name === "excel_modify_workbook_structure") {
    return prepareStructureRecovery(name, args ?? {}, toolCallId);
  }
  if (name === "excel_modify_object") {
    return prepareChartRecovery(args ?? {}, toolCallId);
  }
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
          : item.capture === "comment"
            ? await captureCommentSnapshot(item.target, item.offset)
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
        const snapshot = capture.kind === "comment"
          ? await recoveryLog.appendCommentThread({
              toolName: "comments",
              toolCallId,
              address: capture.address,
              changedCount: 1,
              commentThreadState: capture.state,
            })
          : capture.kind === "format"
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

    await recordMutationDiff(toolCallId, captures).catch(() => {});

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
  const snapshots = await allSnapshots();
  const anchor = snapshotId
    ? snapshots.find((snapshot) => snapshot.id === snapshotId)
    : snapshots[0];
  if (!anchor) throw new Error("No recovery checkpoint is available for this workbook.");
  return snapshots.filter((snapshot) => snapshot.toolCallId === anchor.toolCallId);
}

async function allSnapshots() {
  const snapshots = [
    ...(await recoveryLog.listForCurrentWorkbook(120)),
    ...(await readCustomSnapshots()),
  ];
  return snapshots.sort((a, b) => Number(b.at ?? 0) - Number(a.at ?? 0)).slice(0, 120);
}

export async function workbookHistory({ action = "list", snapshot_id: snapshotId, limit = 20 } = {}) {
  switch (action) {
    case "list":
      return {
        success: true,
        action,
        snapshots: groupSnapshots(await allSnapshots())
          .slice(0, limit)
          .map(compactSnapshotGroup),
      };
    case "restore": {
      const group = await resolveSnapshotGroup(snapshotId);
      const restored = [];
      // Re-insert rows/columns/sheets before writing values back into them.
      const ordered = [...group].sort(
        (a, b) =>
          Number(["modify_structure_state", "custom_state"].includes(b.snapshotKind)) -
          Number(["modify_structure_state", "custom_state"].includes(a.snapshotKind)),
      );
      for (const snapshot of ordered) {
        restored.push(
          snapshot.snapshotKind === "custom_state"
            ? await restoreCustomSnapshot(snapshot)
            : await recoveryLog.restore(snapshot.id),
        );
      }
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
      const deleted = await Promise.all(
        group.map((snapshot) =>
          snapshot.snapshotKind === "custom_state"
            ? deleteCustomSnapshot(snapshot.id)
            : recoveryLog.delete(snapshot.id),
        ),
      );
      return { success: deleted.every(Boolean), action, snapshotIds: group.map((snapshot) => snapshot.id) };
    }
    case "clear": {
      const piRemoved = await recoveryLog.clearForCurrentWorkbook();
      const customRemoved = (await readCustomSnapshots()).length;
      await writeCustomSnapshots([]);
      return { success: true, action, removed: piRemoved + customRemoved };
    }
    default:
      throw new Error(`Unsupported workbook history action: ${action}`);
  }
}
