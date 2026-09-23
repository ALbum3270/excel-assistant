// Exercises the production recovery dispatcher and gateway with isolated I/O.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import { webcrypto } from "node:crypto";
import { createThepExcelGateway } from "../daemon/thepexcel-gateway.mjs";
import { createWorkbookExecution } from "../daemon/workbook-execution.mjs";

const read = (path) =>
  readFileSync(new URL(`../${path}`, import.meta.url), "utf8").replace(/\r\n/g, "\n");
function cut(text, a, b) {
  const start = text.indexOf(a),
    end = text.indexOf(b, start + a.length);
  assert.ok(start >= 0 && end > start, `Missing anchors: ${a} / ${b}`);
  return text.slice(start, end);
}
function run(code, bindings = {}) {
  const context = vm.createContext(bindings);
  vm.runInContext(code, context);
  return context;
}
const recovery = read("taskpane/shared/recovery.js");
const pi = read("taskpane/shared/vendor/pi-recovery.js");

test("native Windows paths keep a # in the workbook recovery identity", async () => {
  const functions = run(cut(recovery, "function workbookName(", "const recoveryLog ="), {
    Office: { context: { document: { url: "" } } },
    TextEncoder,
    crypto: webcrypto,
  });
  const ids = [];
  for (const url of ["C:\\Work\\Budget#A.xlsx", "C:\\Work\\Budget#B.xlsx"]) {
    functions.Office.context.document.url = url;
    ids.push((await functions.currentWorkbookContext()).workbookId);
  }
  assert.notEqual(ids[0], ids[1]);
});

test("a cloud workbook keeps the identity it already had in the log", async () => {
  // The fix for '#' in native paths must not renumber documents that were
  // already recorded: a URL is cut at '?' or '#', not re-serialized, which
  // would percent-encode spaces and lowercase the host.
  const functions = run(cut(recovery, "function workbookName(", "const recoveryLog ="), {
    Office: { context: { document: { url: "" } } },
    TextEncoder,
    crypto: webcrypto,
  });
  const idFor = async (url) => {
    functions.Office.context.document.url = url;
    return (await functions.currentWorkbookContext()).workbookId;
  };
  const legacy = async (url) => {
    functions.Office.context.document.url = url.split(/[?#]/, 1)[0];
    return (await functions.currentWorkbookContext()).workbookId;
  };
  const shared = "https://Contoso.SharePoint.com/sites/fin/Shared Documents/Book.xlsx";
  assert.equal(await idFor(shared), await legacy(shared), "space or host case changed the id");
  assert.equal(await idFor(`${shared}?web=1`), await idFor(shared), "query string changed the id");
  assert.equal(await idFor(`${shared}#anchor`), await idFor(shared), "fragment changed the id");
});

test("restore resolves every snapshot in an operation before applying the history display limit", async () => {
  const piSnapshots = [
    {
      id: "structure",
      toolCallId: "operation",
      at: 2,
      snapshotKind: "modify_structure_state",
      address: "Sheet1!A1",
    },
    {
      id: "values",
      toolCallId: "operation",
      at: 1,
      snapshotKind: "range_values",
      address: "Sheet1!A1",
    },
  ];
  const customSnapshots = Array.from({ length: 119 }, (_, index) => ({
    id: `recent-${index}`,
    toolCallId: `call-${index}`,
    at: index + 3,
    snapshotKind: "custom_state",
    address: "Sheet1!A1",
  }));
  const restored = [];
  const functions = run(
    recovery
      .slice(recovery.indexOf("function compactSnapshotGroup("))
      .replace("export async function workbookHistory", "async function workbookHistory"),
    {
      readCustomSnapshots: async () => customSnapshots,
      recoveryLog: {
        listForCurrentWorkbook: async () => piSnapshots,
        restore: async (id) => {
          restored.push(id);
          return { restoredSnapshotId: id, address: "Sheet1!A1", changedCount: 1 };
        },
      },
    },
  );
  const result = await functions.workbookHistory({ action: "restore", snapshot_id: "structure" });
  assert.equal(result.commitStatus, "committed");
  assert.deepEqual(new Set(restored), new Set(["structure", "values"]));
});

// Actual recovery plan and actual captureRange. A1 is the supported top-left
// shorthand for a 2x2 matrix; before the write, A1 itself is still a 1x1 range.
test("copy recovery uses the matrix size when the source is a single starting cell", async () => {
  const ranges = {
    A1: { address: "Sheet1!A1", rowCount: 1, columnCount: 1 },
    D1: { address: "Sheet1!D1", rowCount: 1, columnCount: 1 },
  };
  const range = (state) => ({
    ...state,
    isNullObject: false,
    load() {},
    getCell() {
      return {
        getResizedRange: (r, c) => range({ ...state, rowCount: r + 1, columnCount: c + 1 }),
      };
    },
  });
  const functions = run(
    cut(recovery, "function recoveryPlan(", "async function captureAutoFilterState(") +
      cut(recovery, "async function captureRange(", "async function captureRangeSnapshot("),
    {
      matrixShape: (cells) => ({ rows: cells.length, columns: cells[0].length }),
      notedCells: () => [],
      MAX_COMMENT_CAPTURES: 20,
      CELL_FORMAT_PROPERTIES: {},
      resolveWorksheet: async (_ctx, target) => ({
        a1: target.address,
        worksheet: { getRange: (a) => range(ranges[a]) },
      }),
    },
  );
  const plan = functions.recoveryPlan("excel_set_cell_range", {
    sheetId: 1,
    range: "A1",
    cells: [
      [{ value: 1 }, { value: 2 }],
      [{ value: 3 }, { value: 4 }],
    ],
    copyToRange: "D1",
    allow_overwrite: true,
  });
  for (const item of plan.slice(1)) {
    const captured = await functions.captureRange({ sync: async () => {} }, item.target);
    assert.equal(captured.rowCount, 2, `${item.capture} recovery must cover both rows`);
    assert.equal(captured.columnCount, 2, `${item.capture} recovery must cover both columns`);
  }
});

// Actual grouped history dispatcher and actual Pi restore function. Only the
// Excel cell storage/apply boundary is in memory. Start after deleting row 2.
test("grouped row recovery reverses its execution order on every undo and redo", async () => {
  let cells = ["A", "C", ""];
  let serial = 0;
  const scope = { workbookId: "book" };
  const snapshots = [
    {
      id: "structure",
      toolCallId: "delete-row",
      workbookId: "book",
      snapshotKind: "modify_structure_state",
      address: "Sheet1!2:2",
      changedCount: 1,
      modifyStructureState: { kind: "rows_present", position: 2, count: 1, data: "B" },
    },
    {
      id: "values",
      toolCallId: "delete-row",
      workbookId: "book",
      snapshotKind: "range_values",
      address: "Sheet1!A1:A3",
      changedCount: 3,
      beforeValues: [["A"], ["B"], ["C"]],
      beforeFormulas: [["A"], ["B"], ["C"]],
    },
  ];
  const restore = run(
    cut(pi, "function resolveSnapshotKind(", "// src/workbook/recovery/log-store.ts"),
  );
  const append = (snapshotKind) => async (args) => {
    const item = {
      ...args,
      id: `inverse-${++serial}`,
      snapshotKind,
      workbookId: "book",
      at: serial,
    };
    snapshots.unshift(item);
    return item;
  };
  const dependencies = {
    toRestoreValues: (values) => values,
    countChangedCells: () => 3,
    applySnapshot: async (_address, values) => {
      const old = Array.from({ length: values.length }, (_, i) => [cells[i] ?? ""]);
      values.forEach((row, i) => {
        cells[i] = row[0];
      });
      return { values: old, formulas: old };
    },
    applyModifyStructureSnapshot: async (_address, state) => {
      const i = state.position - 1;
      if (state.kind === "rows_present") {
        cells.splice(i, 0, state.data ?? "");
        return { ...state, kind: "rows_absent" };
      }
      const [data] = cells.splice(i, 1);
      return { ...state, kind: "rows_present", data };
    },
    appendRangeSnapshot: append("range_values"),
    appendModifyStructureSnapshot: append("modify_structure_state"),
  };
  const history = run(
    recovery
      .slice(recovery.indexOf("export async function workbookHistory("))
      .replace("export async", "async"),
    {
      resolveSnapshotGroup: async (id) => {
        const anchor = snapshots.find((s) => s.id === id);
        return snapshots.filter((s) => s.toolCallId === anchor.toolCallId);
      },
      recoveryLog: {
        restore: (id, options) =>
          restore.restoreWorkbookRecoverySnapshot({
            snapshot: snapshots.find((s) => s.id === id),
            scope,
            dependencies,
            ...options,
          }),
      },
      Date: { now: () => serial + 1000 },
    },
  );
  const clean = () => {
    const copy = [...cells];
    while (copy.at(-1) === "") copy.pop();
    return copy;
  };
  let id = "structure";
  for (let i = 0; i < 6; i++) {
    const restored = await history.workbookHistory({ action: "restore", snapshot_id: id });
    assert.deepEqual(clean(), i % 2 === 0 ? ["A", "B", "C"] : ["A", "C"], `restore ${i + 1}`);
    id = restored.inverseSnapshotIds[0];
    // Old snapshots may have been deleted/pruned; ordering must survive that.
    for (let j = snapshots.length - 1; j >= 0; j--) {
      if (!restored.inverseSnapshotIds.includes(snapshots[j].id)) snapshots.splice(j, 1);
    }
  }
});

// Real gateway/coordinator; this client honours signal cancellation, as the MCP
// client does. The old regression test's info request ignored that signal.
test("cancelling a signal-aware COM info request leaves writes unlocked at the same revision", async () => {
  const execution = createWorkbookExecution();
  let notifyStarted;
  const started = new Promise((resolve) => {
    notifyStarted = resolve;
  });
  const client = {
    listTools: async () => ({
      tools: [
        {
          name: "excel_range",
          inputSchema: {
            type: "object",
            properties: { action: { type: "string" }, workbook: { type: "string" } },
          },
        },
      ],
    }),
    getInstructions: () => "",
    callTool: (_request, _schema, { signal }) =>
      new Promise((_resolve, reject) => {
        signal.addEventListener("abort", () => reject(signal.reason), { once: true });
        notifyStarted();
      }),
  };
  const gateway = await createThepExcelGateway(null, execution, { client });
  const abort = new AbortController();
  const path = "C:/audit/book.xlsx";
  const server = gateway.createSessionServer({ workbookPath: path, signal: abort.signal });
  const pending = server.instance._registeredTools.excel_range.handler({ action: "write" });
  await started;
  abort.abort();
  await pending;
  await new Promise(setImmediate);
  const state = execution.snapshot(path);
  assert.equal(state.revision, 0);
  assert.equal(state.blocked, null);
  const next = await execution.run(path, { write: true, expectedRevision: 0 }, async () => ({
    success: true,
  }));
  assert.equal(next.revision, 1);
});
