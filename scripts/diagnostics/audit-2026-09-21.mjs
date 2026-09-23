// Audit reproductions, not business fixes or production tests.
// Run: node scripts/diagnostics/audit-2026-09-21.mjs
// Uses real functions; browser/COM/process I/O is replaced in isolated VMs.
// A reproduced finding is the CURRENT bug, not a passing acceptance test.
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { mkdtemp, mkdir, writeFile, readFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join, win32 } from "node:path";
import { EventEmitter } from "node:events";
import vm from "node:vm";
import { z } from "zod";
import { createWorkbookExecution, canonicalWorkbookId } from "../../daemon/workbook-execution.mjs";
import { backupWorkbookFile } from "../../daemon/com-backup.mjs";
import { createComputeShell } from "../../daemon/compute-tool.mjs";
import { needsApproval } from "../../daemon/approval.mjs";
import { createOfficeBridgeMcp } from "../../daemon/office-tools.mjs";

const source = (path) =>
  readFileSync(new URL(`../../${path}`, import.meta.url), "utf8").replace(/\r\n/g, "\n");
function between(text, start, end) {
  const a = text.indexOf(start),
    b = text.indexOf(end, a + start.length);
  assert.ok(a >= 0 && b > a, `Missing source anchors: ${start} / ${end}`);
  return text.slice(a, b);
}
function evaluate(code, bindings = {}) {
  const sandbox = vm.createContext(bindings);
  vm.runInContext(code, sandbox);
  return sandbox;
}
const indexSource = source("daemon/index.mjs");
const paneSource = source("taskpane/shared/taskpane.js");
const recoverySource = source("taskpane/shared/recovery.js");
const piSource = source("taskpane/shared/vendor/pi-recovery.js");
const officeSource = source("taskpane/shared/vendor/office-agents-excel-api.js");
const results = [];
async function check(id, body) {
  try {
    const evidence = await body();
    results.push({ id, reproduced: true, evidence });
  } catch (error) {
    results.push({ id, reproduced: false, error: error.stack });
    process.exitCode = 1;
  }
}

await check("B01-partial-write-revision", async () => {
  const execution = createWorkbookExecution();
  let cell = "before";
  await assert.rejects(
    execution.run("book", { write: true }, async () => {
      cell = "partially changed";
      throw Object.assign(new Error("later sync failed"), {
        commitStatus: "unknown",
        executionSettled: true,
      });
    }),
  );
  const after = execution.snapshot("book");
  let staleWriteRan = false;
  await execution.run("book", { write: true, expectedRevision: 0 }, async () => {
    staleWriteRan = true;
    return { success: true };
  });
  assert.equal(after.revision, 0);
  assert.equal(after.blocked, null);
  assert.equal(staleWriteRan, true);
  return { cell, revisionAfterPartialFailure: after.revision, staleWriteRan };
});

await check("B02-backup-collision", async () => {
  const root = await mkdtemp(join(tmpdir(), "excel-audit-"));
  const oldDir = process.env.EXCEL_COM_BACKUP_DIR,
    oldMode = process.env.EXCEL_COM_BACKUP;
  try {
    process.env.EXCEL_COM_BACKUP_DIR = join(root, "backups");
    process.env.EXCEL_COM_BACKUP = "on";
    await mkdir(join(root, "a"));
    await mkdir(join(root, "b"));
    const a = join(root, "a", "Budget.xlsx"),
      b = join(root, "b", "Budget.xlsx");
    await writeFile(a, "workbook-A-original");
    await writeFile(b, "workbook-B-original");
    const now = () => new Date("2026-09-21T12:00:00.123Z");
    const first = await backupWorkbookFile(a, { now });
    const second = await backupWorkbookFile(b, { now });
    const reused = await backupWorkbookFile(a, { now });
    const content = await readFile(first.file, "utf8");
    assert.equal(first.file, second.file);
    assert.equal(content, "workbook-B-original");
    assert.equal(reused.status, "reused");
    await writeFile(a, "workbook-A-new-state-longer");
    const sameSecond = await backupWorkbookFile(a, { now });
    assert.equal(first.file, sameSecond.file);
    return {
      differentWorkbooksShareFile: true,
      firstBackupContents: content,
      reuseReturnsWrongWorkbook: true,
      sameSecondSaveReplacesPreviousBackup: true,
    };
  } finally {
    if (oldDir === undefined) delete process.env.EXCEL_COM_BACKUP_DIR;
    else process.env.EXCEL_COM_BACKUP_DIR = oldDir;
    if (oldMode === undefined) delete process.env.EXCEL_COM_BACKUP;
    else process.env.EXCEL_COM_BACKUP = oldMode;
    // Only this call's mkdtemp directory is removed.
    await rm(root, { recursive: true, force: true });
  }
});

await check("B03-csv-type-loss", async () => {
  let cells;
  const shell = createComputeShell(async (name, args) => {
    if (name === "excel_get_range_as_csv")
      return {
        csv: "00123,9007199254740993,TRUE,=1+1",
        rowCount: 1,
        columnCount: 4,
        sheetName: "Sheet1",
      };
    cells = args.cells;
    return { success: true, commitStatus: "committed", writtenRange: args.range };
  });
  const result = await shell({
    command: "sheet-to-csv 1 A1:D1 data.csv && csv-to-sheet data.csv 1 A2 --force",
  });
  assert.equal(result.exitCode, 0);
  assert.deepEqual(cells[0], [
    { value: 123 },
    { value: 9007199254740992 },
    { value: true },
    { formula: "=1+1" },
  ]);
  return { sourceStrings: ["00123", "9007199254740993", "TRUE", "=1+1"], writtenCells: cells[0] };
});

await check("B04-script-approval-bypass", async () => {
  const calls = [];
  const shell = createComputeShell(async (name, args) => {
    calls.push(name);
    return { success: true, commitStatus: "committed", writtenRange: args.range };
  });
  await shell({
    command: "printf '42\\n' > data.csv; printf 'csv-to-sheet data.csv 1 A1 --force\\n' > later.sh",
  });
  const input = { command: "bash later.sh" };
  const approval = needsApproval("mcp__office__excel_bash", input);
  const result = await shell(input);
  assert.equal(approval, false);
  assert.equal(result.exitCode, 0);
  assert.deepEqual(calls, ["excel_set_cell_range"]);
  return { command: input.command, approval, writes: calls };
});

await check("B05-export-tool-result", async () => {
  const transcript = source("daemon/transcript.mjs");
  const parser = evaluate(between(transcript, "const MAX_RESULT_CHARS", "/**\n * Reconstruct"), {
    stripContextHeader: (text) => text,
  });
  const archive = evaluate(
    between(
      indexSource,
      "function portableTranscriptEvents(",
      "async function exportConversation(",
    ),
  );
  const events = parser.eventsFromLine({
    type: "user",
    message: {
      role: "user",
      content: [{ type: "tool_result", tool_use_id: "tool-1", content: '{"success":true}' }],
    },
  });
  assert.equal(events[0].kind, "tool_result");
  let error;
  try {
    archive.portableTranscriptEvents(events);
  } catch (e) {
    error = e.message;
  }
  assert.equal(error, "Conversation archive contains an unsupported event");
  return { inputKind: events[0].kind, error };
});

await check("B06-export-byte-budget", async () => {
  const events = Array.from({ length: 19 }, () => ({
    kind: "assistant",
    text: "中".repeat(100_000),
  }));
  const archive = evaluate(
    between(
      indexSource,
      "function portableTranscriptEvents(",
      "async function importConversation(",
    ),
    {
      getSessionRecord: async () => ({ archive_events: events, title: "中文记录" }),
      documentKeyForPane: () => "book",
    },
  );
  const exported = await archive.exportConversation("pane", "excel", "session");
  const serialized = JSON.stringify(exported, null, 2);
  const bytes = Buffer.byteLength(serialized, "utf8");
  assert.ok(bytes > 2_000_000);
  assert.ok(serialized.length < 2_000_000);
  return {
    exportedEvents: exported.events.length,
    characters: serialized.length,
    bytes,
    uiImportByteLimit: 2_000_000,
  };
});

await check("B07-recovery-persistence-breaks-result", async () => {
  const messages = [];
  let writes = 0,
    commits = 0;
  const custom = evaluate(
    between(
      recoverySource,
      "async function prepareCustomRecovery(",
      "async function restoreCustomState(",
    ),
    {
      captureAutoFilterState: async () => ({
        address: "Sheet1!autofilter",
        state: { kind: "autofilter", enabled: false },
      }),
      appendCustomSnapshot: async () => {
        commits++;
        throw new Error("QuotaExceededError");
      },
    },
  );
  const pane = evaluate(
    between(
      paneSource,
      "async function runOfficeTool(",
      "// ---------------------------------------------------------------------------\n// Selection tracking",
    ),
    {
      cancelledToolCalls: new Set(),
      WRITE_TOOLS: new Set(["excel_autofilter"]),
      isMutationCall: () => true,
      refuseInCellEditMode: async () => {},
      prepareMutationRecovery: custom.prepareCustomRecovery,
      toolExcelAutoFilter: async () => {
        writes++;
        return { success: true };
      },
      runWorkbookWrite: async (_id, _name, fn) => ({ result: await fn(), revision: 1 }),
      CANCELLED_TOOL_RESULT: Symbol(),
      wsSend: (msg) => messages.push(msg),
    },
  );
  await assert.rejects(
    pane.runOfficeTool({ id: "t", name: "excel_autofilter", args: {} }),
    /QuotaExceededError/,
  );
  assert.equal(writes, 1);
  assert.equal(commits, 2);
  assert.equal(messages.length, 0);
  return {
    workbookWrites: writes,
    attemptedCheckpointCommits: commits,
    resultMessages: messages.length,
  };
});

await check("B08-repeat-budget-survives-state-change", async () => {
  let blocked = true,
    calls = 0;
  const server = createOfficeBridgeMcp(
    {
      callTaskpaneTool: async (name) => {
        if (name === "excel_get_cell_ranges") return { success: true, workbookRevision: 1 };
        calls++;
        if (blocked) throw new Error("Excel is in cell-edit mode");
        return { success: true, commitStatus: "committed", workbookRevision: 2 };
      },
    },
    "excel",
    "audit-pane",
  );
  const tools = server.instance._registeredTools;
  const args = { sheetId: 1, range: "A1", clearType: "contents" };
  for (let i = 0; i < 3; i++) await tools.excel_clear_cell_range.handler(args);
  blocked = false; // User pressed Enter; another turn/read now sees new workbook state.
  await tools.excel_get_cell_ranges.handler({ sheetId: 1, ranges: ["A1"] });
  const result = await tools.excel_clear_cell_range.handler(args);
  assert.equal(calls, 3);
  assert.equal(result.isError, true);
  return {
    dispatchCount: calls,
    stateNoLongerBlocked: !blocked,
    fourthCall: result.content[0].text,
  };
});

await check("B09-recovery-limit-after-data-load", async () => {
  const loads = [];
  const range = {
    address: "Sheet1!A1:A1000000",
    rowCount: 1_000_000,
    columnCount: 1,
    isNullObject: false,
    load: (properties) => loads.push(properties),
  };
  const recovery = evaluate(
    between(
      recoverySource,
      "async function captureRangeSnapshot(",
      "async function qualifiedAddress(",
    ),
    {
      Excel: { run: (fn) => fn({ sync: async () => {} }) },
      resolveWorksheet: async () => ({ worksheet: { getRange: () => range }, a1: "A1:A1000000" }),
      MAX_RECOVERY_CELLS: 20_000,
    },
  );
  await assert.rejects(recovery.captureRangeSnapshot({ address: "A1:A1000000" }), /exceeds/);
  assert.ok(loads.some((properties) => properties.includes("values,formulas")));
  return { cells: range.rowCount, limit: 20_000, loadsBeforeRefusal: loads };
});

await check("B10-com-write-starts-after-stop", async () => {
  let releaseInfo, sawInfo;
  const infoStarted = new Promise((resolve) => {
    sawInfo = resolve;
  });
  const infoWait = new Promise((resolve) => {
    releaseInfo = resolve;
  });
  const calls = [];
  const execution = createWorkbookExecution();
  const abort = new AbortController();
  const gateway = evaluate(
    source("daemon/thepexcel-gateway.mjs")
      .slice(source("daemon/thepexcel-gateway.mjs").indexOf("const UPSTREAM_TIMEOUT_MS"))
      .replace("export async function", "async function"),
    {
      win32,
      z,
      canonicalWorkbookId,
      needsApproval,
      getDefaultEnvironment: () => ({}),
      StdioClientTransport: class {
        async close() {}
      },
      Client: class {
        async connect() {}
        async listTools() {
          return {
            tools: [
              {
                name: "excel_range",
                inputSchema: {
                  type: "object",
                  properties: { action: { type: "string" }, workbook: { type: "string" } },
                },
              },
            ],
          };
        }
        getInstructions() {
          return "";
        }
        async callTool(request) {
          calls.push(request.name);
          if (request.name === "excel_workbook") {
            sawInfo();
            await infoWait;
          }
          return {
            content: [{ type: "text", text: JSON.stringify({ path: "C:\\audit\\book.xlsx" }) }],
          };
        }
      },
      tool: (name, _description, _shape, handler) => ({ name, handler }),
      createSdkMcpServer: (args) => args,
      backupWorkbookFile: async () => ({ status: "copied", file: "isolated-mock" }),
    },
  );
  const instance = await gateway.createThepExcelGateway(
    { type: "stdio", command: "mock" },
    execution,
  );
  const server = instance.createSessionServer({
    workbookPath: "C:\\audit\\book.xlsx",
    signal: abort.signal,
  });
  const response = server.tools[0].handler({ action: "write" });
  await infoStarted;
  abort.abort();
  const stopped = await response;
  assert.equal(stopped.isError, true);
  assert.deepEqual(calls, ["excel_workbook"]);
  releaseInfo();
  await new Promise(setImmediate);
  assert.deepEqual(calls, ["excel_workbook", "excel_range"]);
  return { responseAtStop: JSON.parse(stopped.content[0].text).code, callsAfterStop: calls };
});

await check("B11-grouped-restore-splits-inverses", async () => {
  const inverses = [];
  const pi = evaluate(
    between(piSource, "function resolveSnapshotKind(", "// src/workbook/recovery/log-store.ts"),
  );
  const groups = evaluate(
    between(recoverySource, "function groupSnapshots(", "async function resolveSnapshotGroup("),
  );
  const snapshots = ["one", "two"].map((id) => ({
    id,
    workbookId: "book",
    toolCallId: "same-operation",
    address: "Sheet1!A1",
    beforeValues: [[1]],
    beforeFormulas: [[1]],
  }));
  for (const snapshot of snapshots)
    await pi.restoreWorkbookRecoverySnapshot({
      snapshot,
      scope: { workbookId: "book" },
      dependencies: {
        toRestoreValues: (values) => values,
        applySnapshot: async () => ({ values: [[2]], formulas: [[2]] }),
        countChangedCells: () => 1,
        appendRangeSnapshot: async (args) => {
          const result = { ...args, id: `inverse-${inverses.length}` };
          inverses.push(result);
          return result;
        },
      },
    });
  assert.equal(groups.groupSnapshots(snapshots).length, 1);
  assert.equal(groups.groupSnapshots(inverses).length, 2);
  return { originalGroups: 1, inverseGroups: 2, inverseCallIds: inverses.map((x) => x.toolCallId) };
});

await check("B12-copy-expanded-range-unguarded", async () => {
  // copyFrom expansion follows Microsoft's documented Range.copyFrom contract.
  // This verifies our guard's decisions, not the real Office host implementation.
  const data = new Map([
    ["0,0", 1],
    ["0,1", 2],
    ["1,0", 3],
    ["1,1", 4],
    ["1,4", "KEEP"],
  ]);
  const ranges = new Map();
  function range(address, row, col, rows, cols) {
    const result = {
      address: `Sheet1!${address}`,
      rowCount: rows,
      columnCount: cols,
      load() {},
      get values() {
        return Array.from({ length: rows }, (_, r) =>
          Array.from({ length: cols }, (_, c) => data.get(`${row + r},${col + c}`) ?? ""),
        );
      },
      get formulas() {
        return this.values;
      },
      copyFrom(from) {
        const values = from.values;
        for (let r = 0; r < from.rowCount; r++)
          for (let c = 0; c < from.columnCount; c++)
            data.set(`${row + r},${col + c}`, values[r][c]);
      },
    };
    ranges.set(address, result);
  }
  range("A1:B2", 0, 0, 2, 2);
  range("D1", 0, 3, 1, 1);
  const office = evaluate(
    between(officeSource, "function columnIndexToLetter(", "function excelColorToHex(") +
      between(officeSource, "async function copyTo(", "async function modifySheetStructure("),
    {
      Excel: { run: (fn) => fn({ sync: async () => {} }), RangeCopyType: { all: "All" } },
      getWorksheetById: async () => ({ getRange: (address) => ranges.get(address) }),
    },
  );
  const result = await office.copyTo(1, "A1:B2", "D1", false);
  assert.equal(result.success, true);
  assert.equal(data.get("1,4"), 4);
  return { allowOverwrite: false, uncheckedCell: "E2", before: "KEEP", after: data.get("1,4") };
});

await check("B13-failed-insert-creates-deletion-checkpoint", async () => {
  const saved = [];
  const recovery = evaluate(
    between(recoverySource, "function structureSpan(", "// Before/after cell diffs"),
    {
      Excel: { run: (fn) => fn({ sync: async () => {} }) },
      resolveWorksheet: async () => ({
        worksheet: { id: "sheet-1", name: "Sheet1", position: 0, visibility: "Visible", load() {} },
      }),
      recoveryLog: {
        appendModifyStructure: async (args) => {
          saved.push(args);
          return { id: "phantom", ...args };
        },
      },
    },
  );
  const commit = await recovery.prepareStructureRecovery(
    "excel_modify_sheet_structure",
    { operation: "insert", dimension: "rows", reference: "2", count: 1 },
    "failed-call",
  );
  // runOfficeTool's catch calls commitRecovery() even if Excel made no change.
  const result = await commit();
  assert.equal(result.status, "checkpoint_created");
  let deleted = false;
  const pi = evaluate(
    between(piSource, "async function applyRowsState(", "async function applyColumnsState("),
    {
      normalizePositiveInteger: (n) => n,
      loadSheetById: async () => ({
        id: "sheet-1",
        name: "Sheet1",
        getRange: () => ({
          delete: () => {
            deleted = true;
          },
        }),
      }),
      hasValueDataInRange: async () => false, // Existing blank row; data below it moves up.
    },
  );
  await pi.applyRowsState({ sync: async () => {} }, saved[0].modifyStructureState);
  assert.equal(deleted, true);
  return {
    failedOperationCheckpoint: result.status,
    inverseKind: saved[0].modifyStructureState.kind,
    deletedExistingBlankRow: deleted,
  };
});

await check("B14-pending-restart-defeats-stop", async () => {
  const pending = [];
  let spawned = 0;
  const app = evaluate(
    between(
      source("app/main.mjs"),
      "async function startDaemon()",
      "// --------------------------------------------------------------------------\n// IPC:",
    ),
    {
      findInitialWorkspace: async () => null,
      currentWorkspace: null,
      daemonStatus: "stopped",
      daemonProcess: null,
      stableTimer: null,
      restartAttempts: 0,
      MAX_RESTART: 5,
      STABLE_AFTER_MS: 30000,
      updateTray() {},
      DAEMON_ENTRY: "mock",
      PROJECT_ROOT: "mock",
      process: { execPath: "mock", env: {} },
      logStream: null,
      handleDaemonMessage() {},
      setTimeout: (fn, ms) => {
        const timer = { fn, ms };
        pending.push(timer);
        return timer;
      },
      clearTimeout: (timer) => {
        timer.cancelled = true;
      },
      spawn: () => {
        spawned++;
        const child = new EventEmitter();
        child.stdout = new EventEmitter();
        child.stderr = new EventEmitter();
        child.kill = () => {};
        return child;
      },
    },
  );
  await app.startDaemon();
  app.daemonProcess.emit("exit", 1, null); // Scheduled auto restart in 1 second.
  app.restartDaemon(); // User chooses "Start daemon" immediately.
  await new Promise(setImmediate);
  app.stopDaemon(); // Stop the explicit new process.
  await pending.find((timer) => timer.ms === 1000).fn();
  assert.equal(spawned, 3);
  return {
    spawned,
    afterExplicitStop: app.daemonStatus,
    reason: "pending auto-restart is never cancelled",
  };
});

const report = JSON.stringify(
  { findings: results.length, reproduced: results.filter((r) => r.reproduced).length, results },
  null,
  2,
);
if (process.argv.includes("--save")) {
  await writeFile(
    new URL("../../docs/systematic-bug-audit-2026-09-21.repro.json", import.meta.url),
    `${report}\n`,
    "utf8",
  );
}
console.log(report);
