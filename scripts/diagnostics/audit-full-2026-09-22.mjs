// Post-fix audit verification using actual repository functions with isolated
// transport/Office/process fixtures.
import assert from 'node:assert/strict';
import { readFileSync, writeFileSync } from 'node:fs';
import { EventEmitter } from 'node:events';
import { webcrypto } from 'node:crypto';
import vm from 'node:vm';
import { createOfficeBridgeMcp } from '../../daemon/office-tools.mjs';
import { createComputeShell } from '../../daemon/compute-tool.mjs';
import { createController, turnFailed } from '../../taskpane/app/core/controller.js';
import { csvPreview } from '../../taskpane/app/core/csv-preview.js';

const read = (path) => readFileSync(new URL(`../../${path}`, import.meta.url), 'utf8').replace(/\r\n/g, '\n');
const cut = (source, start, end) => {
  const a = source.indexOf(start), b = source.indexOf(end, a + start.length);
  assert.ok(a >= 0 && b > a, `Missing anchors: ${start} / ${end}`);
  return source.slice(a, b);
};
const evaluate = (source, bindings) => {
  const ctx = vm.createContext(bindings);
  vm.runInContext(source, ctx);
  return ctx;
};
const results = [];
async function check(id, fn) {
  try { results.push({ id, fixed: true, evidence: await fn() }); }
  catch (error) { results.push({ id, fixed: false, error: error.stack }); process.exitCode = 1; }
}

await check('F01-valid-quoted-sheet-reference-rejected', async () => {
  const calls = [];
  const server = createOfficeBridgeMcp({ async callTaskpaneTool(name, args) {
    calls.push({ name, args }); return { success: true, commitStatus: 'committed' };
  } }, 'excel', 'audit');
  const handler = server.instance._registeredTools.excel_set_cell_range.handler;
  const dispatches = [];
  for (const formula of ["='Plan (draft'!A1", "='Plan )'!A1", '=Table1[Cost (USD]']) {
    const result = await handler({ sheetId: 1, range: 'B1', cells: [[{ formula }]] });
    assert.notEqual(result.isError, true);
    dispatches.push(formula);
  }
  assert.equal(calls.length, 3);
  return { bridgeCalls: calls.length, dispatches };
});

await check('F02-large-csv-import-stack-overflow', async () => {
  let writes = 0, writtenRows = 0;
  const shell = createComputeShell(async (_name, args) => {
    writes++; writtenRows += args.cells.length;
    return { success: true, commitStatus: 'committed', writtenRange: args.range };
  });
  const result = await shell({
    command: `python3 - <<'PY'
with open('big.csv', 'w') as stream:
    stream.write('x\\n' * 200000)
PY
csv-to-sheet big.csv 1 A1 --force --text`,
  });
  assert.equal(result.exitCode, 0, result.stderr);
  assert.equal(writtenRows, 200000);
  return { inputRows: 200000, writes, writtenRows, exitCode: result.exitCode };
});

await check('F03-config-reload-leaves-turn-in-flight', async () => {
  const sent = [];
  const pane = createController({
    makeBridge: () => ({ ready: true, send: (msg) => { sent.push(msg); return true; }, request: async () => ({ ok: true }) }),
    makeRunner: () => ({ run() {}, cancel() {} }),
    excel: { readSelection: async () => null }, storage: null, session: null,
  });
  await pane.sendUserTurn('first task');
  let cancelled = 0, restarted = 0;
  const source = read('daemon/index.mjs');
  const daemon = evaluate(cut(source, 'async function restartSession(', 'function handleAgentMessage('), {
    sessionFor: () => ({ cwd: 'audit', sessionId: 's1' }),
    cancelPaneSession: () => { cancelled++; },
    console: { log() {} },
    bridge: { sendAssistantEvent: (event) => pane.handleServerMessage({ type: 'assistant_event', ...event }) },
    scheduleSessionStart: () => { restarted++; },
  });
  await daemon.restartSession('audit', 'excel', { reason: 'model_changed' });
  pane.handleServerMessage({ type: 'assistant_event', event: 'session_init', session_id: 's1', model: 'new-model' });
  await pane.sendUserTurn('next task');
  const state = pane.store.getState();
  assert.equal(cancelled, 1);
  assert.equal(restarted, 1);
  assert.equal(state.turnInFlight, true);
  assert.equal(state.queue.length, 0);
  assert.equal(sent.filter((msg) => msg.type === 'user_message').length, 2);
  assert.equal(turnFailed({ subtype: 'success', is_error: true }), true); // N06 fixed in the new pane
  return { cancelled, restarted, turnInFlight: state.turnInFlight, queued: state.queue.map((t) => t.text), dispatched: sent.filter((m) => m.type === 'user_message').map((m) => m.text), N06FixedInNewPane: true };
});

await check('F04-native-path-hash-collides', async () => {
  const source = read('taskpane/shared/recovery.js');
  const ctx = evaluate(cut(source, 'function workbookName(', 'const recoveryLog ='), {
    Office: { context: { document: { url: '' } } }, TextEncoder, crypto: webcrypto,
  });
  const ids = [];
  for (const url of ['C:\\Work\\Budget#A.xlsx', 'C:\\Work\\Budget#B.xlsx']) {
    ctx.Office.context.document.url = url;
    ids.push({ url, ...await ctx.currentWorkbookContext() });
  }
  assert.notEqual(ids[0].workbookId, ids[1].workbookId);
  return { contexts: ids, affectedStore: 'customWorkbookId uses this URL hash without a document-token suffix' };
});

await check('F05-history-truncation-splits-one-restore', async () => {
  const source = read('taskpane/shared/recovery.js');
  const pi = [
    { id: 'group-structure', toolCallId: 'one-operation', at: 2, snapshotKind: 'modify_structure_state', address: 'Sheet1!A1' },
    { id: 'group-values', toolCallId: 'one-operation', at: 1, snapshotKind: 'range_values', address: 'Sheet1!A1' },
  ];
  const custom = Array.from({ length: 119 }, (_, i) => ({ id: `recent-${i}`, toolCallId: `call-${i}`, at: i + 3, snapshotKind: 'custom_state', address: 'Sheet1!A1' }));
  const restored = [];
  const ctx = evaluate(source.slice(source.indexOf('function compactSnapshotGroup(')).replace('export async function workbookHistory', 'async function workbookHistory'), {
    readCustomSnapshots: async () => custom,
    recoveryLog: {
      listForCurrentWorkbook: async () => pi,
      restore: async (id) => { restored.push(id); return { restoredSnapshotId: id, address: 'Sheet1!A1', changedCount: 1 }; },
    },
  });
  const result = await ctx.workbookHistory({ action: 'restore', snapshot_id: 'group-structure' });
  assert.equal(result.success, true);
  assert.deepEqual(new Set(restored), new Set(['group-structure', 'group-values']));
  return { snapshotsStillStored: pi.map((s) => s.id), actuallyRestored: restored, reportedCommitStatus: result.commitStatus };
});

await check('F06-csv-preview-splits-quoted-fields', async () => {
  const result = csvPreview('"a,b",c\n"line1\nline2",d', 'D5:E6');
  assert.equal(result.rows.length, 2);
  assert.equal(result.columns.length, 2);
  return { expectedRows: 2, expectedColumns: 2, actual: result };
});

await check('F07-stop-does-not-cancel-manual-restart', async () => {
  const source = read('app/main.mjs');
  let spawned = 0;
  const ctx = evaluate(cut(source, 'let restartTimer = null;', 'async function findInitialWorkspace') +
    cut(source, 'function startDaemon() {', '// --------------------------------------------------------------------------\n// IPC:'), {
    findInitialWorkspace: async () => null, currentWorkspace: null, daemonStatus: 'stopped', daemonProcess: null,
    stableTimer: null, restartAttempts: 0, MAX_RESTART: 3, STABLE_AFTER_MS: 45000, updateTray() {},
    DAEMON_ENTRY: 'daemon', PROJECT_ROOT: 'root', process: { execPath: 'node', env: {} }, logStream: null,
    handleDaemonMessage() {}, setTimeout: () => ({}), clearTimeout() {},
    spawn() {
      spawned++;
      const proc = new EventEmitter();
      proc.stdout = new EventEmitter(); proc.stderr = new EventEmitter();
      proc.kill = () => true; // Real child-process exit is asynchronous.
      return proc;
    },
  });
  await ctx.startDaemon();
  const old = ctx.daemonProcess;
  ctx.restartDaemon();
  ctx.stopDaemon();
  old.emit('exit', null, 'SIGTERM');
  await new Promise(setImmediate);
  assert.equal(spawned, 1);
  assert.equal(ctx.daemonStatus, 'stopped');
  return { spawnedAfterStop: spawned - 1, finalStatus: ctx.daemonStatus };
});

const report = { generatedAt: new Date().toISOString(), note: 'Post-fix verification using isolated fixtures; no live Excel or model calls.', results };
if (process.argv.includes('--save')) writeFileSync(new URL('../../docs/full-project-audit-2026-09-22.repro.json', import.meta.url), JSON.stringify(report, null, 2) + '\n');
console.log(JSON.stringify(report, null, 2));
