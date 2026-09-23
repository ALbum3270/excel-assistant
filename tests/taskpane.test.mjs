import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createController, addTurnUsage, turnFailed } from "../taskpane/app/core/controller.js";
import {
  createOfficeRunner,
  describeOfficeToolError,
  refuseInCellEditMode,
  withMutationReceipt,
} from "../taskpane/app/core/office-runner.js";
import { createContextSnapshot, readSelection } from "../taskpane/app/core/excel-tools.js";
import { describeArgs } from "../taskpane/app/core/labels.js";
import { groupTimeline } from "../taskpane/app/core/tool-groups.js";

function memoryStorage() {
  const data = new Map();
  return {
    getItem: (key) => (data.has(key) ? data.get(key) : null),
    setItem: (key, value) => data.set(key, String(value)),
  };
}

function deferred() {
  let resolve;
  const promise = new Promise((done) => {
    resolve = done;
  });
  return { promise, resolve };
}

// A controller wired to an in-memory bridge and Excel. `requests` answers
// bridge.request by type.
function createHarness({
  selection = { address: "Sheet1!B2", text: "42" },
  captureFails = false,
  requests = {},
} = {}) {
  const sent = [];
  let reads = 0;
  const bridge = {
    ready: true,
    token: "t",
    connect: async () => {},
    send(message) {
      sent.push(message);
      return true;
    },
    request(type, payload) {
      const answer = requests[type];
      return typeof answer === "function"
        ? answer(payload)
        : Promise.resolve(answer ?? { ok: true });
    },
  };
  const pane = createController({
    makeBridge: () => bridge,
    makeRunner: () => ({ run() {}, cancel() {} }),
    excel: {
      async readSelection() {
        reads++;
        if (captureFails) throw new Error("selection unavailable");
        return selection;
      },
      navigateToRange: async () => {},
    },
    storage: memoryStorage(),
    session: memoryStorage(),
  });
  const state = () => pane.store.getState();
  return { pane, bridge, sent, state, reads: () => reads };
}

test("a turn snapshots Excel selection immediately before submit", async () => {
  const h = createHarness();
  assert.equal(await h.pane.sendUserTurn("calculate"), true);
  assert.equal(h.reads(), 1);
  const message = h.sent.find((entry) => entry.type === "user_message");
  assert.equal(message.selection.address, "Sheet1!B2");
  assert.equal(message.selection.text, "42");
  assert.equal(h.state().turnInFlight, true);
  assert.equal(h.state().items.at(-1).kind, "user");
});

test("the selection read loads the address, size and first value only", async () => {
  const loads = [];
  const range = {
    address: "Sheet1!B2",
    rowCount: 1,
    columnCount: 1,
    load: (properties) => loads.push(["selection", properties]),
    getCell: () => ({
      values: [[42]],
      load: (properties) => loads.push(["first-cell", properties]),
    }),
  };
  const excel = {
    run: (callback) => callback({ workbook: { getSelectedRange: () => range }, async sync() {} }),
  };
  assert.deepEqual(await readSelection(excel), { text: "42", address: "Sheet1!B2" });
  assert.deepEqual(loads, [
    ["selection", "address, rowCount, columnCount"],
    ["first-cell", "values"],
  ]);
});

test("a detached turn explicitly submits null selection", async () => {
  const h = createHarness();
  h.pane.detachSelection();
  assert.equal(await h.pane.sendUserTurn("ignore selection"), true);
  assert.equal(h.reads(), 0);
  assert.equal(h.sent.find((entry) => entry.type === "user_message").selection, null);
  // The detach choice applies to one turn only.
  assert.equal(h.state().selection.attach, true);
});

test("a failed submit-time capture does not reuse stale selection", async () => {
  const h = createHarness({ captureFails: true });
  assert.equal(await h.pane.sendUserTurn("calculate"), true);
  assert.equal(h.sent.find((entry) => entry.type === "user_message").selection, null);
});

test("follow-ups wait for the current turn and keep their submit-time selection", async () => {
  const h = createHarness();
  await h.pane.sendUserTurn("first");
  h.sent.length = 0;
  assert.equal(await h.pane.sendUserTurn("next step"), true);
  assert.equal(
    h.sent.filter((entry) => entry.type === "user_message").length,
    0,
    "queued, not sent",
  );
  assert.equal(h.state().queue.length, 1);
  assert.equal(h.state().queue[0].selection.address, "Sheet1!B2");

  h.pane.handleServerMessage({
    type: "assistant_event",
    event: "turn_complete",
    subtype: "success",
  });
  const message = h.sent.find((entry) => entry.type === "user_message");
  assert.equal(message.text, "next step");
  assert.equal(message.selection.address, "Sheet1!B2");
  assert.equal(h.state().queue.length, 0);
});

test("a turn that reports is_error shows the error and does not run the queued follow-up", async () => {
  // Additional audit N06: the SDK can report subtype "success" with is_error true.
  const h = createHarness();
  await h.pane.sendUserTurn("first");
  await h.pane.sendUserTurn("queued");
  h.sent.length = 0;
  h.pane.handleServerMessage({
    type: "assistant_event",
    event: "turn_complete",
    subtype: "success",
    is_error: true,
    error: "provider rejected the request",
  });
  assert.equal(h.state().items.at(-1).kind, "error");
  assert.equal(h.sent.filter((entry) => entry.type === "user_message").length, 0);
  assert.equal(h.state().queue.length, 1, "the follow-up waits for the user");

  // A Stop is not a failure: the queued message runs.
  assert.equal(turnFailed({ interrupted: true, is_error: true }), false);
  assert.equal(turnFailed({ subtype: "error_max_turns" }), true);
  assert.equal(turnFailed({ subtype: "success" }), false);
});

test("sending from a view-only conversation clears its transcript", async () => {
  const h = createHarness();
  h.pane.handleServerMessage({
    type: "transcript_replay",
    session_id: "s1",
    resume_compatible: false,
    events: [{ kind: "user", text: "old" }],
  });
  assert.equal(h.state().readOnlyHistory, true);
  assert.equal(await h.pane.sendUserTurn("new task"), true);
  assert.equal(h.state().readOnlyHistory, false);
  assert.deepEqual(
    h.state().items.map((item) => [item.kind, item.text]),
    [["user", "new task"]],
  );
  assert.equal(h.sent.find((entry) => entry.type === "user_message").text, "new task");
});

test("a replayed transcript restores tool receipts", () => {
  const h = createHarness();
  h.pane.handleServerMessage({
    type: "transcript_replay",
    events: [
      { kind: "tool", name: "mcp__office__excel_set_cell_range", id: "t1", input: { range: "A1" } },
      {
        kind: "tool_result",
        id: "t1",
        text: JSON.stringify({ commitStatus: "committed", workbookRevision: 3 }),
      },
      { kind: "tool", name: "mcp__office__excel_copy_to", id: "t2", input: {} },
      { kind: "tool_result", id: "t2", isError: true, text: "failed" },
    ],
  });
  const tools = h.state().items.filter((item) => item.kind === "tool");
  assert.equal(tools[0].state, "success");
  assert.equal(tools[0].write, true);
  assert.equal(tools[0].result.workbookRevision, 3);
  assert.equal(tools[1].state, "error");
});

test("streamed assistant text joins one message until something else is shown", () => {
  const h = createHarness();
  h.pane.handleServerMessage({ type: "assistant_text", delta: "Hel" });
  h.pane.handleServerMessage({ type: "assistant_text", delta: "lo" });
  h.pane.handleServerMessage({
    type: "tool_call",
    id: "t1",
    name: "excel_get_cell_ranges",
    args: {},
  });
  h.pane.handleServerMessage({ type: "assistant_text", delta: "Done" });
  assert.deepEqual(
    h.state().items.map((item) => (item.kind === "assistant" ? item.text : item.kind)),
    ["Hello", "tool", "Done"],
  );
});

test("session cost is read, not summed, across turns", () => {
  let usage = {
    turns: 0,
    input: 0,
    output: 0,
    cacheRead: 0,
    cacheCreate: 0,
    cost: 0,
    priced: false,
    contextWindow: 0,
  };
  usage = addTurnUsage(usage, { input_tokens: 10, output_tokens: 5 }, 0.01, {
    m: { inputTokens: 10, outputTokens: 5, contextWindow: 128000 },
  });
  usage = addTurnUsage(usage, { input_tokens: 20, output_tokens: 5 }, 0.03, {
    m: { inputTokens: 30, outputTokens: 10, contextWindow: 128000 },
  });
  assert.equal(usage.turns, 2);
  assert.equal(usage.cost, 0.03);
  assert.equal(usage.input, 30);
  assert.equal(usage.contextWindow, 128000);
});

test("auto-context bounds text and shares an unfinished overview read", async () => {
  const overview = deferred();
  let overviewCalls = 0;
  const selectionAddresses = [];
  const contextSnapshot = createContextSnapshot({
    getMetadata: async () => ({ sheetsMetadata: [{ name: "Sheet1", id: 1 }] }),
    overview() {
      overviewCalls++;
      return overview.promise;
    },
    async readSelection(address) {
      selectionAddresses.push(address);
      return { text: "S".repeat(15000) + "\nErrors nearby: B3=#DIV/0!" };
    },
    changeTracker: { flush: () => "changed A1" },
  });
  const first = contextSnapshot({ selectionAddress: "Sheet1!B2" });
  const second = contextSnapshot({ selectionAddress: "Sheet1!B2" });
  overview.resolve("W".repeat(20000));
  const [a, b] = await Promise.all([first, second]);
  assert.equal(overviewCalls, 1);
  assert.deepEqual(selectionAddresses, ["Sheet1!B2", "Sheet1!B2"]);
  for (const result of [a, b]) {
    assert.ok(result.workbook.length <= 4010 + 8000);
    assert.ok(result.selection.length <= 12000);
    assert.match(result.selection, /Errors nearby: B3=#DIV\/0!/);
    assert.match(result.workbook, /sheetId for mcp__office__excel_\* tools: Sheet1=1/);
  }
});

test("an old context response cannot overwrite a newly selected workspace", async () => {
  const oldRequest = deferred();
  const newRequest = deferred();
  const calls = [];
  const h = createHarness({
    requests: {
      get_context() {
        calls.push(1);
        return calls.length === 1 ? oldRequest.promise : newRequest.promise;
      },
    },
  });
  h.pane.handleServerMessage({ type: "assistant_event", event: "cwd_changed", cwd: "C:\\old" });
  const oldLoad = h.pane.loadContext();
  h.pane.handleServerMessage({ type: "assistant_event", event: "cwd_changed", cwd: "C:\\new" });
  const newLoad = h.pane.loadContext(true);
  newRequest.resolve({ cwd: "C:\\new", entries: [{ path: "new.txt" }] });
  await newLoad;
  oldRequest.resolve({ cwd: "C:\\old", entries: [{ path: "old.txt" }] });
  await oldLoad;

  const context = h.state().context;
  assert.equal(context.cwd, "C:\\new");
  assert.deepEqual(
    context.entries.map((entry) => entry.path),
    ["new.txt"],
  );
});

test("an Excel InvalidArgument on a formula write keeps Excel's own wording", async () => {
  const invalid = Object.assign(new Error("参数无效或缺少，或格式不正确。"), {
    code: "InvalidArgument",
  });
  const formulaWrite = await describeOfficeToolError(invalid, {
    cells: [[{ formula: '=IF(C2!="",C2,B2)' }]],
  });
  const valueWrite = await describeOfficeToolError(invalid, { cells: [[{ value: 1 }]] });
  // Excel already said what it rejected; the note translates the code instead of
  // listing rules Excel never complained about.
  assert.match(formulaWrite, /参数无效或缺少/);
  assert.match(formulaWrite, /argument is missing, extra or of the wrong kind/);
  assert.doesNotMatch(formulaWrite, /<> not !=/);
  assert.equal(valueWrite, invalid.message);
});

test("every mutation receives a uniform receipt with its verification level", () => {
  const formula = withMutationReceipt(
    "excel_set_cell_range",
    { range: "B2:B5" },
    { success: true, writtenRange: "B2:B5", commitStatus: "committed" },
    "receipt-1",
  );
  assert.equal(formula.verification.status, "read_back");
  assert.deepEqual(formula.affectedTargets, ["B2:B5"]);
  assert.equal(formula.verification.semanticCheckRequired, true);

  const sort = withMutationReceipt(
    "excel_sort_range",
    { address: "A1:D10" },
    { sheet: "Sheet1", address: "Sheet1!A1:D10" },
    "receipt-2",
  );
  assert.equal(sort.commitStatus, "committed");
  assert.equal(sort.verification.status, "commit_acknowledged");
  assert.deepEqual(sort.affectedTargets, ["Sheet1!A1:D10"]);
});

test("argument summaries include the material arguments of each write", () => {
  const copy = describeArgs("excel_copy_to", {
    sheetId: 1,
    sourceRange: "A1:A5",
    destinationRange: "B1:B5",
  });
  assert.match(copy, /sourceRange: A1:A5/);
  assert.match(copy, /destinationRange: B1:B5/);
  const object = describeArgs("mcp__office__excel_modify_object", {
    operation: "update",
    id: "Chart 1",
    properties: { title: "Revenue", chartType: "line" },
  });
  assert.match(object, /id: Chart 1/);
  assert.match(object, /"title":"Revenue"/);
});

function runnerHarness({ execute, commit }) {
  const sent = [];
  const settled = [];
  const runner = createOfficeRunner({
    send: (message) => sent.push(message),
    onSettled: (id, outcome) => settled.push([id, outcome]),
    executeTool: execute,
    prepareMutationRecovery: async () => commit,
    takeMutationDiff: () => undefined,
    runWorkbookWrite: async (_id, _name, run) => ({ result: await run(), revision: 1 }),
    refuseInCellEditMode: async () => {},
    describeError: async (error) => error.message,
    log: { error() {} },
  });
  return { runner, sent, settled };
}

const WRITE = {
  id: "write-1",
  name: "excel_set_cell_range",
  args: { sheetId: 1, range: "A1", cells: [[{ value: 1 }]] },
};

test("a failed mutation persists its captured recovery checkpoint", async () => {
  let recoveryResult;
  let recoveryFailure;
  const h = runnerHarness({
    async execute() {
      throw new Error("format sync failed after values were written");
    },
    async commit(result, failure) {
      recoveryResult = result;
      recoveryFailure = failure;
      return { status: "checkpoint_created", snapshotIds: ["before-write"] };
    },
  });
  await h.runner.run(WRITE);
  assert.equal(recoveryResult, undefined);
  assert.equal(recoveryFailure.commitStatus, "unknown");
  assert.equal(h.sent[0].ok, false);
  assert.equal(h.sent[0].error.commitStatus, "unknown");
  assert.equal(h.sent[0].error.recovery.status, "checkpoint_created");
  assert.equal(h.settled[0][1].ok, false);
});

test("a backup that cannot be saved does not swallow the write's result", async () => {
  let commits = 0;
  const h = runnerHarness({
    execute: async () => ({ success: true, commitStatus: "committed" }),
    async commit() {
      commits += 1;
      throw new Error("QuotaExceededError");
    },
  });
  await h.runner.run(WRITE);
  assert.equal(h.sent.length, 1, "exactly one tool_result");
  assert.equal(h.sent[0].ok, true, "the write happened and is reported");
  assert.equal(h.sent[0].result.recovery.status, "not_available");
  assert.match(h.sent[0].result.recovery.reason, /QuotaExceededError/);
  assert.equal(h.sent[0].result.workbookRevision, 1);
  assert.equal(commits, 1, "no second commit on a failure path");
});

test("a call cancelled before it starts never runs and is reported settled", async () => {
  let executed = false;
  const h = runnerHarness({
    execute: async () => {
      executed = true;
    },
    commit: null,
  });
  h.runner.cancel("write-1", 10);
  await h.runner.run(WRITE);
  assert.equal(executed, false);
  assert.deepEqual(h.sent, [{ type: "tool_settled", id: "write-1" }]);
});

test("a write refused for cell-edit mode reports that nothing was committed", async () => {
  const editing = Object.assign(new Error("Excel 处于单元格编辑模式。"), {
    code: "InvalidOperationInCellEditMode",
  });
  const excel = {
    async run() {
      throw editing;
    },
  };
  const error = await refuseInCellEditMode(excel).then(
    () => null,
    (e) => e,
  );
  assert.equal(error.commitStatus, "not_committed");
  assert.match(error.message, /nothing was changed/);
  // Any other probe failure is left for the write itself to report.
  excel.run = async () => {
    throw new Error("GeneralException");
  };
  assert.equal(await refuseInCellEditMode(excel), undefined);
});

test("consecutive completed calls to one tool group, a running call never does", () => {
  const tool = (id, local, state = "success") => ({ kind: "tool", id, local, state });
  const grouped = groupTimeline([
    tool("a", "excel_get_cell_ranges"),
    tool("b", "excel_get_cell_ranges"),
    tool("c", "excel_get_cell_ranges"),
    tool("d", "excel_set_cell_range"),
    tool("e", "excel_set_cell_range", "running"),
    { kind: "assistant", id: "m" },
    tool("f", "excel_copy_to"),
    tool("g", "excel_copy_to"),
  ]);
  assert.deepEqual(
    grouped.map((entry) =>
      entry.kind === "group"
        ? `${entry.items.length}${entry.collapsed ? "c" : "o"}`
        : entry.item.id,
    ),
    ["3c", "d", "e", "m", "2o"],
  );
});

test("the vendored input box ignores an IME Enter that reports keyCode 229", () => {
  const source = readFileSync(
    new URL("../taskpane/app/vendor/ai-elements/prompt-input.tsx", import.meta.url),
    "utf8",
  );
  assert.match(
    source,
    /if \(isComposing \|\| e\.nativeEvent\.isComposing \|\| e\.nativeEvent\.keyCode === 229\) \{\s*return;/,
  );
  assert.match(source, /if \(e\.shiftKey\) \{\s*return;/);
});

test("the theme follows Excel until the user picks one, and the pick is remembered", () => {
  const storage = memoryStorage();
  const make = () =>
    createController({
      makeBridge: () => ({
        ready: false,
        send() {},
        request: async () => ({}),
        connect: async () => {},
      }),
      makeRunner: () => ({ run() {}, cancel() {} }),
      excel: { readSelection: async () => null },
      storage,
      session: memoryStorage(),
    });
  const first = make();
  assert.equal(first.store.getState().settings.theme, "auto");
  first.setTheme("dark");
  assert.equal(make().store.getState().settings.theme, "dark");
  first.setTheme("bogus");
  assert.equal(make().store.getState().settings.theme, "auto");
});

test("undo and redo toggle on the card of the change, without adding cards", async () => {
  const restores = [];
  let next = 0;
  const h = createHarness({
    requests: {
      workbook_history({ args }) {
        restores.push(args.snapshot_id);
        next += 1;
        return Promise.resolve({ ok: true, result: { inverseSnapshotIds: [`inverse-${next}`] } });
      },
    },
  });
  h.pane.handleServerMessage({
    type: "transcript_replay",
    events: [
      { kind: "tool", name: "mcp__office__excel_set_cell_range", id: "w1", input: { range: "A1" } },
      {
        kind: "tool_result",
        id: "w1",
        text: JSON.stringify({ success: true, recovery: { snapshotIds: ["before-w1"] } }),
      },
    ],
  });
  const card = () => h.state().items.find((item) => item.id === "w1");
  const count = h.state().items.length;

  await h.pane.toggleUndo("w1");
  assert.equal(card().undo.undone, true);
  await h.pane.toggleUndo("w1");
  assert.equal(card().undo.undone, false);
  await h.pane.toggleUndo("w1");
  assert.equal(card().undo.undone, true);
  // Each click restores the reverse point the previous click left.
  assert.deepEqual(restores, ["before-w1", "inverse-1", "inverse-2"]);
  assert.equal(h.state().items.length, count, "no card is added");

  // The pane-started call the daemon sends back runs without a card of its own.
  h.pane.handleServerMessage({
    type: "tool_call",
    id: "p1",
    name: "excel_workbook_history",
    args: { action: "restore" },
    origin: "pane",
  });
  assert.equal(h.state().items.length, count);
});

test("an older daemon's untagged echo of the pane's own restore draws no card", async () => {
  let answer;
  const h = createHarness({
    requests: {
      workbook_history: () => new Promise((resolve) => (answer = resolve)),
    },
  });
  h.pane.handleServerMessage({
    type: "transcript_replay",
    events: [
      { kind: "tool", name: "mcp__office__excel_set_cell_range", id: "w1", input: {} },
      {
        kind: "tool_result",
        id: "w1",
        text: JSON.stringify({ recovery: { snapshotIds: ["s1"] } }),
      },
    ],
  });
  const count = h.state().items.length;
  const undo = h.pane.toggleUndo("w1");
  // While the pane's request is open, the call it causes has no origin tag.
  h.pane.handleServerMessage({
    type: "tool_call",
    id: "echo",
    name: "excel_workbook_history",
    args: { action: "restore", snapshot_id: "s1" },
  });
  answer({ ok: true, result: { inverseSnapshotIds: ["s2"] } });
  await undo;
  assert.equal(h.state().items.length, count);
  // Afterwards, a restore the assistant asks for is shown as usual.
  h.pane.handleServerMessage({
    type: "tool_call",
    id: "agent",
    name: "excel_workbook_history",
    args: { action: "restore", snapshot_id: "s2" },
  });
  assert.equal(h.state().items.length, count + 1);
});
