import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";

const source = readFileSync(new URL("../taskpane/shared/taskpane.js", import.meta.url), "utf8");

test("Enter submits, but IME candidate confirmation does not", () => {
  const start = source.indexOf('$input.addEventListener("keydown"');
  const end = source.indexOf("// Theme —", start);
  assert.ok(start >= 0 && end > start, "taskpane keydown handler source anchors must exist");

  let handler;
  const submissions = [];
  const sandbox = {
    $input: {
      addEventListener(name, fn) {
        assert.equal(name, "keydown");
        handler = fn;
      },
    },
    $composer: { dispatchEvent: (event) => submissions.push(event.type) },
    Event: class Event {
      constructor(type) {
        this.type = type;
      }
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(source.slice(start, end), sandbox, {
    filename: "taskpane-keydown-under-test.js",
  });

  const invoke = (overrides = {}) => {
    let prevented = false;
    handler({
      key: "Enter",
      shiftKey: false,
      isComposing: false,
      keyCode: 13,
      preventDefault: () => {
        prevented = true;
      },
      ...overrides,
    });
    return prevented;
  };

  assert.equal(invoke(), true);
  assert.deepEqual(submissions, ["submit"]);

  assert.equal(invoke({ isComposing: true, keyCode: 229 }), false);
  assert.equal(invoke({ isComposing: false, keyCode: 229 }), false);
  assert.equal(invoke({ shiftKey: true }), false);
  assert.equal(invoke({ key: "a", keyCode: 65 }), false);
  assert.deepEqual(submissions, ["submit"]);
});

function createTurnHarness({ attached = true, captureFails = false } = {}) {
  const start = source.indexOf("let selectionDebounce = null;");
  const end = source.indexOf("// Theme", start);
  assert.ok(start >= 0 && end > start, "selection/composer source anchors must exist");

  const sent = [];
  const loads = [];
  const replays = [];
  let excelRuns = 0;
  const sandbox = {
    attachSelection: attached,
    lastSelection: { address: "Sheet1!A1", text: "old" },
    wsReady: true,
    turnInFlight: false,
    submitPending: false,
    readOnlyHistory: false,
    queuedTurns: [],
    $turnQueue: null,
    $turnQueueList: null,
    $send: { textContent: "Send" },
    clearTimeout,
    setTimeout,
    Excel: {
      async run(callback) {
        excelRuns++;
        if (captureFails) throw new Error("selection unavailable");
        const range = {
          address: "Sheet1!B2",
          values: [[42]],
          rowCount: 1,
          columnCount: 1,
          load(properties) {
            loads.push(["selection", properties]);
          },
          getCell() {
            return {
              values: [[42]],
              load(properties) {
                loads.push(["first-cell", properties]);
              },
            };
          },
        };
        return callback({ workbook: { getSelectedRange: () => range }, async sync() {} });
      },
    },
    $composer: { addEventListener() {}, dispatchEvent() {} },
    $input: { value: "", addEventListener() {} },
    Event,
    appendEvent() {},
    appendUserMessage() {},
    setAgentStatus() {},
    setComposerDisabled() {},
    refreshSelectionChip() {},
    renderTranscriptReplay(...args) {
      replays.push(args);
    },
    wsSend(message) {
      sent.push(message);
    },
    beginTurn() {
      sandbox.turnInFlight = true;
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(source.slice(start, end), sandbox, {
    filename: "taskpane-selection-under-test.js",
  });
  return { sandbox, sent, loads, replays, excelRuns: () => excelRuns };
}

test("a turn snapshots Excel selection immediately before submit", async () => {
  const h = createTurnHarness();
  assert.equal(await h.sandbox.sendUserTurn("calculate"), true);
  assert.equal(h.excelRuns(), 1);
  const message = h.sent.find((entry) => entry.type === "user_message");
  assert.equal(message.selection.address, "Sheet1!B2");
  assert.equal(message.selection.text, "42");
  assert.deepEqual(h.loads, [
    ["selection", "address, rowCount, columnCount"],
    ["first-cell", "values"],
  ]);
});

test("a detached turn explicitly submits null selection", async () => {
  const h = createTurnHarness({ attached: false });
  assert.equal(await h.sandbox.sendUserTurn("ignore selection"), true);
  assert.equal(h.excelRuns(), 0);
  const message = h.sent.find((entry) => entry.type === "user_message");
  assert.equal(message.selection, null);
});

test("a failed submit-time capture does not reuse stale selection", async () => {
  const h = createTurnHarness({ captureFails: true });
  assert.equal(await h.sandbox.sendUserTurn("calculate"), true);
  const message = h.sent.find((entry) => entry.type === "user_message");
  assert.equal(message.selection, null);
});

test("follow-ups wait for the current turn and keep their submit-time selection", async () => {
  const h = createTurnHarness();
  h.sandbox.turnInFlight = true;
  assert.equal(await h.sandbox.sendUserTurn("next step"), true);
  assert.equal(h.sent.length, 1, "selection context update is the only immediate message");
  assert.equal(h.sandbox.queuedTurns.length, 1);
  assert.equal(h.sandbox.queuedTurns[0].selection.address, "Sheet1!B2");

  h.sandbox.turnInFlight = false;
  assert.equal(h.sandbox.drainTurnQueue(), true);
  const message = h.sent.find((entry) => entry.type === "user_message");
  assert.equal(message.text, "next step");
  assert.equal(message.selection.address, "Sheet1!B2");
});

test("sending from a view-only conversation clears its transcript", async () => {
  const h = createTurnHarness();
  h.sandbox.readOnlyHistory = true;
  assert.equal(await h.sandbox.sendUserTurn("new task"), true);
  assert.equal(h.replays.length, 1);
  assert.equal(h.replays[0][0].length, 0);
  assert.equal(h.replays[0][1], false);
  assert.equal(h.sandbox.readOnlyHistory, false);
  assert.equal(h.sent.find((entry) => entry.type === "user_message").text, "new task");
});

test("auto-context bounds text and shares an unfinished overview read", async () => {
  const start = source.indexOf("const changeTracker = new ChangeTracker();");
  const end = source.indexOf("async function runOfficeTool(msg)", start);
  assert.ok(start >= 0 && end > start);

  let resolveOverview;
  const pendingOverview = new Promise((resolve) => {
    resolveOverview = resolve;
  });
  let overviewCalls = 0;
  const selectionAddresses = [];
  const sandbox = {
    ChangeTracker: class {
      flush() {
        return "changed A1";
      }
    },
    createWorkbookCoordinator: () => ({
      runWrite: async (_context, execute) => ({ result: await execute(), revision: 1 }),
    }),
    buildOverview() {
      overviewCalls++;
      return pendingOverview;
    },
    getWorkbookMetadata: async () => ({ sheetsMetadata: [{ name: "Sheet1", id: 1 }] }),
    async readSelectionContext(address) {
      selectionAddresses.push(address);
      return { text: "S".repeat(15000) + "\nErrors nearby: B3=#DIV/0!" };
    },
    setTimeout,
    clearTimeout,
    Promise,
  };
  vm.createContext(sandbox);
  vm.runInContext(source.slice(start, end), sandbox);
  const first = sandbox.contextSnapshot({ selectionAddress: "Sheet1!B2" });
  const second = sandbox.contextSnapshot({ selectionAddress: "Sheet1!B2" });
  resolveOverview("W".repeat(20000));
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
  const start = source.indexOf("let contextCache = null;");
  const end = source.indexOf("function renderContext()", start);
  assert.ok(start >= 0 && end > start, "context loader source anchors must exist");

  const requests = [];
  const deferred = () => {
    let resolve;
    const promise = new Promise((done) => {
      resolve = done;
    });
    return { promise, resolve };
  };
  const oldRequest = deferred();
  const newRequest = deferred();
  const sandbox = {
    currentWorkspaceCwd: "C:\\old",
    document: { getElementById: () => ({ innerHTML: "" }) },
    renderContext() {},
    showListMessage() {},
    sendRequest() {
      const request = requests.length === 0 ? oldRequest : newRequest;
      requests.push(request);
      return request.promise;
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${source.slice(start, end)}\n` +
      `globalThis.contextApi = { loadContext, switchTo(cwd) { resetContextForWorkspace(cwd); currentWorkspaceCwd = cwd; }, state() { return { contextCache, contextCacheCwd }; } };`,
    sandbox,
    { filename: "taskpane-context-under-test.js" },
  );

  const oldLoad = sandbox.contextApi.loadContext();
  sandbox.contextApi.switchTo("C:\\new");
  const newLoad = sandbox.contextApi.loadContext(true);
  newRequest.resolve({ cwd: "C:\\new", entries: [{ path: "new.txt" }] });
  await newLoad;
  oldRequest.resolve({ cwd: "C:\\old", entries: [{ path: "old.txt" }] });
  await oldLoad;

  const state = sandbox.contextApi.state();
  assert.equal(state.contextCacheCwd, "C:\\new");
  assert.equal(state.contextCache.length, 1);
  assert.equal(state.contextCache[0].path, "new.txt");
});

test("an Excel InvalidArgument on a formula write keeps Excel's own wording", async () => {
  const start = source.indexOf("function hasFormulaInput(args)");
  const end = source.indexOf("const changeTracker = new ChangeTracker();", start);
  assert.ok(start >= 0 && end > start);
  const sandbox = { getWorkbookMetadata: async () => ({ sheetsMetadata: [] }) };
  vm.createContext(sandbox);
  vm.runInContext(source.slice(start, end), sandbox);
  const invalid = Object.assign(new Error("参数无效或缺少，或格式不正确。"), {
    code: "InvalidArgument",
  });

  const formulaWrite = await sandbox.describeOfficeToolError(invalid, {
    cells: [[{ formula: '=IF(C2!="",C2,B2)' }]],
  });
  const valueWrite = await sandbox.describeOfficeToolError(invalid, { cells: [[{ value: 1 }]] });

  // Excel already said what it rejected ("参数无效或缺少"); the note translates
  // the code instead of listing rules Excel never complained about.
  assert.match(formulaWrite, /参数无效或缺少/);
  assert.match(formulaWrite, /argument is missing, extra or of the wrong kind/);
  assert.doesNotMatch(formulaWrite, /<> not !=/);
  assert.equal(valueWrite, invalid.message);
});

test("every mutation receives a uniform receipt with its verification level", () => {
  const start = source.indexOf("const WRITE_TOOLS = new Set(");
  const end = source.indexOf("async function runOfficeTool(msg)", start);
  assert.ok(start >= 0 && end > start);
  const sandbox = {
    createWorkbookCoordinator: () => ({
      runWrite: async (_context, execute) => ({ result: await execute(), revision: 1 }),
    }),
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${source.slice(start, end)}\n` +
      `globalThis.makeReceipt = withMutationReceipt; globalThis.summarizeMutation = mutationSummary;`,
    sandbox,
  );
  const formula = sandbox.makeReceipt(
    "excel_set_cell_range",
    { range: "B2:B5" },
    { success: true, writtenRange: "B2:B5", commitStatus: "committed" },
    "receipt-1",
  );
  assert.equal(formula.verification.status, "read_back");
  assert.deepEqual(Array.from(formula.affectedTargets), ["B2:B5"]);
  assert.equal(formula.verification.semanticCheckRequired, true);
  assert.equal(
    sandbox.summarizeMutation("excel_set_cell_range", {
      cells: [[{ value: 1 }, { formula: "=A1*2" }]],
    }),
    "1\u00d72 cells \u00b7 1 formulas \u00b7 1 values",
  );

  const sort = sandbox.makeReceipt(
    "excel_sort_range",
    { address: "A1:D10" },
    { sheet: "Sheet1", address: "Sheet1!A1:D10" },
    "receipt-2",
  );
  assert.equal(sort.commitStatus, "committed");
  assert.equal(sort.verification.status, "commit_acknowledged");
  assert.deepEqual(Array.from(sort.affectedTargets), ["Sheet1!A1:D10"]);
});

test("approval summaries include the material arguments of each write", () => {
  const start = source.indexOf("const APPROVAL_FIELDS_BY_TOOL = {");
  const end = source.indexOf("function appendApprovalRequest(msg)", start);
  assert.ok(start >= 0 && end > start);
  const sandbox = {};
  vm.createContext(sandbox);
  vm.runInContext(
    `${source.slice(start, end)}\nglobalThis.describeApproval = describeApproval;`,
    sandbox,
  );

  const copy = sandbox.describeApproval("excel_copy_to", {
    sheetId: 1,
    sourceRange: "A1:A5",
    destinationRange: "B1:B5",
  });
  assert.match(copy, /sourceRange: A1:A5/);
  assert.match(copy, /destinationRange: B1:B5/);

  const object = sandbox.describeApproval("excel_modify_object", {
    operation: "update",
    id: "Chart 1",
    properties: { title: "Revenue", chartType: "line" },
  });
  assert.match(object, /id: Chart 1/);
  assert.match(object, /"title":"Revenue"/);
});

test("a failed mutation persists its captured recovery checkpoint", async () => {
  const start = source.indexOf("async function settleRecovery(");
  const end = source.indexOf("// Selection tracking", start);
  assert.ok(start >= 0 && end > start);
  const sent = [];
  let recoveryResult;
  let recoveryFailure;
  const sandbox = {
    cancelledToolCalls: new Set(),
    WRITE_TOOLS: new Set(["excel_set_cell_range"]),
    async prepareMutationRecovery() {
      return async (result, failure) => {
        recoveryResult = result;
        recoveryFailure = failure;
        return { status: "checkpoint_created", snapshotIds: ["before-write"] };
      };
    },
    async setCellRange() {
      throw new Error("format sync failed after values were written");
    },
    isMutationCall: () => true,
    CANCELLED_TOOL_RESULT: Symbol("cancelled-tool-result"),
    async runWorkbookWrite(_id, _name, execute) {
      return { result: await execute(), revision: 1 };
    },
    withMutationReceipt: (_name, _args, result) => result,
    describeOfficeToolError: async (error) => error.message,
    updateToolCardSuccess() {},
    updateToolCardFailure() {},
    takeMutationDiff: () => undefined,
    refuseInCellEditMode: async () => {},
    wsSend: (message) => sent.push(message),
    console: { error() {} },
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${source.slice(start, end)}\nglobalThis.runOfficeTool = runOfficeTool;`,
    sandbox,
  );

  await sandbox.runOfficeTool({
    id: "write-1",
    name: "excel_set_cell_range",
    args: { sheetId: 1, range: "A1", cells: [[{ value: 1 }]] },
  });

  assert.equal(recoveryResult, undefined);
  assert.equal(recoveryFailure.commitStatus, "unknown");
  assert.equal(sent[0].ok, false);
  assert.equal(sent[0].error.commitStatus, "unknown");
  assert.equal(sent[0].error.recovery.status, "checkpoint_created");
});

test("a write refused for cell-edit mode reports that nothing was committed", async () => {
  const start = source.indexOf("const CELL_EDIT_MODE_MESSAGE =");
  const end = source.indexOf("async function runOfficeTool(msg)", start);
  assert.ok(start >= 0 && end > start);
  const editing = Object.assign(new Error("Excel 处于单元格编辑模式。"), {
    code: "InvalidOperationInCellEditMode",
  });
  const sandbox = {
    Excel: {
      async run() {
        throw editing;
      },
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(`${source.slice(start, end)}
globalThis.probe = refuseInCellEditMode;`, sandbox);

  const error = await sandbox.probe().then(() => null, (e) => e);
  assert.equal(error.commitStatus, "not_committed");
  assert.match(error.message, /nothing was changed/);

  // Any other probe failure is left for the write itself to report.
  sandbox.Excel.run = async () => {
    throw new Error("GeneralException");
  };
  assert.equal(await sandbox.probe(), undefined);
});

test("a backup that cannot be saved does not swallow the write's result", async () => {
  const start = source.indexOf("async function settleRecovery(");
  const end = source.indexOf("// Selection tracking", start);
  const sent = [];
  let commits = 0;
  const sandbox = {
    cancelledToolCalls: new Set(),
    WRITE_TOOLS: new Set(["excel_set_cell_range"]),
    async prepareMutationRecovery() {
      return async () => {
        commits += 1;
        throw new Error("QuotaExceededError");
      };
    },
    async setCellRange() {
      return { success: true, commitStatus: "committed" };
    },
    isMutationCall: () => true,
    refuseInCellEditMode: async () => {},
    CANCELLED_TOOL_RESULT: Symbol("cancelled-tool-result"),
    async runWorkbookWrite(_id, _name, execute) {
      return { result: await execute(), revision: 1 };
    },
    withMutationReceipt: (_name, _args, result) => result,
    describeOfficeToolError: async (error) => error.message,
    updateToolCardSuccess() {},
    updateToolCardFailure() {},
    takeMutationDiff: () => undefined,
    wsSend: (message) => sent.push(message),
    console: { error() {} },
  };
  vm.createContext(sandbox);
  vm.runInContext(`${source.slice(start, end)}
globalThis.runOfficeTool = runOfficeTool;`, sandbox);
  await sandbox.runOfficeTool({
    id: "write-1",
    name: "excel_set_cell_range",
    args: { sheetId: 1, range: "A1", cells: [[{ value: 1 }]] },
  });
  assert.equal(sent.length, 1, "exactly one tool_result");
  assert.equal(sent[0].ok, true, "the write happened and is reported");
  assert.equal(sent[0].result.recovery.status, "not_available");
  assert.match(sent[0].result.recovery.reason, /QuotaExceededError/);
  assert.equal(commits, 1, "no second commit on a failure path");
});
