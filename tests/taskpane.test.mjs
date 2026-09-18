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
  let excelRuns = 0;
  const sandbox = {
    attachSelection: attached,
    lastSelection: { address: "Sheet1!A1", text: "old" },
    wsReady: true,
    turnInFlight: false,
    submitPending: false,
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
  return { sandbox, sent, loads, excelRuns: () => excelRuns };
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
