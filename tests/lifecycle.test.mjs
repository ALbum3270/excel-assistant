import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { resolve } from "node:path";
import vm from "node:vm";
import { once } from "node:events";
import WebSocket from "ws";
import { createBridge } from "../daemon/bridge.mjs";
import { ApprovalManager, needsApproval } from "../daemon/approval.mjs";

// index.mjs boots HTTP servers, reads credentials and starts the SDK when
// imported. Evaluate its actual lifecycle/handler code without that boot
// path. Only external I/O and the SDK are replaced; message queues use the
// real bridge. Explicit anchors fail loudly if production code moves.
const source = readFileSync(new URL("../daemon/index.mjs", import.meta.url), "utf8");
function section(from, to) {
  const start = source.indexOf(from);
  const end = source.indexOf(to, start + from.length);
  assert.ok(start >= 0 && end > start, `Missing source anchors: ${from} / ${to}`);
  return source.slice(start, end);
}
const lifecycle = section("const sessions = new Map();", "const samePath =");
const input = section("async function* userMessageStream(", "// Which model provider");
const agent = section("function providerLabel()", "// Kick off.");
const tick = () => new Promise((done) => setImmediate(done));
function deferred() {
  let resolve, reject;
  const promise = new Promise((yes, no) => {
    resolve = yes;
    reject = no;
  });
  return { promise, resolve, reject };
}
async function until(predicate) {
  for (let n = 0; n < 100; n++) {
    if (predicate()) return;
    await tick();
  }
  assert.fail("Lifecycle did not reach the expected state");
}

function harness(t, overrides = {}) {
  const events = [],
    queries = [],
    received = [],
    reads = [],
    writes = [],
    clients = [],
    keys = new Set();
  let bridge, handlers;
  const sandbox = {
    AbortController,
    URL,
    setImmediate,
    resolve,
    samePath: (a, b) =>
      String(a || "")
        .replaceAll("/", "\\")
        .toLowerCase() ===
      String(b || "")
        .replaceAll("/", "\\")
        .toLowerCase(),
    console: { log() {}, warn() {}, error() {} },
    process: { env: {} },
    matterFolder: resolve("fallback"),
    WS_PORT: 0,
    BRIDGE_TOKEN: "test-token",
    HTTP_ORIGIN: "test",
    diag() {},
    userMcpServers: {},
    agentPlugins: [],
    agentConfig: {},
    SESSION_COMPATIBILITY_KEY: "test-compat",
    customPermissionHandler() {},
    ApprovalManager,
    needsApproval,
    touchFolder: async () => {},
    buildSystemPromptAppend: async () => "test prompt",
    createOfficeBridgeMcp: () => ({}),
    getSessionId: async () => null,
    getSessionRecord: async (host, documentKey, sessionId) => {
      const id = sessionId ?? (await sandbox.getSessionId(host, documentKey));
      return id ? { session_id: id, compatibility_key: "test-compat" } : null;
    },
    saveSessionId: async () => {},
    clearSessionId: async () => {},
    listSessions: async () => ({ active_session_id: null, sessions: [] }),
    activateSession: async () => false,
    deleteSession: async () => false,
    deleteTranscript: async () => false,
    resolveWorkspaceRoot: async () => resolve("document-folder"),
    stat: async () => ({ isDirectory: () => true }),
    ensureWorkspaceMarker: async () => false,
    readTranscript: async () => ({ events: [], truncated: false }),
    getContextEntries: async (cwd) => {
      reads.push(cwd);
      return [];
    },
    setContextEntries: async (cwd) => {
      writes.push(cwd);
      return { saved: [], errors: [] };
    },
    query: async function* ({ prompt }) {
      for await (const message of prompt) {
        received.push(message.message.content);
        yield { type: "result", subtype: "success" };
      }
    },
    ...overrides,
    createBridge: (options) => {
      handlers = options;
      bridge = createBridge({ ...options, port: 0, token: "test-token", allowedOrigins: [] });
      bridge.sendAssistantEvent = (event, key) => {
        events.push({ ...event, key });
        sandbox.onAssistantEvent?.(event, key);
      };
      bridge.sendAssistantText = (text, key) => events.push({ event: "text", text, key });
      bridge.sendToTaskpane = (event, key) => events.push({ ...event, key });
      bridge.isTaskpaneConnected = () => true;
      return bridge;
    },
  };
  const queryImpl = sandbox.query;
  sandbox.query = (args) => {
    queries.push(args);
    return queryImpl(args);
  };
  vm.createContext(sandbox);
  vm.runInContext(
    lifecycle +
      input +
      agent +
      `
    globalThis.api = { sessionFor, startSessionForFolder, cancelPaneSession,
      startNewConversation, activateConversation, removeConversation,
      ensureLoopForMessage, scheduleSessionStart,
      cwdForKey, onPaneConnect, awaitWorkspaceResolution, sendTranscriptReplayTo };
  `,
    sandbox,
    { filename: "daemon-lifecycle-under-test.mjs" },
  );
  const api = sandbox.api;
  t.after(async () => {
    for (const key of keys) api.cancelPaneSession(key);
    for (const client of clients) client.terminate();
    await tick();
    await bridge.close();
  });
  const key = "excel\0test-book";
  keys.add(key);
  return {
    sandbox,
    api,
    bridge,
    events,
    queries,
    received,
    reads,
    writes,
    key,
    async connect() {
      await until(() => bridge.address());
      const client = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
      clients.push(client);
      await once(client, "open");
      const welcome = once(client, "message");
      client.send(
        JSON.stringify({
          type: "hello",
          token: "test-token",
          host: "excel",
          active_doc: "test-book",
        }),
      );
      await welcome;
      return client;
    },
    send(text, target = key) {
      keys.add(target);
      const started = handlers.onUserMessage(target, "excel");
      bridge.pushUserMessage(text, target);
      return started;
    },
    async handle(name, msg = {}, target = key) {
      keys.add(target);
      let response;
      await handlers.extraHandlers[name](
        msg,
        (value) => {
          response = value;
        },
        target,
        "excel",
      );
      return response;
    },
  };
}

test("initialization failure ends once; the next message succeeds after recovery", async (t) => {
  let failed = true;
  const h = harness(t, {
    buildSystemPromptAppend: async () => {
      if (failed) throw new Error("ENOENT: prompt missing");
      return "restored";
    },
  });
  await h.send("first");
  await until(() => h.events.some((e) => e.event === "error"));
  assert.equal(h.api.sessionFor(h.key), null);
  assert.equal(h.bridge.hasPendingUserMessages(h.key), false);
  failed = false;
  await h.send("retry");
  await until(() => h.received.length === 1);
  assert.match(h.received[0], /retry$/);
  assert.equal(h.events.filter((e) => e.event === "error").length, 1);
});

test("recents storage failure does not prevent a turn", async (t) => {
  const h = harness(t, {
    touchFolder: async () => {
      throw new Error("EACCES");
    },
  });
  await h.send("works");
  await until(() => h.received.length === 1);
  assert.equal(
    h.events.some((e) => e.event === "error"),
    false,
  );
});

for (const outcome of ["failure", "success"]) {
  test(`obsolete initialization ${outcome} cannot affect its replacement`, async (t) => {
    const oldPrompt = deferred();
    let attempts = 0;
    const h = harness(t, {
      buildSystemPromptAppend: () =>
        ++attempts === 1 ? oldPrompt.promise : Promise.resolve("new prompt"),
    });
    await h.send("old");
    await until(() => attempts === 1);
    await h.handle("stop_agent");
    await h.send("new");
    await until(() => h.received.length === 1);
    const current = h.api.sessionFor(h.key);
    if (outcome === "failure") oldPrompt.reject(new Error("obsolete failure"));
    else oldPrompt.resolve("obsolete prompt");
    await tick();
    await tick();
    assert.equal(h.api.sessionFor(h.key), current);
    assert.equal(h.queries.length, 1);
    assert.equal(
      h.events.some((e) => e.event === "error"),
      false,
    );
    await h.send("next");
    await until(() => h.received.length === 2);
    assert.match(h.received[1], /next$/);
  });
}

test("second consumed turn without a result fails once; third turn retries lazily", async (t) => {
  let h;
  let attempt = 0;
  h = harness(t, {
    query: async function* ({ prompt }) {
      if (++attempt === 1) {
        await prompt.next();
        yield { type: "result", subtype: "success" };
        await prompt.next();
        yield { type: "assistant", message: { content: [] } };
      } else {
        await prompt.next();
        yield { type: "result", subtype: "success" };
      }
    },
  });
  await h.send("first");
  await until(() => h.events.some((e) => e.subtype === "success"));
  await h.send("second");
  await until(() => h.api.sessionFor(h.key) === null);
  assert.equal(h.events.filter((e) => e.event === "error").length, 1);
  assert.equal(h.queries.length, 1, "no background restart after failure");
  await h.send("third");
  await until(() => h.events.filter((e) => e.subtype === "success").length === 2);
  assert.equal(h.queries.length, 2);
});

test("accepted but unconsumed input gets a terminal failure when SDK output ends", async (t) => {
  const end = deferred();
  const h = harness(t, {
    query: async function* ({ prompt }) {
      await prompt.next();
      yield { type: "result", subtype: "success" };
      await end.promise;
    },
  });
  await h.send("first");
  await until(() => h.events.some((e) => e.subtype === "success"));
  await h.send("not consumed");
  assert.equal(h.bridge.hasPendingUserMessages(h.key), true);
  end.resolve();
  await until(() => h.api.sessionFor(h.key) === null);
  assert.equal(h.events.filter((e) => e.event === "error").length, 1);
  assert.equal(h.bridge.hasPendingUserMessages(h.key), false);
});

test("Stop releases the old input waiter; the first retry reaches a fresh consumer", async (t) => {
  const h = harness(t);
  await h.send("first");
  await until(() => h.received.length === 1);
  await tick();
  await h.handle("stop_agent");
  await h.send("retry");
  await until(() => h.received.length === 2);
  assert.match(h.received[1], /retry$/);
  assert.equal(h.queries.length, 2);
  assert.equal(h.events.filter((e) => e.interrupted).length, 1);
});

test("New chat cancels a scheduled old resume before it starts", async (t) => {
  const h = harness(t);
  h.api.scheduleSessionStart(resolve("old"), "old-id", h.key, "excel", "test");
  await h.api.startNewConversation(h.key, "excel");
  await tick();
  assert.equal(h.queries.length, 0);
  await h.send("fresh");
  await until(() => h.queries.length === 1);
  assert.equal(h.queries[0].options.resume, undefined);
});

test("Stop while resolving the document cancels the pending first start", async (t) => {
  const directory = deferred();
  const h = harness(t, { resolveWorkspaceRoot: () => directory.promise });
  const hello = h.api.onPaneConnect(h.key, "excel", "book");
  const message = h.send("cancel me");
  await h.handle("stop_agent");
  directory.resolve(resolve("book"));
  await hello;
  await message;
  await tick();
  assert.equal(h.queries.length, 0);
  await h.send("retry");
  await until(() => h.received.length === 1);
  assert.match(h.received[0], /retry$/);
});

test("cwd/context reads and writes wait for document resolution", async (t) => {
  const directory = deferred();
  const h = harness(t, { resolveWorkspaceRoot: () => directory.promise });
  const hello = h.api.onPaneConnect(h.key, "excel", "book");
  const cwd = h.handle("get_cwd_state"),
    read = h.handle("get_context");
  const write = h.handle("set_context", { entries: [] });
  await tick();
  assert.equal(h.reads.length + h.writes.length, 0);
  directory.resolve(resolve("book"));
  await hello;
  await write;
  assert.equal((await cwd).current_cwd, resolve("book"));
  assert.equal((await read).cwd, resolve("book"));
  assert.deepEqual(h.reads, [resolve("book")]);
  assert.deepEqual(h.writes, [resolve("book")]);
});

test("context writes are rejected after the workspace changes", async (t) => {
  const h = harness(t);
  const result = await h.handle("set_context", {
    entries: [{ path: "stale.txt" }],
    expected_cwd: resolve("old-workspace"),
  });
  assert.equal(result.ok, false);
  assert.match(result.error, /Workspace changed/);
  assert.deepEqual(h.writes, []);
});

test("slow transcript does not block input and cannot overwrite the new message", async (t) => {
  const transcript = deferred();
  let reading = false;
  const h = harness(t, {
    getSessionId: async () => "prior-session",
    readTranscript: () => {
      reading = true;
      return transcript.promise;
    },
  });
  const hello = h.api.onPaneConnect(h.key, "excel", "book");
  await until(() => reading);
  await h.send("new input");
  await until(() => h.received.length === 1);
  transcript.resolve({ events: [{ kind: "user", text: "old" }], truncated: false });
  await hello;
  assert.equal(
    h.events.some((e) => e.type === "transcript_replay"),
    false,
  );
});

test("explicit workspace switch follows hello resolution and stays pinned on reconnect", async (t) => {
  const directory = deferred();
  const h = harness(t, { resolveWorkspaceRoot: () => directory.promise });
  const hello = h.api.onPaneConnect(h.key, "excel", "book");
  const change = h.handle("set_cwd", { cwd: resolve("picked") });
  directory.resolve(resolve("book"));
  await hello;
  assert.equal((await change).ok, true);
  await h.api.onPaneConnect(h.key, "excel", "book");
  assert.equal(h.api.cwdForKey(h.key), resolve("picked"));
  assert.equal(h.queries.length, 0, "changing folders must not launch an idle SDK query");
});

test("cancelling one pane does not drop another pane's input", async (t) => {
  const h = harness(t);
  await h.send("other", "excel\0other-book");
  await h.handle("stop_agent");
  await until(() => h.received.length === 1);
  assert.match(h.received[0], /other$/);
});

test("two workbooks in one workspace resume independent document sessions", async (t) => {
  const lookups = [];
  const h = harness(t, {
    getSessionId: async (host, documentKey) => {
      lookups.push([host, documentKey]);
      return documentKey === "book-a" ? "session-a" : "session-b";
    },
  });
  const keyA = "excel\0book-a";
  const keyB = "excel\0book-b";
  await h.send("from a", keyA);
  await until(() => h.queries.length === 1);
  await h.send("from b", keyB);
  await until(() => h.queries.length === 2);
  assert.deepEqual(lookups, [
    ["excel", "book-a"],
    ["excel", "book-b"],
  ]);
  assert.deepEqual(
    h.queries.map((queryArgs) => queryArgs.options.resume),
    ["session-a", "session-b"],
  );
  assert.equal(h.queries[0].options.cwd, h.queries[1].options.cwd);
});

test("New chat clears only the requesting workbook's saved session", async (t) => {
  const cleared = [];
  const h = harness(t, {
    clearSessionId: async (...args) => cleared.push(args),
  });
  await h.api.startNewConversation("excel\0book-a", "excel");
  assert.deepEqual(cleared, [["excel", "book-a"]]);
});

test("conversation history can list and activate a prior workbook session", async (t) => {
  const activated = [];
  const history = {
    active_session_id: "current",
    sessions: [
      { session_id: "current", title: "Current", compatibility_key: "test-compat" },
      { session_id: "prior", title: "Prior", compatibility_key: "test-compat" },
    ],
  };
  const h = harness(t, {
    listSessions: async () => history,
    activateSession: async (...args) => {
      activated.push(args);
      return true;
    },
    getSessionId: async () => "prior",
    readTranscript: async () => ({
      events: [{ kind: "user", text: "old request" }],
      truncated: false,
    }),
  });
  const listed = await h.handle("list_sessions");
  assert.equal(listed.active_session_id, "current");
  assert.equal(listed.sessions.length, 2);

  const result = await h.handle("activate_session", { session_id: "prior" });
  assert.equal(result.ok, true);
  assert.equal(result.read_only, false);
  assert.deepEqual(activated, [["excel", "test-book", "prior"]]);
  assert.equal(
    h.events.some((event) => event.type === "transcript_replay" && event.session_id === "prior"),
    true,
  );
});

test("an incompatible saved session is viewable but a new message starts fresh", async (t) => {
  const cleared = [];
  const activated = [];
  const h = harness(t, {
    getSessionId: async () => "old-session",
    getSessionRecord: async () => ({ session_id: "old-session", compatibility_key: "old-setup" }),
    listSessions: async () => ({
      active_session_id: "old-session",
      sessions: [{ session_id: "old-session", compatibility_key: "old-setup" }],
    }),
    clearSessionId: async (...args) => cleared.push(args),
    activateSession: async (...args) => activated.push(args),
    readTranscript: async () => ({
      events: [{ kind: "user", text: "previous" }],
      truncated: false,
    }),
  });
  const listed = await h.handle("list_sessions");
  assert.equal(listed.sessions[0].resume_compatible, false);
  const selected = await h.handle("activate_session", { session_id: "old-session" });
  assert.equal(selected.read_only, true);
  assert.deepEqual(activated, []);
  assert.equal(
    h.events.findLast((event) => event.type === "transcript_replay")?.resume_compatible,
    false,
  );
  await h.send("fresh question");
  await until(() => h.received.length === 1);
  assert.equal(h.queries[0].options.resume, undefined);
  assert.deepEqual(cleared, [
    ["excel", "test-book"],
    ["excel", "test-book"],
  ]);
});

test("deleting the active conversation removes its transcript and clears the replay", async (t) => {
  const deleted = [];
  const h = harness(t, {
    listSessions: async () => ({
      active_session_id: "remove-me",
      sessions: [{ session_id: "remove-me", title: "Remove me" }],
    }),
    deleteTranscript: async (sessionId) => {
      deleted.push(["transcript", sessionId]);
      return true;
    },
    deleteSession: async (...args) => {
      deleted.push(["index", ...args]);
      return true;
    },
  });
  const result = await h.handle("delete_session", { session_id: "remove-me" });
  assert.equal(result.ok, true);
  assert.equal(result.transcript_deleted, true);
  assert.deepEqual(deleted, [
    ["transcript", "remove-me"],
    ["index", "excel", "test-book", "remove-me"],
  ]);
  assert.equal(
    h.events.some(
      (event) =>
        event.type === "transcript_replay" &&
        event.session_id === null &&
        event.events.length === 0,
    ),
    true,
  );
});

test("actual WebSocket hello plus immediate message waits for the document folder", async (t) => {
  const directory = deferred();
  const h = harness(t, { resolveWorkspaceRoot: () => directory.promise });
  const client = await h.connect();
  client.send(JSON.stringify({ type: "user_message", text: "from socket" }));
  await tick();
  assert.equal(h.queries.length, 0);
  directory.resolve(resolve("socket-book"));
  await until(() => h.received.length === 1);
  assert.equal(h.queries[0].options.cwd, resolve("socket-book"));
  assert.match(h.received[0], /from socket$/);
});

test("New chat during resume lookup cannot revive the old session id", async (t) => {
  const lookup = deferred();
  let lookups = 0;
  const h = harness(t, {
    getSessionId: () => (++lookups === 1 ? lookup.promise : Promise.resolve(null)),
  });
  const oldMessage = h.send("old");
  await until(() => lookups === 1);
  await h.handle("new_session");
  lookup.resolve("old-id");
  await oldMessage;
  await tick();
  assert.equal(h.queries.length, 0);
  await h.send("fresh");
  await until(() => h.received.length === 1);
  assert.equal(h.queries[0].options.resume, undefined);
});

test("switching directories cancels old initialization and starts only on new input", async (t) => {
  const prompt = deferred();
  let attempts = 0;
  const h = harness(t, {
    buildSystemPromptAppend: () => (++attempts === 1 ? prompt.promise : Promise.resolve("new")),
  });
  await h.send("old");
  await until(() => attempts === 1);
  await h.handle("set_cwd", { cwd: resolve("picked") });
  await h.send("new");
  await until(() => h.received.length === 1);
  prompt.reject(new Error("old prompt failed"));
  await tick();
  assert.equal(h.queries.length, 1);
  assert.equal(h.queries[0].options.cwd, resolve("picked"));
  assert.equal(
    h.events.some((e) => e.event === "error"),
    false,
  );
});

test("SDK error result retains its reason and does not produce a second terminal event", async (t) => {
  const h = harness(t, {
    query: async function* ({ prompt }) {
      await prompt.next();
      yield { type: "result", subtype: "error_max_turns", errors: ["Turn limit reached"] };
    },
  });
  await h.send("request");
  await until(() => h.events.some((e) => e.subtype === "error_max_turns"));
  await tick();
  const terminal = h.events.filter((e) =>
    ["turn_complete", "error", "auth_error"].includes(e.event),
  );
  assert.equal(terminal.length, 1);
  assert.equal(terminal[0].error, "Turn limit reached");
});

test("SDK exception ends the request and a retry has a new consumer", async (t) => {
  let attempt = 0;
  const h = harness(t, {
    query: async function* ({ prompt }) {
      await prompt.next();
      if (++attempt === 1) throw new Error("401 Unauthorized");
      yield { type: "result", subtype: "success" };
    },
  });
  await h.send("first");
  await until(() => h.events.some((e) => e.event === "auth_error"));
  await h.send("retry");
  await until(() => h.events.some((e) => e.subtype === "success"));
  assert.equal(h.queries.length, 2);
});

test("a retry sent synchronously from a stream failure event sees no stale session", async (t) => {
  let attempt = 0,
    h;
  h = harness(t, {
    query: async function* ({ prompt }) {
      await prompt.next();
      if (++attempt === 1) return;
      yield { type: "result", subtype: "success" };
    },
    onAssistantEvent(event) {
      if (event.event === "error" && event.subtype === "stream_ended") h.send("instant retry");
    },
  });
  await h.send("first");
  await until(() => h.events.some((e) => e.subtype === "success"));
  assert.equal(h.queries.length, 2);
});

test("a retry sent synchronously from an SDK exception sees no stale session", async (t) => {
  let attempt = 0,
    h;
  h = harness(t, {
    query: async function* ({ prompt }) {
      await prompt.next();
      if (++attempt === 1) throw new Error("transport failed");
      yield { type: "result", subtype: "success" };
    },
    onAssistantEvent(event) {
      if (event.event === "error" && !event.subtype) h.send("instant retry");
    },
  });
  await h.send("first");
  await until(() => h.events.some((e) => e.subtype === "success"));
  assert.equal(h.queries.length, 2);
});

test("evaluation owns its model before reset and restores the pane model afterward", async () => {
  const reset = deferred();
  const key = "excel\0book.xlsx";
  let canceled = 0;
  const bridge = {
    listPanes: () => [{ key, host: "excel", activeDoc: "book.xlsx" }],
    sendAssistantEvent() {},
    sendAssistantText() {},
    pushUserMessage() {},
    clearUserMessages() {},
  };
  const evalSandbox = {
    bridge,
    modelByKey: new Map([[key, "sonnet"]]),
    ALLOWED_MODELS: new Set(["haiku", "sonnet", "opus"]),
    modelArgFor: (paneKey) => evalSandbox.modelByKey.get(paneKey) ?? "sonnet",
    startNewConversation: () => reset.promise,
    ensureLoopForMessage: async () => {},
    sessionFor: () => ({}),
    cancelPaneSession: () => canceled++,
    locateSessionFile: async () => null,
    performance,
    setTimeout,
    clearTimeout,
    setInterval,
    clearInterval,
  };
  const evalCode = section(
    "const evalObservers = new Map();",
    "// Everything that changes agent behavior",
  );
  vm.createContext(evalSandbox);
  vm.runInContext(
    `${evalCode}\nglobalThis.evalApi = { runEvalPrompt, observer: (key) => evalObservers.get(key) };`,
    evalSandbox,
    { filename: "daemon-eval-model-under-test.mjs" },
  );

  const resultPromise = evalSandbox.evalApi.runEvalPrompt({
    doc: "book.xlsx",
    prompt: "task",
    model: "haiku",
  });
  const observer = evalSandbox.evalApi.observer(key);
  assert.equal(observer.tier, "haiku", "model lock must exist before history reset finishes");
  assert.equal(evalSandbox.modelByKey.get(key), "haiku");
  observer.restoreModel = "opus"; // a UI choice received while the evaluation owns the pane
  reset.resolve();
  await tick();
  bridge.sendAssistantEvent({ event: "session_init", model: "qwen-flash", session_id: "s" }, key);
  bridge.sendAssistantEvent({ event: "turn_complete", subtype: "success" }, key);

  const result = await resultPromise;
  assert.equal(result.model, "qwen-flash");
  assert.equal(evalSandbox.modelByKey.get(key), "opus");
  assert.equal(evalSandbox.evalApi.observer(key), undefined);
  assert.equal(canceled, 1, "temporary evaluation loop must not consume later pane messages");
});
