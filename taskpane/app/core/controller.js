// Pane state and everything the pane does, without any DOM: daemon messages in,
// state out, and the actions the view calls. The React view only renders
// `store.getState()` and calls these actions.
import { createStore } from "./store.js";
import { isMutationCall } from "./office-runner.js";
import { localToolName, statusKeyForTool } from "./labels.js";
import { resolveLanguage, setLanguage, t } from "./i18n.js";
import { defaultPresets, migratePresets, presetText } from "./presets.js";
import { docDirFromActiveUrl, isInOrUnder } from "../../shared/paths.js";

const SETTINGS_KEY = "claude-code-office-settings-v1";
const PRESETS_KEY = "claude-code-office-presets-v1:excel";
const TIERS = ["haiku", "sonnet", "opus"];
export const ARCHIVE_MAX_BYTES = 2_000_000;

function defaultSettings() {
  return {
    showDiagnostics: false,
    // Global, sticky model tier; the daemon maps it to the provider's model.
    model: "sonnet",
    // Ask before each workbook change (approve-before-apply).
    approveWrites: false,
    // "auto" follows Excel's theme; the header toggle pins "light" or "dark".
    theme: "auto",
    // "auto" follows Office's display language.
    language: "auto",
  };
}

function readJson(storage, key) {
  try {
    const raw = storage?.getItem(key);
    return raw ? JSON.parse(raw) : null;
  } catch {
    return null;
  }
}

function writeJson(storage, key, value) {
  try {
    storage?.setItem(key, JSON.stringify(value));
  } catch {
    /* storage full or unavailable: the setting lasts for this session only */
  }
}

function newId(prefix = "i") {
  return (
    globalThis.crypto?.randomUUID?.() ?? `${prefix}_${Math.random().toString(36).slice(2, 12)}`
  );
}

function emptyUsage() {
  return {
    turns: 0,
    input: 0,
    output: 0,
    cacheRead: 0,
    cacheCreate: 0,
    cost: 0,
    priced: false,
    contextWindow: 0,
    last: null,
  };
}

// Model usage is cumulative within one SDK query stream (including subagents and
// compaction). The last-turn figure is main-agent-only.
export function addTurnUsage(previous, usage, cost, modelUsage) {
  if (!usage) return previous;
  const last = {
    input: Number(usage.input_tokens || 0),
    output: Number(usage.output_tokens || 0),
    cacheRead: Number(usage.cache_read_input_tokens || 0),
    cacheCreate: Number(usage.cache_creation_input_tokens || 0),
  };
  const next = { ...previous, turns: previous.turns + 1, last };
  const models = Object.values(modelUsage ?? {});
  if (models.length) {
    const sum = (field) => models.reduce((total, item) => total + Number(item[field] || 0), 0);
    next.input = sum("inputTokens");
    next.output = sum("outputTokens");
    next.cacheRead = sum("cacheReadInputTokens");
    next.cacheCreate = sum("cacheCreationInputTokens");
    next.contextWindow = Math.max(...models.map((item) => Number(item.contextWindow || 0)));
  } else {
    next.input += last.input;
    next.output += last.output;
    next.cacheRead += last.cacheRead;
    next.cacheCreate += last.cacheCreate;
  }
  // SDK result cost is cumulative for one query() stream. Reading the latest
  // value is correct; summing each turn would count earlier turns repeatedly.
  next.cost = Number(cost || 0);
  next.priced = next.cost > 0;
  return next;
}

// One answer to "did this turn succeed?" for both the error shown and whether
// queued follow-ups run. A Stop is not a failure: the next queued message runs.
export function turnFailed(message) {
  return (
    !message.interrupted &&
    Boolean(message.is_error || (message.subtype && message.subtype !== "success"))
  );
}

/**
 * @param {object} deps
 * @param {(handlers) => {connect, send, request, ready}} deps.makeBridge
 * @param {(hooks) => {run, cancel}} deps.makeRunner
 * @param {{readSelection, watchSelection, navigateToRange}} deps.excel
 * @param {{start?: () => void}} [deps.changeTracker]
 * @param {(snapshotArgs) => Promise<any>} [deps.localHistory]  pane-side workbook_history
 */
export function createController({
  makeBridge,
  makeRunner,
  excel,
  changeTracker = null,
  storage = globalThis.localStorage,
  session = globalThis.sessionStorage,
  now = () => Date.now(),
}) {
  const storedSettings = { ...defaultSettings(), ...(readJson(storage, SETTINGS_KEY) ?? {}) };
  if (!TIERS.includes(storedSettings.model)) storedSettings.model = "sonnet";
  const storedPresets = readJson(storage, PRESETS_KEY);
  // Built-in presets are stored by id and shown in the current language;
  // untouched copies saved by older versions become built-ins again.
  const presets = Array.isArray(storedPresets)
    ? migratePresets(storedPresets)
    : defaultPresets(newId);
  writeJson(storage, PRESETS_KEY, presets);

  const store = createStore({
    view: "chat",
    connection: { state: "idle", key: "conn.connecting" },
    liveModel: null,
    pendingModel: null,
    tierModels: {},
    provider: null,
    agent: { state: "idle", key: "agent.ready", since: 0 },
    authError: null,
    hostDark: false,
    hostLanguage: null,
    items: [],
    readOnlyHistory: false,
    selection: { attach: true, last: null },
    queue: [],
    turnInFlight: false,
    submitPending: false,
    usage: emptyUsage(),
    settings: storedSettings,
    presets,
    draft: null,
    workspace: { cwd: null, mismatch: false, error: null, activeDocUrl: null },
    context: { entries: null, cwd: null, loading: false, error: null, dialog: null },
    backups: { snapshots: [], busy: false, status: "", error: false, loaded: false },
    history: { open: false, loading: false, sessions: [], activeId: null, error: null },
  });
  const get = store.getState;
  const set = store.setState;

  // ---- Timeline --------------------------------------------------------------
  let assistantItemId = null;

  function append(item) {
    const entry = { id: item.id ?? newId(), ...item };
    set((state) => ({ ...state, items: [...state.items, entry] }));
    if (entry.kind !== "assistant") assistantItemId = null;
    return entry.id;
  }

  function updateItem(id, update) {
    set((state) => {
      const index = state.items.findIndex((item) => item.id === id);
      if (index < 0) return state;
      const items = state.items.slice();
      items[index] = {
        ...items[index],
        ...(typeof update === "function" ? update(items[index]) : update),
      };
      return { ...state, items };
    });
  }

  function findItem(id) {
    return get().items.find((item) => item.id === id);
  }

  // Pane-written entries carry a key and follow a language switch; daemon text
  // is shown as sent.
  const notice = (key, params) => append({ kind: "notice", key, params });
  const event = (key, params) => append({ kind: "event", key, params });
  const error = (key, params) => append({ kind: "error", key, params });
  const errorText = (text) => append({ kind: "error", text });
  const noticeText = (text) => append({ kind: "notice", text });

  function appendAssistantDelta(delta) {
    if (assistantItemId && findItem(assistantItemId)) {
      updateItem(assistantItemId, (item) => ({ text: item.text + delta }));
    } else {
      assistantItemId = append({ kind: "assistant", text: delta });
    }
  }

  function toolItem(name, args, callId = null) {
    const local = localToolName(name);
    return {
      kind: "tool",
      id: callId ?? newId("tool"),
      callId,
      name,
      local,
      args,
      office: Boolean(callId) || /^mcp__office__/.test(name || ""),
      write: isMutationCall(local, args && typeof args === "object" ? args : {}),
      // Only bridge calls (with an id) report back; others are announcements.
      state: callId ? "running" : "announced",
      result: null,
      error: null,
      diff: null,
      undo: null,
      unlock: null,
    };
  }

  function replayTranscript(events, truncated) {
    const items = [];
    if (truncated) items.push({ kind: "divider", id: newId(), key: "divider.truncated" });
    const tools = new Map();
    for (const ev of events) {
      if (ev.kind === "user") items.push({ kind: "user", id: newId(), text: ev.text });
      else if (ev.kind === "assistant")
        items.push({ kind: "assistant", id: newId(), text: ev.text });
      else if (ev.kind === "tool") {
        const item = { ...toolItem(ev.name, ev.input, ev.id ?? null), replayed: true };
        if (ev.id) tools.set(ev.id, item);
        items.push(item);
      } else if (ev.kind === "tool_result") {
        const item = tools.get(ev.id);
        if (!item) continue;
        if (ev.isError) {
          item.state = "error";
          item.error = { message: ev.text.slice(0, 400) };
          continue;
        }
        item.state = "success";
        try {
          item.result = JSON.parse(ev.text);
        } catch {
          // A truncated or non-JSON result: the card still shows it completed.
        }
      }
    }
    if (events.length > 0) items.push({ kind: "divider", id: newId(), key: "divider.end" });
    assistantItemId = null;
    set({ items });
  }

  // ---- Agent status ------------------------------------------------------------
  function setAgent(state, key, params) {
    set((current) => ({
      ...current,
      agent: {
        state,
        key,
        params,
        since: state === "working" ? current.agent.since || now() : 0,
      },
    }));
  }

  // ---- Bridge ----------------------------------------------------------------
  let panePersistId;
  const paneRuntimeId = newId("runtime");
  try {
    panePersistId = session.getItem("cc-pane-id");
    if (!panePersistId) {
      panePersistId = newId("p");
      session.setItem("cc-pane-id", panePersistId);
    }
  } catch {
    panePersistId = newId("p");
  }

  const bridge = makeBridge({
    onOpen() {
      bridge.send({
        type: "hello",
        token: bridge.token,
        host: "excel",
        active_doc: get().workspace.activeDocUrl,
        pane_id: panePersistId,
        runtime_id: paneRuntimeId,
        selection: get().selection.attach ? get().selection.last : null,
      });
      // Record the sticky model and approval choice right after the hello binds
      // the pane key, so the lazy first-message session start uses them.
      sendModel();
      sendApproval();
    },
    onClose() {
      set({ connection: { state: "err", key: "conn.reconnecting" } });
      // Release the composer if a turn was mid-flight when the connection
      // dropped — the turn_complete it waits for will never arrive.
      endTurn({ drainQueue: false });
    },
    onMessage: handleServerMessage,
  });

  const runner = makeRunner({
    send: (message) => bridge.send(message),
    onSettled(id, outcome) {
      if (outcome.ok) {
        updateItem(id, { state: "success", result: outcome.result, diff: outcome.diff ?? null });
      } else {
        updateItem(id, {
          state: outcome.error?.commitStatus === "unknown" ? "unknown" : "error",
          error: outcome.error,
        });
      }
    },
  });

  const request = (type, payload) => bridge.request(type, payload);

  function sendModel() {
    bridge.send({ type: "set_model", model: get().settings.model });
  }

  function sendApproval() {
    bridge.send({ type: "set_approval", enabled: Boolean(get().settings.approveWrites) });
  }

  // ---- Turns -----------------------------------------------------------------
  function beginTurn() {
    set({ turnInFlight: true });
  }

  function endTurn({ drainQueue = true } = {}) {
    set({ turnInFlight: false });
    if (drainQueue) drainTurnQueue();
  }

  function dispatchCapturedTurn({ text, selection }) {
    append({ kind: "user", text });
    bridge.send({ type: "user_message", text, selection });
    setAgent("working", "tool.working");
    beginTurn();
  }

  function drainTurnQueue() {
    const { turnInFlight, queue } = get();
    if (turnInFlight || !bridge.ready || queue.length === 0) return false;
    const [next, ...rest] = queue;
    set({ queue: rest });
    dispatchCapturedTurn(next);
    return true;
  }

  let selectionTimer = null;

  async function captureSelection() {
    try {
      const selection = await excel.readSelection();
      set((state) => ({ ...state, selection: { ...state.selection, last: selection } }));
      bridge.send({ type: "context_update", selection: get().selection.attach ? selection : null });
      return selection;
    } catch {
      // Selection may be transient; ignore.
      return null;
    }
  }

  function onSelectionChanged() {
    clearTimeout(selectionTimer);
    selectionTimer = setTimeout(captureSelection, 100);
  }

  // Snapshot the selection immediately before the message. This closes the
  // debounce window where Excel has moved but context_update still carries the
  // previous range. Detaching explicitly sends null for this turn.
  async function sendUserTurn(text) {
    text = String(text ?? "").trim();
    if (get().submitPending || !text) return false;
    if (!bridge.ready) {
      event("event.notConnected");
      return false;
    }
    set({ submitPending: true });
    try {
      const selection = get().selection.attach ? await captureSelection() : null;
      if (!bridge.ready) {
        event("event.disconnectedBeforeSend");
        return false;
      }
      if (get().readOnlyHistory) {
        set({ readOnlyHistory: false });
        replayTranscript([], false);
      }
      const turn = { id: newId("q"), text, selection };
      if (get().turnInFlight || get().queue.length > 0) {
        set((state) => ({ ...state, queue: [...state.queue, turn] }));
        drainTurnQueue();
      } else {
        dispatchCapturedTurn(turn);
      }
      // The detach choice applies to one turn; the next turn starts attached.
      set((state) => ({ ...state, selection: { ...state.selection, attach: true } }));
      return true;
    } finally {
      set({ submitPending: false });
    }
  }

  function detachSelection() {
    set((state) => ({ ...state, selection: { ...state.selection, attach: false } }));
    bridge.send({ type: "context_update", selection: null });
  }

  function removeQueued(id) {
    set((state) => ({ ...state, queue: state.queue.filter((turn) => turn.id !== id) }));
  }

  function clearQueue() {
    set({ queue: [] });
  }

  // Stop aborts the current turn; the daemon answers with turn_complete
  // interrupted=true, and the next message starts a fresh resuming loop.
  function stop() {
    if (!bridge.ready || get().agent.state !== "working") return;
    setAgent("working", "agent.stopping");
    bridge.send({ type: "stop_agent" });
  }

  // ---- Server messages ---------------------------------------------------------
  function handleServerMessage(msg) {
    switch (msg.type) {
      case "welcome":
        set({ connection: { state: "ok", key: "conn.connected" } });
        setAgent("idle", "agent.ready");
        refreshWorkspaceFromDaemon();
        refreshModelLabels();
        break;
      case "transcript_replay": {
        const readOnly = Boolean(msg.session_id) && msg.resume_compatible === false;
        set({ readOnlyHistory: readOnly });
        replayTranscript(msg.events || [], Boolean(msg.truncated));
        if (readOnly) {
          notice(msg.archived ? "notice.archive" : "notice.incompatible");
        }
        break;
      }
      case "assistant_text":
        appendAssistantDelta(msg.delta);
        break;
      case "assistant_event":
        handleAssistantEvent(msg);
        break;
      case "tool_call":
        // A call the pane started itself (its undo button) is not something the
        // assistant did, so it runs without a card of its own.
        if (!startedByPane(msg)) append(toolItem(msg.name, msg.args, msg.id));
        runner.run(msg);
        break;
      case "tool_cancel":
        runner.cancel(msg.id);
        updateItem(msg.id, {
          state: "unknown",
          error: {
            messageKey: "tool.cancelled",
            commitStatus: "unknown",
          },
        });
        break;
      default:
        break;
    }
  }

  function handleAssistantEvent(msg) {
    switch (msg.event) {
      case "tool_use_announce":
        // Office tools get a richer card from the matching tool_call, where the
        // bridge id attaches the actual result and receipt.
        if (!String(msg.tool || "").startsWith("mcp__office__"))
          append(toolItem(msg.tool, msg.input));
        setAgent("working", ...statusKeyForTool(msg.tool));
        break;
      case "turn_complete": {
        set((state) => ({
          ...state,
          usage: addTurnUsage(state.usage, msg.usage, msg.total_cost_usd, msg.model_usage),
        }));
        const failed = turnFailed(msg);
        if (failed) {
          if (msg.error) errorText(msg.error);
          else error("error.turnEnded", { subtype: msg.subtype || "error" });
          setAgent("idle", "agent.stoppedError");
        } else {
          setAgent("idle", msg.interrupted ? "agent.stopped" : "agent.ready");
        }
        assistantItemId = null;
        endTurn({ drainQueue: !failed });
        break;
      }
      case "approval_request":
        set((state) => ({
          ...state,
          items: [
            ...state.items.filter((item) => item.id !== `approval:${msg.request_id}`),
            {
              kind: "approval",
              id: `approval:${msg.request_id}`,
              requestId: msg.request_id,
              tool: msg.tool,
              input: msg.input,
              decision: null,
              busy: false,
              error: null,
            },
          ],
        }));
        assistantItemId = null;
        break;
      case "approval_resolved":
        updateItem(`approval:${msg.request_id}`, { decision: msg.decision, busy: false });
        break;
      case "info":
      case "session_incompatible":
        notice(msg.message);
        break;
      case "context_compacting":
        setAgent("working", "agent.compacting");
        break;
      case "context_compacted": {
        const before = Number(msg.pre_tokens || 0).toLocaleString();
        const after = msg.post_tokens ? ` → ${Number(msg.post_tokens).toLocaleString()}` : "";
        event("event.compacted", { before, after });
        setAgent("working", "tool.working");
        break;
      }
      case "context_compaction_complete":
        setAgent("working", "tool.working");
        break;
      case "context_compaction_failed":
        if (msg.error) noticeText(t("notice.compactionFailed", { error: msg.error }));
        else notice("notice.compactionFailed", { error: "?" });
        setAgent("working", "tool.working");
        break;
      case "context_overflow_recovering":
        notice("notice.overflow");
        setAgent("working", "agent.recovering");
        break;
      case "context_overflow_compacted":
        setAgent("working", "agent.retrying");
        break;
      case "error":
        if (msg.error) errorText(msg.error);
        else error("error.noResult");
        setAgent("idle", "agent.stoppedError");
        endTurn({ drainQueue: false });
        break;
      case "auth_error":
        set({ authError: String(msg.error || "-") });
        if (bridge.ready) set({ connection: { state: "err", key: "conn.signin" } });
        endTurn({ drainQueue: false });
        break;
      case "session_init":
        set({
          usage: emptyUsage(),
          readOnlyHistory: false,
          liveModel: msg.model || get().liveModel,
          pendingModel: null,
        });
        event("event.session", { id: msg.session_id?.slice(0, 8), model: msg.model });
        break;
      case "cwd_changed":
        setWorkspaceDisplay(msg.cwd);
        event(msg.resumed ? "event.workspaceResumed" : "event.workspace", {
          name: msg.cwd.split(/[\\/]/).filter(Boolean).pop(),
        });
        if (get().view === "settings") loadContext(true);
        break;
      case "config_reloaded":
        event(msg.reason === "context_changed" ? "event.reloadedContext" : "event.reloadedConfig");
        break;
      default:
        break;
    }
  }

  // ---- Approvals ---------------------------------------------------------------
  async function respondApproval(requestId, decision) {
    const id = `approval:${requestId}`;
    updateItem(id, { busy: true, error: null });
    try {
      const response = await request("approval_response", {
        approval_request_id: requestId,
        decision,
      });
      updateItem(id, {
        busy: false,
        decision: response.decision ?? decision,
        error: response.ok ? null : response.error || t("error.approvalNotDelivered"),
      });
    } catch (failure) {
      updateItem(id, { busy: false, error: failure.message });
    }
  }

  // ---- Restore points and unknown outcomes -------------------------------------
  // The pane's own history requests (undo, redo, the restore-points page) come
  // back from the daemon as an excel_workbook_history tool call. Newer daemons
  // tag it origin "pane"; while such a request is open, an untagged one is
  // recognised too, so an older daemon cannot turn each click into a card.
  let paneHistoryRequests = 0;

  function startedByPane(msg) {
    return (
      msg.origin === "pane" || (msg.name === "excel_workbook_history" && paneHistoryRequests > 0)
    );
  }

  async function daemonWorkbookHistory(args) {
    paneHistoryRequests += 1;
    try {
      const response = await request("workbook_history", { args });
      if (!response.ok) throw new Error(response.error || t("error.history"));
      return response.result;
    } finally {
      paneHistoryRequests -= 1;
    }
  }

  function turnBusy() {
    return get().turnInFlight || get().submitPending;
  }

  // Undo and redo on the card of the change itself. Restoring a restore point
  // creates a reverse one, so each click restores the point the previous click
  // left: undo, redo, undo again, all on the same card.
  function undoPoint(item) {
    if (item.undo) return item.undo.snapshotId;
    return (
      item.result?.recovery?.snapshotIds?.[0] ?? item.error?.recovery?.snapshotIds?.[0] ?? null
    );
  }

  async function toggleUndo(itemId) {
    const item = findItem(itemId);
    const snapshotId = item && undoPoint(item);
    if (!snapshotId) return;
    const undone = Boolean(item.undo?.undone);
    if (turnBusy()) {
      updateItem(itemId, { undo: { snapshotId, undone, error: t("error.restoreWhileBusy") } });
      return;
    }
    updateItem(itemId, { undo: { snapshotId, undone, busy: true, error: null } });
    try {
      const restored = await daemonWorkbookHistory({ action: "restore", snapshot_id: snapshotId });
      updateItem(itemId, {
        undo: {
          snapshotId: restored.inverseSnapshotIds?.[0] ?? null,
          undone: !undone,
          error: null,
        },
      });
      if (get().view === "backups") loadBackups();
    } catch (failure) {
      updateItem(itemId, {
        undo: { snapshotId, undone, error: failure?.message || String(failure) },
      });
    }
  }

  async function acknowledgeWorkbook(itemId) {
    updateItem(itemId, { unlock: { busy: true } });
    try {
      const response = await request("acknowledge_workbook");
      if (!response.ok) throw new Error(response.error || t("error.unlock"));
      updateItem(itemId, { unlock: { busy: false, revision: response.workbookRevision } });
    } catch (failure) {
      updateItem(itemId, { unlock: { busy: false, error: failure?.message || String(failure) } });
    }
  }

  // ---- Backups view ------------------------------------------------------------
  function setBackups(update) {
    set((state) => ({ ...state, backups: { ...state.backups, ...update } }));
  }

  async function loadBackups(status = "") {
    if (get().backups.busy) return;
    setBackups({ busy: true, status: status || t("backups.loading"), error: false });
    try {
      const result = await daemonWorkbookHistory({ action: "list", limit: 120 });
      setBackups({ snapshots: result.snapshots || [], status, loaded: true });
    } catch (failure) {
      setBackups({ status: failure?.message || String(failure), error: true });
    } finally {
      setBackups({ busy: false });
    }
  }

  async function backupAction(action, snapshotId = null) {
    if (get().backups.busy) return;
    if (action === "restore" && turnBusy()) {
      setBackups({ status: t("error.restoreWhileBusy"), error: true });
      return;
    }
    setBackups({
      busy: true,
      error: false,
      status: t(
        { restore: "backups.restoring", delete: "backups.deleting", clear: "backups.clearing" }[
          action
        ],
      ),
    });
    try {
      const result = await daemonWorkbookHistory({ action, snapshot_id: snapshotId });
      const listed = await daemonWorkbookHistory({ action: "list", limit: 120 });
      const done = {
        restore: () =>
          t("backups.restored", { targets: result.addresses?.join(", ") || t("backups.workbook") }),
        delete: () => t("backups.deleted"),
        clear: () => t("backups.cleared", { count: result.removed || 0 }),
      }[action]();
      setBackups({ snapshots: listed.snapshots || [], status: done });
    } catch (failure) {
      setBackups({ status: failure?.message || String(failure), error: true });
    } finally {
      setBackups({ busy: false });
    }
  }

  // ---- Conversation history ---------------------------------------------------
  function setHistory(update) {
    set((state) => ({ ...state, history: { ...state.history, ...update } }));
  }

  async function refreshHistory() {
    setHistory({ loading: true, error: null });
    try {
      const history = await request("list_sessions");
      if (!history.ok) throw new Error(history.error || t("error.historyLoad"));
      setHistory({ sessions: history.sessions || [], activeId: history.active_session_id });
    } catch (failure) {
      setHistory({ error: failure.message });
    } finally {
      setHistory({ loading: false });
    }
  }

  function openHistory() {
    setHistory({ open: true });
    refreshHistory();
  }

  async function activateSession(sessionId) {
    clearQueue();
    const result = await request("activate_session", { session_id: sessionId });
    if (!result.ok) throw new Error(result.error || t("error.sessionOpen"));
    set({ readOnlyHistory: Boolean(result.read_only), usage: emptyUsage() });
    setAgent("idle", "agent.ready");
    setHistory({ open: false });
  }

  async function exportSession(sessionId) {
    const result = await request("export_session", { session_id: sessionId });
    if (!result.ok) throw new Error(result.error || t("error.sessionExport"));
    return result.archive;
  }

  async function deleteSession(sessionId) {
    if (sessionId === get().history.activeId) clearQueue();
    const result = await request("delete_session", { session_id: sessionId });
    if (!result.ok) throw new Error(result.error || t("error.sessionDelete"));
    await refreshHistory();
  }

  async function importSession(file) {
    if (file.size > ARCHIVE_MAX_BYTES) throw new Error(t("error.archiveTooLarge"));
    const archive = JSON.parse(await file.text());
    const result = await request("import_session", { archive });
    if (!result.ok) throw new Error(result.error || t("error.sessionImport"));
    await refreshHistory();
  }

  async function newChat() {
    try {
      clearQueue();
      const result = await request("new_session");
      if (!result.ok) throw new Error(result.error || t("error.newChat"));
      set({ usage: emptyUsage() });
      setAgent("idle", "agent.ready");
    } catch (failure) {
      errorText(failure.message);
    }
  }

  // ---- Settings, model, provider ------------------------------------------------
  function updateSettings(update) {
    const settings = { ...get().settings, ...update };
    set({ settings });
    writeJson(storage, SETTINGS_KEY, settings);
    return settings;
  }

  function setModel(tier) {
    updateSettings({ model: tier });
    sendModel();
    // A live loop restarts to apply the new model; show that until the next
    // session_init confirms. Before any turn the choice simply applies.
    if (get().liveModel) set({ pendingModel: tier });
  }

  function setApproveWrites(enabled) {
    updateSettings({ approveWrites: Boolean(enabled) });
    sendApproval();
  }

  function setShowDiagnostics(enabled) {
    updateSettings({ showDiagnostics: Boolean(enabled) });
  }

  function applyLanguage() {
    setLanguage(resolveLanguage(get().settings.language, get().hostLanguage));
  }

  function setLanguagePreference(language) {
    updateSettings({ language: ["zh", "en"].includes(language) ? language : "auto" });
    applyLanguage();
  }

  function setTheme(theme) {
    updateSettings({ theme: ["light", "dark"].includes(theme) ? theme : "auto" });
  }

  async function refreshModelLabels() {
    try {
      const response = await request("get_models");
      if (!response.ok) return;
      set({
        tierModels: Object.fromEntries(Object.entries(response.models).filter(([, id]) => id)),
        provider: {
          name: response.provider || "Anthropic",
          baseUrl: response.base_url || "",
          models: response.models || {},
          credentialConfigured: Boolean(response.credential_configured),
        },
      });
    } catch {
      /* older daemon without get_models — keep the tier names */
    }
  }

  async function saveProvider(settings) {
    const result = await request("set_provider", { settings });
    if (!result.ok) throw new Error(result.error || t("error.providerSave"));
    const shadowed = result.shadowed?.length
      ? ` ${t("provider.shadowed", { names: result.shadowed.join(", ") })}`
      : "";
    refreshModelLabels();
    return (
      (result.restarting ? t("provider.savedRestarting") : t("provider.savedRestart")) + shadowed
    );
  }

  // ---- Presets -------------------------------------------------------------------
  function savePresets(next) {
    set({ presets: next });
    writeJson(storage, PRESETS_KEY, next);
  }

  function upsertPreset(data, id = null) {
    const list = get().presets;
    savePresets(
      id
        ? list.map((preset) => {
            if (preset.id !== id) return preset;
            const { builtin: _builtin, ...own } = preset;
            return { ...own, ...data };
          })
        : [...list, { id: newId("p"), ...data }],
    );
  }

  function deletePreset(id) {
    savePresets(get().presets.filter((preset) => preset.id !== id));
  }

  function togglePin(id) {
    savePresets(
      get().presets.map((preset) =>
        preset.id === id ? { ...preset, pinned: !preset.pinned } : preset,
      ),
    );
  }

  async function usePreset(preset) {
    set({ view: "chat" });
    const { prompt } = presetText(preset);
    if (preset.auto_send) await sendUserTurn(prompt);
    else set({ draft: { text: prompt, at: now() } });
  }

  // ---- Workspace -------------------------------------------------------------------
  // The workspace is the folder the open document lives in, followed
  // automatically. An explicit pick overrides that until the document moves to a
  // different folder (lastFollowedDocDir), so a reconnect does not undo it.
  let lastFollowedDocDir = null;

  function mismatchFor(cwd, activeDocUrl) {
    if (!activeDocUrl || !cwd) return false;
    // Cloud-hosted documents have no folder to compare, so they never mismatch.
    const docDir = docDirFromActiveUrl(activeDocUrl);
    return Boolean(docDir) && !isInOrUnder(docDir, cwd);
  }

  function setWorkspaceDisplay(cwd) {
    if (get().workspace.cwd !== cwd) resetContextForWorkspace(cwd);
    set((state) => ({
      ...state,
      workspace: {
        ...state.workspace,
        cwd,
        mismatch: mismatchFor(cwd, state.workspace.activeDocUrl),
      },
    }));
  }

  function setWorkspaceError(message) {
    set((state) => ({ ...state, workspace: { ...state.workspace, error: message } }));
  }

  async function refreshWorkspaceFromDaemon() {
    try {
      const response = await request("get_cwd_state");
      if (response.current_cwd) setWorkspaceDisplay(response.current_cwd);
      await autoFollowDocWorkspace();
    } catch {
      /* ignore on initial boot */
    }
  }

  async function loadWorkspaceSection() {
    setWorkspaceError(null);
    await autoFollowDocWorkspace();
    try {
      const response = await request("get_cwd_state");
      if (response.current_cwd) setWorkspaceDisplay(response.current_cwd);
    } catch (failure) {
      setWorkspaceError(t("error.workspaceLoad", { error: failure.message }));
    }
  }

  async function autoFollowDocWorkspace() {
    const { activeDocUrl, cwd } = get().workspace;
    if (!activeDocUrl) return;
    const docDir = docDirFromActiveUrl(activeDocUrl);
    if (!docDir || docDir === lastFollowedDocDir) return;
    lastFollowedDocDir = docDir;
    if (cwd && isInOrUnder(docDir, cwd)) return;
    await switchWorkspace(null, { autodetectFromDoc: true });
  }

  async function switchWorkspace(cwd, { autodetectFromDoc = false } = {}) {
    setWorkspaceError(null);
    try {
      const activeDocUrl = get().workspace.activeDocUrl;
      const response = await request(
        "set_cwd",
        autodetectFromDoc ? { autodetect_from_doc: activeDocUrl } : { cwd },
      );
      if (!response.ok) throw new Error(response.error || t("error.workspaceSwitch"));
      if (!autodetectFromDoc) lastFollowedDocDir = docDirFromActiveUrl(activeDocUrl) || null;
      // The daemon emits cwd_changed, which updates the header. Clear the chat
      // so the new session is visibly separate from the old one.
      set({ items: [] });
      assistantItemId = null;
      loadWorkspaceSection();
    } catch (failure) {
      setWorkspaceError(failure.message);
    }
  }

  async function pickPath({ start_path = null, include_files = false, title = null } = {}) {
    const response = await request("pick_path", { default_path: start_path, include_files, title });
    if (!response.ok) throw new Error(response.error || t("error.picker"));
    if (response.canceled) return null;
    return { path: response.path, kind: response.kind };
  }

  async function changeWorkspace() {
    try {
      const picked = await pickPath({ title: t("workspace.pickTitle") });
      if (picked) await switchWorkspace(picked.path);
    } catch (failure) {
      setWorkspaceError(failure.message);
    }
  }

  // ---- Context files ---------------------------------------------------------------
  // A response for a workspace the pane has since left must not fill (or be
  // saved into) the new one: each load carries a token and the cwd it was for.
  let contextToken = Symbol("context");
  let contextLoading = null;

  function setContext(update) {
    set((state) => ({ ...state, context: { ...state.context, ...update } }));
  }

  function resetContextForWorkspace(cwd) {
    contextToken = Symbol("context");
    contextLoading = null;
    setContext({ entries: null, cwd: cwd ?? null, loading: false, error: null });
  }

  async function loadContext(force = false) {
    const requestCwd = get().workspace.cwd;
    const { entries, cwd } = get().context;
    if (entries && cwd === requestCwd && !force) return;
    if (contextLoading?.cwd === requestCwd && !force) return contextLoading.promise;
    const token = Symbol("context");
    contextToken = token;
    const promise = (async () => {
      setContext({ loading: true, error: null });
      try {
        const response = await request("get_context");
        const currentCwd = get().workspace.cwd;
        if (
          contextToken !== token ||
          currentCwd !== requestCwd ||
          (response.cwd && currentCwd && response.cwd !== currentCwd)
        ) {
          return;
        }
        setContext({
          entries: Array.isArray(response.entries) ? response.entries : [],
          cwd: response.cwd ?? requestCwd,
        });
      } catch (failure) {
        if (contextToken === token && get().workspace.cwd === requestCwd) {
          setContext({ error: t("error.contextLoad", { error: failure.message }) });
        }
      } finally {
        if (contextToken === token) {
          contextLoading = null;
          setContext({ loading: false });
        }
      }
    })();
    contextLoading = { cwd: requestCwd, promise };
    return promise;
  }

  async function saveContext(entries) {
    const targetCwd = get().context.cwd ?? get().workspace.cwd;
    if (!targetCwd || targetCwd !== get().workspace.cwd) return false;
    setContext({ error: null });
    try {
      const response = await request("set_context", {
        entries: entries.map((entry) => ({ ...entry })),
        expected_cwd: targetCwd,
      });
      if (!response.ok) throw new Error(response.error || t("error.contextSave"));
      if (get().workspace.cwd !== targetCwd || get().context.cwd !== targetCwd) return false;
      if (response.errors?.length) {
        setContext({
          error: t("error.contextPartial", {
            entries: response.errors.map((e) => `${e.path} — ${e.error}`).join("; "),
          }),
        });
      }
      setContext({ entries: Array.isArray(response.saved) ? response.saved : entries });
      return true;
    } catch (failure) {
      if (get().workspace.cwd === targetCwd)
        setContext({ error: t("error.contextSaveFailed", { error: failure.message }) });
      return false;
    }
  }

  async function removeContextEntry(index) {
    const entries = get().context.entries;
    if (!entries) return;
    await saveContext(entries.filter((_, i) => i !== index));
  }

  // Pick first (a single-mode native dialog is reliable on every OS), then the
  // dialog collects an optional description.
  async function beginAddContext(includeFiles) {
    try {
      const picked = await pickPath({
        start_path: get().workspace.cwd,
        include_files: includeFiles,
      });
      if (!picked) return;
      setContext({
        dialog: {
          path: picked.path,
          kind: includeFiles ? "file" : "folder",
          cwd: get().workspace.cwd,
          error: null,
        },
      });
    } catch (failure) {
      setContext({ error: failure.message });
    }
  }

  async function browseContextPath(includeFiles, startPath) {
    const picked = await pickPath({
      start_path: startPath || get().workspace.cwd,
      include_files: includeFiles,
    });
    return picked?.path ?? null;
  }

  function closeContextDialog() {
    setContext({ dialog: null });
  }

  async function addContextEntry(path, description) {
    const dialog = get().context.dialog;
    const fail = (message) => setContext({ dialog: { ...dialog, error: message } });
    if (!path) return fail(t("error.pathRequired"));
    const targetCwd = dialog?.cwd;
    const changed = t("error.workspaceChanged");
    if (!targetCwd || get().workspace.cwd !== targetCwd) return fail(changed);
    if (!get().context.entries || get().context.cwd !== targetCwd) await loadContext();
    if (get().workspace.cwd !== targetCwd || get().context.cwd !== targetCwd) return fail(changed);
    const ok = await saveContext([...(get().context.entries ?? []), { path, description }]);
    if (ok) closeContextDialog();
  }

  // ---- Views -------------------------------------------------------------------------
  function setView(view) {
    set({ view });
    if (view === "settings") {
      loadContext();
      loadWorkspaceSection();
    } else if (view === "backups") {
      loadBackups();
    }
  }

  // ---- Boot ----------------------------------------------------------------------------
  function start({ activeDocUrl = null, hostDark = false, hostLanguage = null } = {}) {
    set((state) => ({
      ...state,
      hostDark,
      hostLanguage,
      workspace: { ...state.workspace, activeDocUrl },
    }));
    applyLanguage();
    excel
      .watchSelection?.(onSelectionChanged)
      ?.catch?.((failure) => console.warn("Could not attach Excel selection handler:", failure));
    captureSelection();
    changeTracker?.start?.();
    return bridge.connect();
  }

  return {
    store,
    start,
    // chat
    sendUserTurn,
    stop,
    newChat,
    detachSelection,
    removeQueued,
    respondApproval,
    toggleUndo,
    acknowledgeWorkbook,
    navigate: (address) => excel.navigateToRange(address),
    dismissAuthError: () => set({ authError: null }),
    consumeDraft: () => set({ draft: null }),
    // settings
    setView,
    setModel,
    setApproveWrites,
    setShowDiagnostics,
    setTheme,
    setLanguage: setLanguagePreference,
    saveProvider,
    // presets
    usePreset,
    upsertPreset,
    deletePreset,
    togglePin,
    // history
    openHistory,
    closeHistory: () => setHistory({ open: false }),
    activateSession,
    exportSession,
    deleteSession,
    importSession,
    // backups
    loadBackups,
    restoreBackup: (id) => backupAction("restore", id),
    deleteBackup: (id) => backupAction("delete", id),
    clearBackups: () => backupAction("clear"),
    // workspace and context files
    changeWorkspace,
    loadContext,
    removeContextEntry,
    beginAddContext,
    browseContextPath,
    closeContextDialog,
    addContextEntry,
    // exposed for tests
    handleServerMessage,
    drainTurnQueue,
  };
}
