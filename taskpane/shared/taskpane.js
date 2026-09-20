/* global Office, Excel */

import {
  toolExcelGetSelectedRange,
  toolExcelSelectRange,
  toolExcelSetFormat,
  toolExcelSortRange,
  toolExcelAutoFilter,
  toolExcelCreateTable,
  toolExcelAddTableRows,
} from "./tools-excel.js";
import {
  clearCellRange,
  copyTo,
  getAllObjects,
  getCellRanges,
  getRangeAsCsv,
  getWorkbookMetadata,
  modifyObject,
  modifySheetStructure,
  modifyWorkbookStructure,
  resizeRange,
  searchData,
  setCellRange,
} from "./vendor/office-agents-excel-api.js";
import {
  buildOverview,
  ChangeTracker,
  createExplainFormulaTool,
  createTraceDependenciesTool,
  readSelectionContext,
} from "./vendor/pi-context.js";
import { prepareMutationRecovery, workbookHistory } from "./recovery.js";
import { isInOrUnder, docDirFromActiveUrl } from "./paths.js";
import { marked } from "/npm/marked.esm.js";
import DOMPurify from "/npm/purify.es.mjs";

// pi-for-excel's read-only formula tools, used as-is. They report failures in
// their text output rather than throwing, so the text is the whole result.
const explainFormulaTool = createExplainFormulaTool();
const traceDependenciesTool = createTraceDependenciesTool();

async function runPiTool(piTool, toolCallId, args) {
  const output = await piTool.execute(toolCallId, args ?? {});
  return { text: output.content.map((part) => part.text ?? "").join("\n") };
}

// Daemon endpoints. The HTTP server that loaded this taskpane is on
// HTTP_PORT; the WebSocket bridge is on WS_PORT (one less by daemon convention).
const WS_URL = "ws://127.0.0.1:47833";
const TOKEN_URL = "/bridge-token";

// The bridge token — fetched at boot from the same-origin HTTP server. The
// bridge rejects any WS that doesn't present this token in its first hello.
let bridgeToken = null;

async function fetchBridgeToken() {
  try {
    const res = await fetch(TOKEN_URL, { credentials: "same-origin" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    bridgeToken = (await res.text()).trim();
  } catch (e) {
    console.warn("Failed to fetch bridge token:", e);
    bridgeToken = null;
  }
}

let ws = null;
let wsReady = false;
let lastSelection = null;
let activeDocUrl = null;

// The only supported host. Office.onReady rejects anything else.
const HOST = "excel";

// ---------------------------------------------------------------------------
// UI helpers
// ---------------------------------------------------------------------------
const $messages = document.getElementById("messages");
// Two status indicators, semantically separate:
//   $connectionStatus (topbar) — WS bridge / daemon reachability + auth.
//     Stays put while the user is reading the chat history.
//   $agentStatus (above composer) — agent activity: Ready / Working —
//     <tool>… / Stopped. Lives right where the user's eye is when they
//     send a message.
const $connectionStatus = document.getElementById("connection-status");
const $agentStatus = document.getElementById("agent-status");
const $stopAgent = document.getElementById("stop-agent");
const $activeDoc = document.getElementById("active-doc");
const $input = document.getElementById("input");
const $send = document.getElementById("send");
const $composer = document.getElementById("composer");
const $chip = document.getElementById("selection-chip");
const $chipText = document.getElementById("selection-chip-text");
const $chipDetach = document.getElementById("selection-chip-detach");
const $turnQueue = document.getElementById("turn-queue");
const $turnQueueList = document.getElementById("turn-queue-list");

let assistantTurnElem = null;
let attachSelection = true;

// True while a user turn is mid-flight (we've sent user_message and are
// waiting for turn_complete / error / auth_error). Additional submissions
// stay in queuedTurns so streaming output keeps the correct assistant bubble.
let turnInFlight = false;
let submitPending = false;
const queuedTurns = [];

function setComposerDisabled(disabled) {
  $send.disabled = disabled;
  $input.disabled = disabled;
}

function beginTurn() {
  turnInFlight = true;
  setComposerDisabled(false);
  $send.textContent = "Queue";
}

function endTurn({ drainQueue = true } = {}) {
  flushAssistantRendering();
  turnInFlight = false;
  setComposerDisabled(false);
  $send.textContent = "Send";
  if (drainQueue) drainTurnQueue();
}

// The connection status doubles as the live-model indicator. Horizontal
// space in the topbar is tight (the workspace chip takes the rest), so when
// connected we show ONLY the short model name and let the colored dot carry
// connection state; the full SDK model id goes in the tooltip. The model
// shown is the trustworthy `session_init` value from the daemon — never the
// agent's own (unreliable) self-report.
let connState = "idle";
let connLabel = "Connecting…";
let liveModel = null; // SDK-reported model id of the running loop
let pendingModelShort = null; // target short name while a switch restarts

// Tier alias → real model id when the daemon runs a non-Claude provider.
let tierModels = {};

function shortModelName(idOrAlias) {
  const s = String(idOrAlias || "");
  if (tierModels[s]) return tierModels[s];
  if (/opus/i.test(s)) return "Opus";
  if (/sonnet/i.test(s)) return "Sonnet";
  if (/haiku/i.test(s)) return "Haiku";
  if (!/^claude-/i.test(s)) return s.split("[")[0];
  return s.replace(/^claude-/, "").split(/[-[]/)[0] || s;
}

const TIER_HINTS = { haiku: "fastest", sonnet: "balanced", opus: "most capable" };

async function refreshModelLabels() {
  try {
    const r = await sendRequest("get_models");
    if (!r.ok) return;
    tierModels = Object.fromEntries(Object.entries(r.models).filter(([, id]) => id));
    for (const option of document.querySelectorAll("#composer-model option")) {
      const id = tierModels[option.value];
      const tier = option.value.charAt(0).toUpperCase() + option.value.slice(1);
      option.textContent = `${id ?? tier} · ${TIER_HINTS[option.value]}`;
    }
    renderConnection();
  } catch {
    /* older daemon without get_models — keep the static labels */
  }
}

function renderConnection() {
  let text = connLabel;
  let title = "Connection to the Excel Assistant daemon";
  if (connState === "ok") {
    if (pendingModelShort) {
      text = `↻ ${pendingModelShort}`;
      title = `Switching model → ${pendingModelShort}…`;
    } else if (liveModel) {
      text = shortModelName(liveModel);
      title = `Connected · ${liveModel}`;
    } else {
      // Connected but no turn has run yet — show what the next message
      // will use (the sticky dropdown choice) rather than a bare "Connected".
      text = shortModelName(settings?.model || "sonnet");
      title = `Connected · ${text} on next message`;
    }
  }
  $connectionStatus.className = `status ${connState}`;
  $connectionStatus.textContent = text;
  $connectionStatus.title = title;
}

function setConnectionStatus(state, label) {
  connState = state;
  connLabel = label;
  renderConnection();
}

function setAgentStatus(state, label) {
  $agentStatus.className = `agent-status ${state}`;
  $agentStatus.textContent = label;
  // Stop button is meaningful only while the agent is mid-turn.
  $stopAgent.hidden = state !== "working";
}

// Stop button — abort the current agent turn. The daemon picks up the
// abort, emits turn_complete with interrupted=true (flipping this
// indicator to "Stopped"). The next message starts a fresh resuming loop.
$stopAgent?.addEventListener("click", () => {
  if (!wsReady) return;
  setAgentStatus("working", "Stopping…");
  wsSend({ type: "stop_agent" });
});

// New chat — the daemon drops this pane's conversation and replays an empty
// transcript (clearing the panel); the next message starts a fresh session.
document.getElementById("new-chat")?.addEventListener("click", async () => {
  try {
    clearTurnQueue();
    const r = await sendRequest("new_session");
    if (!r.ok) throw new Error(r.error || "Could not start a new chat");
    setAgentStatus("idle", "Ready");
  } catch (e) {
    appendError(e.message);
  }
});

const $historyModal = document.getElementById("history-modal");
const $historyList = document.getElementById("history-list");
const $historyClose = document.getElementById("history-modal-close");

function closeHistory() {
  if ($historyModal) $historyModal.hidden = true;
}

function historyTimestamp(value) {
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "" : date.toLocaleString();
}

function historyMessage(text, error = false) {
  if (!$historyList) return;
  $historyList.innerHTML = "";
  const message = document.createElement("div");
  message.className = error ? "history-error" : "history-empty";
  message.textContent = text;
  $historyList.appendChild(message);
}

function renderConversationHistory(history) {
  if (!$historyList) return;
  $historyList.innerHTML = "";
  if (!history.sessions?.length) {
    historyMessage("No saved conversations for this workbook.");
    return;
  }

  for (const session of history.sessions) {
    const active = session.session_id === history.active_session_id;
    const item = document.createElement("div");
    item.className = `history-item${active ? " active" : ""}`;

    const titleRow = document.createElement("div");
    titleRow.className = "history-title-row";
    const title = document.createElement("div");
    title.className = "history-title";
    title.textContent = session.title || "Conversation";
    title.title = title.textContent;
    titleRow.appendChild(title);
    if (active) {
      const badge = document.createElement("span");
      badge.className = "history-active";
      badge.textContent = "Active";
      titleRow.appendChild(badge);
    }

    const meta = document.createElement("div");
    meta.className = "history-meta";
    meta.textContent = historyTimestamp(session.last_used);

    const actions = document.createElement("div");
    actions.className = "history-actions";
    const resume = document.createElement("button");
    resume.type = "button";
    resume.textContent = active ? "Current" : "Continue";
    resume.disabled = active;
    resume.addEventListener("click", async () => {
      resume.disabled = true;
      try {
        clearTurnQueue();
        const result = await sendRequest("activate_session", { session_id: session.session_id });
        if (!result.ok) throw new Error(result.error || "Could not open conversation");
        setAgentStatus("idle", "Ready");
        closeHistory();
      } catch (error) {
        resume.disabled = false;
        historyMessage(error.message, true);
      }
    });

    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "Delete";
    let confirmDelete = false;
    remove.addEventListener("click", async () => {
      if (!confirmDelete) {
        confirmDelete = true;
        remove.textContent = "Delete?";
        remove.classList.add("history-delete-confirm");
        return;
      }
      for (const button of actions.querySelectorAll("button")) button.disabled = true;
      try {
        if (active) clearTurnQueue();
        const result = await sendRequest("delete_session", { session_id: session.session_id });
        if (!result.ok) throw new Error(result.error || "Could not delete conversation");
        await refreshConversationHistory();
      } catch (error) {
        historyMessage(error.message, true);
      }
    });

    actions.append(resume, remove);
    item.append(titleRow, meta, actions);
    $historyList.appendChild(item);
  }
}

async function refreshConversationHistory() {
  historyMessage("Loading...");
  try {
    const history = await sendRequest("list_sessions");
    if (!history.ok) throw new Error(history.error || "Could not load conversation history");
    renderConversationHistory(history);
  } catch (error) {
    historyMessage(error.message, true);
  }
}

document.getElementById("chat-history")?.addEventListener("click", () => {
  if (!$historyModal) return;
  $historyModal.hidden = false;
  refreshConversationHistory();
});
$historyClose?.addEventListener("click", closeHistory);
$historyModal?.addEventListener("click", (event) => {
  if (event.target === $historyModal) closeHistory();
});

// Auth-failure banner. Shown across the top of the panel when the daemon
// emits event: "auth_error". Persists until the user dismisses it; recovery
// is to sign in to Claude Code (or set ANTHROPIC_API_KEY) and relaunch the
// app, neither of which we can do from inside the taskpane.
function showAuthErrorBanner(rawError) {
  let banner = document.getElementById("auth-error-banner");
  if (!banner) {
    banner = document.createElement("div");
    banner.id = "auth-error-banner";
    banner.className = "auth-error-banner";
    banner.innerHTML = `
      <div class="auth-error-head">
        <strong>Sign-in required</strong>
        <button type="button" class="auth-error-dismiss" title="Dismiss">×</button>
      </div>
      <div class="auth-error-body">
        The agent couldn't authenticate with its model provider.
        <ul>
          <li>Using a provider configured in <code>.env</code>: check <code>ANTHROPIC_AUTH_TOKEN</code> and <code>ANTHROPIC_BASE_URL</code>.</li>
          <li>Using Claude: sign in to Claude Code in a terminal (run <code>claude</code>), or set <code>ANTHROPIC_API_KEY</code>.</li>
        </ul>
        Then quit Excel Assistant (tray icon) and reopen it.
      </div>
      <details class="auth-error-raw">
        <summary>Raw error</summary>
        <pre></pre>
      </details>
    `;
    document.body.insertBefore(banner, document.body.firstChild);
    banner.querySelector(".auth-error-dismiss").addEventListener("click", () => {
      banner.hidden = true;
    });
  }
  banner.querySelector(".auth-error-raw pre").textContent = String(rawError || "(no detail)");
  banner.hidden = false;
}

// Map a tool name to a short, user-friendly status label shown in the topbar
// while the agent is mid-turn. Falls back to a generic "Working…" for tools
// the user hasn't seen named before — most filesystem/Bash/MCP tools.
const TOOL_STATUS_LABELS = {
  // Excel task-pane tools
  excel_get_workbook_metadata: "Reading the workbook…",
  excel_get_selected_range: "Reading your selection…",
  excel_get_cell_ranges: "Reading cells…",
  excel_get_range_as_csv: "Reading cells…",
  excel_search_data: "Searching…",
  excel_get_all_objects: "Listing charts and pivots…",
  excel_explain_formula: "Explaining a formula…",
  excel_trace_dependencies: "Tracing formula dependencies…",
  excel_set_cell_range: "Writing cells…",
  excel_clear_cell_range: "Clearing cells…",
  excel_copy_to: "Copying cells…",
  excel_modify_sheet_structure: "Changing rows or columns…",
  excel_modify_workbook_structure: "Changing sheets…",
  excel_resize_range: "Resizing…",
  excel_modify_object: "Updating a chart or pivot…",
  excel_select_range: "Selecting cells…",
  excel_set_format: "Formatting cells…",
  excel_sort_range: "Sorting…",
  excel_autofilter: "Filtering…",
  excel_create_table: "Creating a table…",
  excel_add_table_rows: "Adding table rows…",
  excel_bash: "Running a calculation…",
  // Common Claude Code tools
  Read: "Reading a file…",
  Write: "Writing a file…",
  Edit: "Editing a file…",
  MultiEdit: "Editing a file…",
  Bash: "Running a command…",
  Glob: "Searching files…",
  Grep: "Searching files…",
  WebFetch: "Fetching from the web…",
  WebSearch: "Searching the web…",
};
// Raw MCP server name → user-facing display name. The in-process bridge
// server is named "office" in our code; to the user it's Excel.
function mcpServerDisplayName(raw) {
  if (raw === "office") return "Excel";
  return raw.charAt(0).toUpperCase() + raw.slice(1);
}
function statusForTool(name) {
  if (TOOL_STATUS_LABELS[name]) return TOOL_STATUS_LABELS[name];
  // External and in-process MCP tools both surface with the
  // mcp__<server>__<tool> convention. Strip the prefix and retry the map
  // — catches our own office_*/excel_* tools if the SDK ever wraps them.
  const inner = /^mcp__[^_]+__(.+)$/.exec(name || "");
  if (inner && TOOL_STATUS_LABELS[inner[1]]) return TOOL_STATUS_LABELS[inner[1]];
  const m = /^mcp__([^_]+)__/.exec(name || "");
  if (m) return `Calling ${mcpServerDisplayName(m[1])}…`;
  return "Working…";
}

// Autoscroll is "pinned" only while the user is at (or near) the bottom of
// the chat. The previous behavior unconditionally snapped to bottom on
// every append (streaming deltas, events, tool chips, replay), which made
// scrolling up to read mid-stream impossible — the next token snapped you
// back. nearBottom uses a 48px tolerance so a small natural-scroll wobble
// while reading the latest message still counts as pinned.
let scrollPinned = true;
function nearBottom() {
  return $messages.scrollHeight - $messages.scrollTop - $messages.clientHeight <= 48;
}
function maybeScrollToBottom() {
  if (scrollPinned) $messages.scrollTop = $messages.scrollHeight;
}
function forceScrollToBottom() {
  scrollPinned = true;
  $messages.scrollTop = $messages.scrollHeight;
}
$messages.addEventListener(
  "scroll",
  () => {
    scrollPinned = nearBottom();
  },
  { passive: true },
);

function appendUserMessage(text) {
  const el = document.createElement("div");
  el.className = "msg user";
  el.textContent = text;
  $messages.appendChild(el);
  // The user just sent a message — they want to see what comes next.
  // Override any prior scroll-up.
  forceScrollToBottom();
  assistantTurnElem = null;
}

// Model output is untrusted: render Markdown, then sanitize before inserting.
function renderMarkdown(el, text) {
  el.rawText = text;
  el.innerHTML = DOMPurify.sanitize(marked.parse(text, { gfm: true, breaks: true }));
}

let assistantRenderTimer = null;
const pendingAssistantRenders = new Set();

function flushAssistantRendering() {
  clearTimeout(assistantRenderTimer);
  assistantRenderTimer = null;
  for (const el of pendingAssistantRenders) {
    if (el.isConnected) renderMarkdown(el, el.rawText ?? "");
  }
  pendingAssistantRenders.clear();
  maybeScrollToBottom();
}

function appendAssistantDelta(delta) {
  if (!assistantTurnElem) {
    assistantTurnElem = document.createElement("div");
    assistantTurnElem.className = "msg assistant";
    $messages.appendChild(assistantTurnElem);
  }
  assistantTurnElem.rawText = (assistantTurnElem.rawText ?? "") + delta;
  pendingAssistantRenders.add(assistantTurnElem);
  if (assistantRenderTimer === null) {
    assistantRenderTimer = setTimeout(flushAssistantRendering, 50);
  }
}

function appendEvent(text) {
  const el = document.createElement("div");
  el.className = "msg event";
  el.textContent = text;
  $messages.appendChild(el);
  maybeScrollToBottom();
  assistantTurnElem = null;
}

// Like appendEvent, but a user-facing notice that is always visible (not
// gated by the "Show diagnostics" toggle, which hides .msg.event).
function appendNotice(text) {
  const el = document.createElement("div");
  el.className = "msg notice";
  el.textContent = text;
  $messages.appendChild(el);
  maybeScrollToBottom();
  assistantTurnElem = null;
}

// An always-visible alert bubble for errors that stop the turn (usage-limit
// hit, stream ended without a result, etc.). NOT routed through appendEvent
// because .msg.event is hidden when diagnostics-off — a user hitting their
// limit would otherwise see nothing, the turn would silently stop, and the
// status would flip back to "Ready". This bubble is rendered with a
// prominent border and a ⚠ prefix and is never hidden by the toggle.
function appendError(text) {
  const el = document.createElement("div");
  el.className = "msg error";
  el.textContent = text;
  $messages.appendChild(el);
  maybeScrollToBottom();
  assistantTurnElem = null;
}

// Approve-before-apply card: what the assistant is about to change, with
// approve / approve the rest of this turn / reject.
const APPROVAL_FIELDS_BY_TOOL = {
  excel_set_cell_range: [
    "sheetId",
    "range",
    "cells",
    "copyToRange",
    "resizeWidth",
    "resizeHeight",
    "allow_overwrite",
  ],
  excel_fill_formula: ["sheetId", "range", "formula", "allow_overwrite"],
  excel_copy_to: ["sheetId", "sourceRange", "destinationRange", "allow_overwrite"],
  excel_clear_cell_range: ["sheetId", "range", "clearType"],
  excel_modify_sheet_structure: [
    "sheetId",
    "operation",
    "dimension",
    "reference",
    "count",
    "position",
  ],
  excel_modify_workbook_structure: ["operation", "sheetId", "sheetName", "newName", "tabColor"],
  excel_resize_range: ["sheetId", "range", "width", "height"],
  excel_modify_object: ["sheetId", "operation", "objectType", "id", "properties"],
  excel_set_format: [
    "sheet",
    "address",
    "number_format",
    "bold",
    "italic",
    "font_size",
    "font_name",
    "font_color",
    "fill_color",
    "border",
  ],
  excel_sort_range: ["sheet", "address", "key", "ascending", "has_headers"],
  excel_autofilter: ["sheet", "address", "clear"],
  excel_create_table: ["sheet", "address", "name", "has_headers"],
  excel_add_table_rows: ["table", "values", "index"],
  excel_workbook_history: ["action", "snapshot_id"],
  excel_bash: ["command"],
};

function approvalValue(value) {
  if (Array.isArray(value)) {
    const rows = value.length;
    const columns = Array.isArray(value[0])
      ? Math.max(0, ...value.slice(0, 20).map((row) => row.length))
      : null;
    const sample = JSON.stringify(value.slice(0, 3));
    const shape = columns === null ? `${rows} items` : `${rows}×${columns}`;
    return `${shape}; sample: ${sample.length > 500 ? sample.slice(0, 497) + "..." : sample}`;
  }
  if (value && typeof value === "object") {
    const json = JSON.stringify(value);
    return json.length > 500 ? json.slice(0, 497) + "..." : json;
  }
  const text = String(value);
  return text.length > 500 ? text.slice(0, 497) + "..." : text;
}

function describeApproval(tool, input = {}) {
  const preferred = APPROVAL_FIELDS_BY_TOOL[tool] ?? ["action"];
  const fields = [
    ...preferred,
    ...Object.keys(input).filter((field) => !preferred.includes(field)),
  ];
  const lines = fields
    .filter((field) => input[field] !== undefined && input[field] !== "")
    .map((field) => `${field}: ${approvalValue(input[field])}`);
  return lines.length ? lines.join("\n") : "No arguments";
}

function resolveApprovalCard(requestId, decision, error = null) {
  const el = $messages.querySelector(`[data-approval-request-id="${requestId}"]`);
  if (!el) return;
  const labels = {
    approve: "Approved",
    approve_turn: "Approved for the rest of this turn",
    reject: "Rejected",
    timeout: "Approval timed out",
    cancelled: "Approval cancelled",
    disabled: "Approval turned off; change allowed",
  };
  el.querySelector(".approval-actions").textContent =
    error || labels[decision] || "Approval closed";
}

function appendApprovalRequest(msg) {
  $messages.querySelector(`[data-approval-request-id="${msg.request_id}"]`)?.remove();
  const el = document.createElement("div");
  el.className = "msg approval";
  el.dataset.approvalRequestId = msg.request_id;
  el.innerHTML = `<div class="tool-name"></div><div class="tool-args"></div>
    <div class="approval-actions">
      <button type="button" class="btn-primary btn-small" data-decision="approve">Approve</button>
      <button type="button" class="btn-secondary btn-small" data-decision="approve_turn">Approve rest of turn</button>
      <button type="button" class="btn-secondary btn-small" data-decision="reject">Reject</button>
    </div>`;
  el.querySelector(".tool-name").textContent = `Approve change? ${statusForTool(msg.tool)}`;
  el.querySelector(".tool-args").textContent = describeApproval(msg.tool, msg.input);
  el.querySelector(".approval-actions").addEventListener("click", async (event) => {
    const decision = event.target?.dataset?.decision;
    if (!decision) return;
    for (const button of el.querySelectorAll("button")) button.disabled = true;
    try {
      const response = await sendRequest("approval_response", {
        approval_request_id: msg.request_id,
        decision,
      });
      resolveApprovalCard(msg.request_id, response.decision, response.ok ? null : response.error);
    } catch (error) {
      for (const button of el.querySelectorAll("button")) button.disabled = false;
      el.querySelector(".tool-args").textContent =
        `${describeApproval(msg.tool, msg.input)}\n\n${error.message}`;
    }
  });
  $messages.appendChild(el);
  maybeScrollToBottom();
  assistantTurnElem = null;
}

const toolCards = new Map();

function appendToolUse(name, args, id = null) {
  const el = document.createElement("div");
  el.className = "msg tool";
  el.innerHTML = `
    <div class="tool-head">
      <div class="tool-name"></div>
      <span class="tool-state running">Running</span>
    </div>
    <div class="tool-args"></div>
    <div class="tool-result" hidden></div>`;
  const officeName = /^mcp__office__(.+)$/.exec(name || "")?.[1];
  const localName = officeName || name;
  el.querySelector(".tool-name").textContent =
    id || officeName ? `Excel · ${recoveryOperationLabel(localName)}` : `Tool · ${name}`;
  el.querySelector(".tool-name").title = name;
  const argText = typeof args === "string" ? args : JSON.stringify(args, null, 2);
  el.querySelector(".tool-args").textContent =
    argText.length > 200 ? argText.slice(0, 197) + "..." : argText;
  if (id) {
    el.dataset.toolCallId = id;
    toolCards.set(id, el);
  } else {
    el.querySelector(".tool-state").remove();
  }
  $messages.appendChild(el);
  maybeScrollToBottom();
  assistantTurnElem = null;
  return el;
}

// A complete assistant bubble (replay path — full text, not streamed
// deltas). Resets assistantTurnElem so a subsequent live delta starts a
// fresh bubble rather than appending onto a replayed one.
function toolResultRow(label, value) {
  const row = document.createElement("div");
  row.className = "tool-result-row";
  const key = document.createElement("span");
  key.textContent = label;
  const content = document.createElement("strong");
  content.textContent = value;
  row.append(key, content);
  return row;
}

function setToolCardState(id, state, label) {
  const card = toolCards.get(id);
  if (!card) return null;
  const status = card.querySelector(".tool-state");
  status.className = `tool-state ${state}`;
  status.textContent = label;
  return card;
}

function appendToolCardError(card, message) {
  const result = card?.querySelector(".tool-result");
  if (!result) return;
  result.hidden = false;
  result.classList.add("error");
  const detail = document.createElement("div");
  detail.className = "tool-result-error";
  detail.textContent = message;
  result.appendChild(detail);
}

function appendToolRestoreButton(card, result, id, snapshotId) {
  const restore = document.createElement("button");
  restore.type = "button";
  restore.className = "tool-undo";
  restore.textContent = "Restore backup";
  restore.addEventListener("click", async () => {
    if (turnInFlight || submitPending) {
      appendToolCardError(
        card,
        "Wait for the assistant to finish or stop the turn before restoring.",
      );
      return;
    }
    restore.disabled = true;
    restore.textContent = "Restoring...";
    try {
      const restored = await workbookHistory({ action: "restore", snapshot_id: snapshotId });
      setToolCardState(id, "restored", "Restored");
      restore.textContent = "Restored";
      const targets = restored.addresses?.join(", ") || "the workbook";
      result.append(toolResultRow("Restore", `${targets}; reverse backup saved`));
      if (document.body.dataset.activeTab === "backups") loadRecoveryHistory();
    } catch (error) {
      restore.disabled = false;
      restore.textContent = "Restore backup";
      appendToolCardError(card, error?.message || String(error));
    }
  });
  result.appendChild(restore);
}

function updateToolCardSuccess(id, name, args, receipt) {
  const card = setToolCardState(id, "success", "Completed");
  if (!card || !isMutationCall(name, args)) return;
  const result = card.querySelector(".tool-result");
  result.hidden = false;
  result.classList.remove("error");
  result.innerHTML = "";

  const commitLabels = {
    committed: "Committed",
    not_committed: "Not committed",
    unknown: "Unknown",
  };
  const verificationLabels = {
    read_back: "Read back from Excel",
    commit_acknowledged: "Accepted by Excel",
  };
  result.append(
    toolResultRow(
      "Status",
      commitLabels[receipt?.commitStatus] || receipt?.commitStatus || "Completed",
    ),
  );
  const summary = mutationSummary(name, args);
  if (summary) result.append(toolResultRow("Change", summary));
  if (receipt?.affectedTargets?.length) {
    result.append(toolResultRow("Range", receipt.affectedTargets.join(", ")));
  }
  if (receipt?.verification?.status) {
    result.append(
      toolResultRow(
        "Verification",
        verificationLabels[receipt.verification.status] || receipt.verification.status,
      ),
    );
  }
  if (receipt?.formulaErrorCount || receipt?.formulaErrors?.length) {
    result.append(
      toolResultRow(
        "Formula errors",
        String(receipt.formulaErrorCount || receipt.formulaErrors.length),
      ),
    );
  }

  const snapshotId = receipt?.recovery?.snapshotIds?.[0];
  if (snapshotId) {
    result.append(toolResultRow("Backup", "Saved before this change"));
    appendToolRestoreButton(card, result, id, snapshotId);
  } else if (receipt?.recovery?.status === "not_available") {
    result.append(toolResultRow("Backup", "Not available for this change"));
  }
  maybeScrollToBottom();
}

function updateToolCardFailure(id, error) {
  const commitUnknown = error?.commitStatus === "unknown";
  const card = setToolCardState(
    id,
    commitUnknown ? "unknown" : "error",
    commitUnknown ? "Check workbook" : "Failed",
  );
  if (!card) return;
  appendToolCardError(card, error?.message || String(error));
  if (error?.recovery?.snapshotIds?.length) {
    const result = card.querySelector(".tool-result");
    result.append(toolResultRow("Backup", "Saved before the attempted change"));
    appendToolRestoreButton(card, result, id, error.recovery.snapshotIds[0]);
  }
}

function appendAssistantMessage(text) {
  const el = document.createElement("div");
  el.className = "msg assistant";
  renderMarkdown(el, text);
  $messages.appendChild(el);
  maybeScrollToBottom();
  assistantTurnElem = null;
}

// Rebuild the chat panel from a replayed transcript. Clears whatever is in
// #messages, renders each event through the same bubble helpers the live
// path uses (so the diagnostics CSS gate applies identically), and — only
// when there is prior history — appends a divider so replayed history is
// visually distinct from the live session that follows.
function renderTranscriptReplay(events, truncated) {
  $messages.innerHTML = "";
  toolCards.clear();
  assistantTurnElem = null;

  if (truncated) {
    const t = document.createElement("div");
    t.className = "transcript-truncated";
    t.textContent = "⋯ earlier messages not shown";
    $messages.appendChild(t);
  }

  for (const ev of events) {
    if (ev.kind === "user") appendUserMessage(ev.text);
    else if (ev.kind === "assistant") appendAssistantMessage(ev.text);
    else if (ev.kind === "tool") appendToolUse(ev.name, ev.input);
  }

  if (events.length > 0) {
    const d = document.createElement("div");
    d.className = "transcript-divider";
    d.textContent = "end of earlier conversation";
    $messages.appendChild(d);
    // End of a replay is an explicit "take me to latest" — the user just
    // resumed, they want the cursor at the live tail.
    forceScrollToBottom();
  }
}

function refreshSelectionChip() {
  if (attachSelection && lastSelection && lastSelection.text) {
    const preview =
      lastSelection.text.length > 60 ? lastSelection.text.slice(0, 57) + "..." : lastSelection.text;
    $chipText.textContent = `Selection: "${preview}"`;
    $chip.hidden = false;
  } else {
    $chip.hidden = true;
  }
}

$chipDetach.addEventListener("click", () => {
  attachSelection = false;
  refreshSelectionChip();
  if (wsReady) wsSend({ type: "context_update", selection: null });
});

// ---------------------------------------------------------------------------
// Settings — persisted to localStorage. New settings get added here and
// applied via applySettings().
// ---------------------------------------------------------------------------
const SETTINGS_KEY = "claude-code-office-settings-v1";

function defaultSettings() {
  return {
    showDiagnostics: false,
    // Global, sticky. Cheaper models use less of your monthly Claude
    // programmatic credit. "default" defers to the Claude Code CLI config.
    model: "sonnet", // "haiku" | "sonnet" | "opus" | "default"
    // Ask before each workbook change (approve-before-apply). Off by default.
    approveWrites: false,
  };
}

function loadSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_KEY);
    if (!raw) return defaultSettings();
    const s = { ...defaultSettings(), ...JSON.parse(raw) };
    // Migrate the retired "default" model choice (and any stale value) to
    // the explicit Sonnet default so the dropdown/indicator stay valid.
    if (!["haiku", "sonnet", "opus"].includes(s.model)) s.model = "sonnet";
    return s;
  } catch {
    return defaultSettings();
  }
}

function saveSettings(s) {
  try {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(s));
  } catch {
    /* ignore */
  }
}

let settings = loadSettings();

function applySettings() {
  $messages.dataset.showDiagnostics = String(settings.showDiagnostics);
  const $showDiag = document.getElementById("setting-show-diagnostics");
  if ($showDiag) $showDiag.checked = settings.showDiagnostics;
  const $model = document.getElementById("composer-model");
  if ($model) $model.value = settings.model || "sonnet";
  const $approve = document.getElementById("setting-approve-writes");
  if ($approve) $approve.checked = Boolean(settings.approveWrites);
}

function sendApproval() {
  if (wsReady) wsSend({ type: "set_approval", enabled: Boolean(settings.approveWrites) });
}

document.getElementById("setting-approve-writes")?.addEventListener("change", (e) => {
  settings.approveWrites = e.target.checked;
  saveSettings(settings);
  applySettings();
  sendApproval();
});

// Push the chosen model to the daemon. The SDK model is fixed per agent
// loop, so changing it mid-conversation triggers a resuming restart
// (daemon side); on first connect it's just recorded for the lazy start.
function sendModel() {
  if (wsReady) wsSend({ type: "set_model", model: settings.model || "sonnet" });
}

document.getElementById("setting-show-diagnostics").addEventListener("change", (e) => {
  settings.showDiagnostics = e.target.checked;
  saveSettings(settings);
  applySettings();
});

document.getElementById("composer-model")?.addEventListener("change", (e) => {
  settings.model = e.target.value;
  saveSettings(settings);
  applySettings();
  sendModel();
  // If a loop is already live (a turn has run), the daemon will restart it
  // to apply the new model — show that until the next session_init confirms.
  // If nothing has run yet, renderConnection already shows the new choice.
  if (liveModel) pendingModelShort = shortModelName(settings.model);
  renderConnection();
});

applySettings();

// ---------------------------------------------------------------------------
// WebSocket
// ---------------------------------------------------------------------------
function wsConnect() {
  setConnectionStatus("idle", "Connecting…");
  ws = new WebSocket(WS_URL);

  ws.onopen = () => {
    wsReady = true;
    sendHello();
    // Record the sticky model for this pane key right after the hello
    // binds it, so the lazy first-message session start uses it.
    sendModel();
    sendApproval();
  };

  ws.onclose = () => {
    wsReady = false;
    for (const pending of pendingRequests.values()) {
      clearTimeout(pending.timer);
      pending.reject(new Error("Disconnected from daemon"));
    }
    pendingRequests.clear();
    setConnectionStatus("err", "Disconnected — retrying…");
    // Release the composer if a turn was mid-flight when the connection
    // dropped — otherwise the user is stuck waiting for a turn_complete
    // that will never arrive.
    endTurn({ drainQueue: false });
    // Daemon may have been restarted, in which case it has rotated the
    // bridge token. Re-fetch from /bridge-token before each reconnect
    // attempt so the next hello carries the current token. fetchBridgeToken
    // tolerates the HTTP server being briefly unreachable too (sets the
    // token to null, the hello fails, we loop again).
    setTimeout(() => {
      fetchBridgeToken().then(wsConnect);
    }, 1500);
  };

  ws.onerror = (err) => {
    console.warn("WS error:", err);
  };

  ws.onmessage = (evt) => {
    let msg;
    try {
      msg = JSON.parse(evt.data);
    } catch {
      return;
    }
    handleServerMessage(msg);
  };
}

function wsSend(obj) {
  if (!wsReady) return;
  ws.send(JSON.stringify(obj));
}

// Stable per-pane-instance id. sessionStorage persists for the lifetime
// of this task pane (survives WS reconnects and panel reloads, cleared
// when the pane is closed), so an unsaved/cloud doc with no filesystem
// path still gets a STABLE bridge key across the 1.5s reconnect loop —
// without this the bridge minted a fresh random key every retry and
// leaked per-pane state. Falls back to an in-memory id if
// sessionStorage is unavailable.
let panePersistId;
try {
  panePersistId = sessionStorage.getItem("cc-pane-id");
  if (!panePersistId) {
    panePersistId = crypto?.randomUUID?.() ?? "p_" + Math.random().toString(36).slice(2, 12);
    sessionStorage.setItem("cc-pane-id", panePersistId);
  }
} catch {
  panePersistId = crypto?.randomUUID?.() ?? "p_" + Math.random().toString(36).slice(2, 12);
}

function sendHello() {
  wsSend({
    type: "hello",
    token: bridgeToken,
    host: HOST,
    active_doc: activeDocUrl,
    pane_id: panePersistId,
    selection: attachSelection ? lastSelection : null,
  });
}

async function handleServerMessage(msg) {
  switch (msg.type) {
    case "welcome":
      setConnectionStatus("ok", "Connected");
      setAgentStatus("idle", "Ready");
      refreshWorkspaceFromDaemon();
      refreshModelLabels();
      break;

    case "transcript_replay":
      renderTranscriptReplay(msg.events || [], !!msg.truncated);
      break;

    case "assistant_text":
      appendAssistantDelta(msg.delta);
      break;

    case "assistant_event":
      if (msg.event === "tool_use_announce") {
        // Office tools get a richer card from the matching tool_call below,
        // where the bridge id lets us attach the actual result and receipt.
        if (!String(msg.tool || "").startsWith("mcp__office__")) {
          appendToolUse(msg.tool, msg.input);
        }
        setAgentStatus("working", statusForTool(msg.tool));
      } else if (msg.event === "turn_complete") {
        if (!msg.interrupted && msg.subtype && msg.subtype !== "success") {
          appendError(msg.error || `The agent ended this request with ${msg.subtype}.`);
          setAgentStatus("idle", "Stopped — see message");
        } else {
          setAgentStatus("idle", msg.interrupted ? "Stopped" : "Ready");
        }
        endTurn({
          drainQueue: Boolean(msg.interrupted) || !msg.subtype || msg.subtype === "success",
        });
      } else if (msg.event === "approval_request") {
        appendApprovalRequest(msg);
      } else if (msg.event === "approval_resolved") {
        resolveApprovalCard(msg.request_id, msg.decision);
      } else if (msg.event === "info") {
        appendNotice(msg.message);
      } else if (msg.event === "context_compacting") {
        setAgentStatus("working", "Compacting context...");
      } else if (msg.event === "context_compacted") {
        const before = Number(msg.pre_tokens || 0).toLocaleString();
        const after = msg.post_tokens ? ` to ${Number(msg.post_tokens).toLocaleString()}` : "";
        appendEvent(`Context compacted (${before}${after} tokens).`);
        setAgentStatus("working", "Working...");
      } else if (msg.event === "context_compaction_complete") {
        setAgentStatus("working", "Working...");
      } else if (msg.event === "context_compaction_failed") {
        appendNotice(`Context compaction failed: ${msg.error || "unknown error"}`);
        setAgentStatus("working", "Working...");
      } else if (msg.event === "error") {
        // Always-visible bubble (NOT .msg.event, which the diagnostics
        // toggle hides). Status reads "Stopped — see message" so the
        // user knows the turn ended on this error, not normally.
        appendError(msg.error || "The agent stopped without producing a result.");
        setAgentStatus("idle", "Stopped — see message");
        endTurn({ drainQueue: false });
      } else if (msg.event === "auth_error") {
        showAuthErrorBanner(msg.error);
        if (wsReady) setConnectionStatus("err", "Sign-in required");
        endTurn({ drainQueue: false });
      } else if (msg.event === "session_init") {
        appendEvent(`Session ${msg.session_id?.slice(0, 8)}… (${msg.model})`);
        // Authoritative: this is the model the SDK actually started with.
        liveModel = msg.model || liveModel;
        pendingModelShort = null;
        renderConnection();
      } else if (msg.event === "cwd_changed") {
        setWorkspaceDisplay(msg.cwd);
        appendEvent(
          `Switched to workspace: ${msg.cwd.split(/[\\/]/).filter(Boolean).pop()}${msg.resumed ? " (resumed prior session)" : ""}`,
        );
        if (document.body.dataset.activeTab === "setup") loadContext(true);
      } else if (msg.event === "config_reloaded") {
        const what = msg.reason === "context_changed" ? "context files" : "config";
        appendEvent(`Session reloaded — ${what} updated.`);
      }
      break;

    case "tool_call":
      appendToolUse(msg.name, msg.args, msg.id);
      runOfficeTool(msg);
      break;

    case "tool_cancel":
      cancelledToolCalls.add(msg.id);
      updateToolCardFailure(msg.id, {
        message: "The tool call was cancelled; check the workbook before retrying.",
        commitStatus: "unknown",
      });
      setTimeout(() => cancelledToolCalls.delete(msg.id), 65_000);
      break;

    case "pong":
      break;

    default:
      // Request/response messages keyed by request_id end in "_result".
      // Resolve the matching pending request.
      if (typeof msg.type === "string" && msg.type.endsWith("_result") && msg.request_id) {
        const pending = pendingRequests.get(msg.request_id);
        if (pending) {
          pendingRequests.delete(msg.request_id);
          clearTimeout(pending.timer);
          pending.resolve(msg);
        }
      }
      break;
  }
}

// ---- Request/response helper (for non-tool round-trips) -------------------
const pendingRequests = new Map();
const cancelledToolCalls = new Set();
const REQUEST_TIMEOUT_MS = 10_000;

function sendRequest(type, payload = {}) {
  return new Promise((resolve, reject) => {
    if (!wsReady) {
      reject(new Error("Not connected to daemon"));
      return;
    }
    const request_id = uuid();
    // Native picking may take minutes; allow the daemon's five-minute
    // picker deadline to report its own result before the client expires.
    const timeoutMs = type === "pick_path" ? 310_000 : REQUEST_TIMEOUT_MS;
    const timer = setTimeout(() => {
      if (pendingRequests.has(request_id)) {
        pendingRequests.delete(request_id);
        reject(new Error(`Request "${type}" timed out`));
      }
    }, timeoutMs);
    pendingRequests.set(request_id, { resolve, reject, timer });
    wsSend({ type, request_id, ...payload });
  });
}

// ---------------------------------------------------------------------------
// Tool dispatcher
// ---------------------------------------------------------------------------
// Excel rejects a malformed formula with a generic, localized "invalid
// argument" message. Its code is stable, so name the likely cause.
function hasFormulaInput(args) {
  return JSON.stringify(args?.cells ?? []).includes('"formula"');
}

async function describeOfficeToolError(error, args) {
  // Office.js names the failing API (e.g. "RangeFormat.columnWidth") only in
  // debugInfo; the localized message alone doesn't say what was rejected.
  const location = error?.debugInfo?.errorLocation;
  const message = `${error?.message ?? String(error)}${location ? ` [at ${location}]` : ""}`;
  if (error?.code === "InvalidArgument" && hasFormulaInput(args)) {
    return (
      `${message} (Excel rejected the write as an invalid argument; check the formula uses Excel syntax: ` +
      "<> not !=, = not ==, AND()/OR() not &&/||, text in double quotes, balanced parentheses.)"
    );
  }
  if (!/Worksheet with ID .+ not found/i.test(message)) return message;
  try {
    const metadata = await getWorkbookMetadata();
    const valid = (metadata.sheetsMetadata ?? [])
      .map((sheet) => `${sheet.name}=${sheet.id}`)
      .join(", ");
    return `${message}. Valid worksheets in the current workbook: ${valid || "none"}. Refresh metadata and retry.`;
  } catch {
    return message;
  }
}

// Per-turn workbook context for the daemon, read with pi-for-excel's
// overview, selection and change-tracker readers. Each part is optional:
// a slow or failing read is dropped rather than delaying the turn.
const changeTracker = new ChangeTracker();
let overviewRead = null;

function readOverviewSingleFlight() {
  if (!overviewRead) {
    const read = Promise.resolve().then(buildOverview);
    overviewRead = read;
    read.then(
      () => {
        if (overviewRead === read) overviewRead = null;
      },
      () => {
        if (overviewRead === read) overviewRead = null;
      },
    );
  }
  return overviewRead;
}

function limitContextText(value, maxLength, preserveTail = false) {
  if (!value || value.length <= maxLength) return value;
  const notice = "\n[Context truncated; read the relevant range with an Excel tool.]\n";
  if (!preserveTail) return value.slice(0, maxLength - notice.length) + notice;
  const tailLength = Math.min(2000, Math.floor(maxLength / 4));
  return value.slice(0, maxLength - tailLength - notice.length) + notice + value.slice(-tailLength);
}

function settleWithin(promise, ms) {
  let timer;
  const timeout = new Promise((resolve) => {
    timer = setTimeout(() => resolve(null), ms);
  });
  return Promise.race([promise.catch(() => null), timeout]).finally(() => clearTimeout(timer));
}

async function contextSnapshot({ selectionAddress = null } = {}) {
  const [metadata, overview, selection] = await Promise.all([
    settleWithin(getWorkbookMetadata(), 1500),
    settleWithin(readOverviewSingleFlight(), 2500),
    selectionAddress ? settleWithin(readSelectionContext(selectionAddress), 1500) : null,
  ]);
  const sheetIds = (metadata?.sheetsMetadata ?? []).map((sheet) => `${sheet.name}=${sheet.id}`);
  const workbookParts = [];
  if (sheetIds.length)
    workbookParts.push(
      limitContextText(`sheetId for mcp__office__excel_* tools: ${sheetIds.join(", ")}`, 4000),
    );
  if (overview) workbookParts.push(limitContextText(overview, 8000));
  return {
    workbook: workbookParts.join("\n\n") || null,
    selection: limitContextText(selection?.text ?? null, 12000, true),
    changes: limitContextText(changeTracker.flush(), 3000),
  };
}

const WRITE_TOOLS = new Set([
  "excel_set_cell_range",
  "excel_clear_cell_range",
  "excel_copy_to",
  "excel_modify_sheet_structure",
  "excel_modify_workbook_structure",
  "excel_resize_range",
  "excel_modify_object",
  "excel_set_format",
  "excel_sort_range",
  "excel_autofilter",
  "excel_create_table",
  "excel_add_table_rows",
]);

function isMutationCall(name, args) {
  return WRITE_TOOLS.has(name) || (name === "excel_workbook_history" && args?.action === "restore");
}

function mutationSummary(name, args = {}) {
  switch (name) {
    case "excel_set_cell_range": {
      const rows = Array.isArray(args.cells) ? args.cells.length : 0;
      const columns = rows ? Math.max(...args.cells.map((row) => row.length)) : 0;
      let formulas = 0;
      let values = 0;
      for (const row of args.cells || []) {
        for (const cell of row) {
          if (cell?.formula !== undefined) formulas++;
          else if (cell?.value !== undefined) values++;
        }
      }
      const parts = [rows && columns ? `${rows}×${columns} cells` : "cells"];
      if (formulas) parts.push(`${formulas} formulas`);
      if (values) parts.push(`${values} values`);
      return parts.join(" · ");
    }
    case "excel_clear_cell_range":
      return `Clear ${args.clearType || "contents"}`;
    case "excel_copy_to":
      return `${args.sourceRange || "source"} → ${args.destinationRange || "destination"}`;
    case "excel_set_format":
      return "Apply formatting";
    case "excel_sort_range":
      return "Reorder rows";
    case "excel_resize_range":
      return "Change row height or column width";
    case "excel_modify_sheet_structure":
      return `${args.operation || "Change"} ${args.count || 1} ${args.dimension || "items"}`;
    case "excel_modify_workbook_structure":
      return `${args.operation || "Change"} worksheet`;
    case "excel_modify_object":
      return `${args.operation || "Change"} ${args.objectType || "object"}`;
    case "excel_autofilter":
      return args.clear ? "Clear filter" : "Apply filter";
    case "excel_create_table":
      return "Create table";
    case "excel_add_table_rows":
      return `Add ${args.rows?.length || 0} table rows`;
    case "excel_workbook_history":
      return "Restore previous workbook state";
    default:
      return null;
  }
}

function mutationTargets(name, args, result) {
  switch (name) {
    case "excel_set_cell_range":
      return [result?.writtenRange ?? args.copyToRange ?? args.range].filter(Boolean);
    case "excel_clear_cell_range":
      return [result?.clearedRange ?? args.range].filter(Boolean);
    case "excel_copy_to":
      return [result?.destination ?? args.destinationRange].filter(Boolean);
    case "excel_set_format":
    case "excel_sort_range":
    case "excel_create_table":
      return [result?.address ?? result?.range ?? args.address].filter(Boolean);
    case "excel_resize_range":
      return [args.range ?? "entire worksheet"];
    case "excel_modify_sheet_structure":
      return [`${args.dimension ?? "dimension"}:${args.reference ?? "pane"}:${args.count ?? 1}`];
    case "excel_modify_workbook_structure":
      return [
        result?.sheetName ?? args.newName ?? args.sheetName ?? `sheetId:${args.sheetId ?? "new"}`,
      ];
    case "excel_modify_object":
      return [
        result?.id ??
          args.id ??
          args.properties?.range ??
          args.properties?.anchor ??
          args.objectType,
      ].filter(Boolean);
    case "excel_autofilter":
      return [args.clear ? "worksheet autofilter" : (result?.address ?? args.address)].filter(
        Boolean,
      );
    case "excel_add_table_rows":
      return [args.table].filter(Boolean);
    case "excel_workbook_history":
      return result?.addresses ?? [];
    default:
      return [];
  }
}

function withMutationReceipt(name, args, result, receiptId) {
  if (!isMutationCall(name, args)) return result;
  const success = result?.success !== false;
  const readBack = name === "excel_set_cell_range" || name === "excel_copy_to";
  return {
    ...(result && typeof result === "object" ? result : { result }),
    success,
    commitStatus: result?.commitStatus ?? (success ? "committed" : "not_committed"),
    receiptId,
    operation: name,
    affectedTargets: mutationTargets(name, args, result),
    verification: {
      status: readBack ? "read_back" : "commit_acknowledged",
      semanticCheckRequired: true,
    },
  };
}

async function runOfficeTool(msg) {
  const { id, name, args } = msg;
  if (cancelledToolCalls.delete(id)) return;
  let commitRecovery = null;
  try {
    commitRecovery = WRITE_TOOLS.has(name) ? await prepareMutationRecovery(name, args, id) : null;
    let result;
    switch (name) {
      case "excel_get_selected_range":
        result = await toolExcelGetSelectedRange(args);
        break;
      case "excel_get_workbook_metadata":
        result = await getWorkbookMetadata();
        break;
      case "excel_context_snapshot":
        result = await contextSnapshot(args);
        break;
      case "excel_get_cell_ranges":
        result = await getCellRanges(args.sheetId, args.ranges, {
          includeStyles: args.includeStyles,
          cellLimit: args.cellLimit,
        });
        break;
      case "excel_get_range_as_csv":
        result = await getRangeAsCsv(args.sheetId, args.range, {
          includeHeaders: args.includeHeaders,
          maxRows: args.maxRows,
        });
        break;
      case "excel_search_data":
        result = await searchData(args.searchTerm, {
          sheetId: args.sheetId,
          range: args.range,
          offset: args.offset,
          cursor: args.cursor,
          ...args.options,
        });
        break;
      case "excel_get_all_objects":
        result = await getAllObjects({ sheetId: args.sheetId, id: args.id });
        break;
      case "excel_explain_formula":
        result = await runPiTool(explainFormulaTool, id, args);
        break;
      case "excel_trace_dependencies":
        result = await runPiTool(traceDependenciesTool, id, args);
        break;
      case "excel_set_cell_range":
        result = await setCellRange(args.sheetId, args.range, args.cells, {
          copyToRange: args.copyToRange,
          resizeWidth: args.resizeWidth,
          resizeHeight: args.resizeHeight,
          allowOverwrite: args.allow_overwrite,
        });
        break;
      case "excel_clear_cell_range":
        result = await clearCellRange(args.sheetId, args.range, args.clearType);
        break;
      case "excel_copy_to":
        result = await copyTo(
          args.sheetId,
          args.sourceRange,
          args.destinationRange,
          args.allow_overwrite,
        );
        break;
      case "excel_modify_sheet_structure":
        result = await modifySheetStructure(args.sheetId, {
          operation: args.operation,
          dimension: args.dimension,
          reference: args.reference,
          count: args.count,
          position: args.position,
        });
        break;
      case "excel_modify_workbook_structure":
        result = await modifyWorkbookStructure({
          operation: args.operation,
          sheetId: args.sheetId,
          sheetName: args.sheetName,
          newName: args.newName,
          tabColor: args.tabColor,
        });
        break;
      case "excel_resize_range":
        result = await resizeRange(args.sheetId, {
          range: args.range,
          width: args.width,
          height: args.height,
        });
        break;
      case "excel_modify_object":
        result = await modifyObject({
          operation: args.operation,
          sheetId: args.sheetId,
          objectType: args.objectType,
          id: args.id,
          properties: args.properties,
        });
        break;
      case "excel_select_range":
        result = await toolExcelSelectRange(args);
        break;
      case "excel_set_format":
        result = await toolExcelSetFormat(args);
        break;
      case "excel_sort_range":
        result = await toolExcelSortRange(args);
        break;
      case "excel_autofilter":
        result = await toolExcelAutoFilter(args);
        break;
      case "excel_create_table":
        result = await toolExcelCreateTable(args);
        break;
      case "excel_add_table_rows":
        result = await toolExcelAddTableRows(args);
        break;
      case "excel_workbook_history":
        result = await workbookHistory(args);
        break;
      default:
        throw new Error(`Unknown tool: ${name}`);
    }
    // Office.js cannot interrupt a context.sync already in progress, but a
    // daemon timeout/session stop must prevent a late result from being
    // mistaken for the current turn's result.
    if (cancelledToolCalls.delete(id)) {
      if (commitRecovery) await commitRecovery(result);
      return;
    }
    if (commitRecovery) {
      result = { ...result, recovery: await commitRecovery(result) };
    }
    // Excel.run has synced by the time a tool resolves. Attach one receipt
    // shape to every mutation and distinguish mechanical read-back from the
    // semantic check the agent still has to perform.
    result = withMutationReceipt(name, args, result, id);
    updateToolCardSuccess(id, name, args, result);
    wsSend({ type: "tool_result", id, ok: true, result });
  } catch (err) {
    const recovery = commitRecovery && isMutationCall(name, args) ? await commitRecovery() : null;
    if (cancelledToolCalls.delete(id)) return;
    console.error(`[tool ${name}] failed:`, err);
    const error = {
      message: await describeOfficeToolError(err, args),
      ...(err?.code ? { code: err.code } : {}),
      ...(err?.commitStatus
        ? { commitStatus: err.commitStatus }
        : isMutationCall(name, args)
          ? { commitStatus: "unknown" }
          : {}),
      ...(recovery ? { recovery } : {}),
    };
    updateToolCardFailure(id, error);
    wsSend({
      type: "tool_result",
      id,
      ok: false,
      error,
    });
  }
}

// ---------------------------------------------------------------------------
// Selection tracking — push context_update on changes (debounced).
// ---------------------------------------------------------------------------
let selectionDebounce = null;

async function captureSelection() {
  try {
    const r = await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      const firstCell = range.getCell(0, 0);
      range.load("address, rowCount, columnCount");
      firstCell.load("values");
      await context.sync();
      const cellCount = (range.rowCount || 0) * (range.columnCount || 0);
      const text =
        cellCount === 1
          ? String(firstCell.values?.[0]?.[0] ?? "")
          : `${range.address} (${range.rowCount}×${range.columnCount})`;
      return { text, address: range.address };
    });
    lastSelection = r;
    refreshSelectionChip();
    if (wsReady) {
      wsSend({
        type: "context_update",
        selection: attachSelection ? lastSelection : null,
      });
    }
    return r;
  } catch (e) {
    // Selection may be transient; ignore.
    return null;
  }
}

function onSelectionChanged() {
  if (selectionDebounce) clearTimeout(selectionDebounce);
  selectionDebounce = setTimeout(captureSelection, 100);
}

function renderTurnQueue() {
  if (!$turnQueue || !$turnQueueList) return;
  $turnQueue.hidden = queuedTurns.length === 0;
  $turnQueueList.innerHTML = "";
  queuedTurns.forEach((turn, index) => {
    const item = document.createElement("div");
    item.className = "turn-queue-item";
    const text = document.createElement("span");
    text.className = "turn-queue-text";
    text.textContent = turn.text;
    text.title = turn.text;
    const remove = document.createElement("button");
    remove.type = "button";
    remove.className = "turn-queue-remove";
    remove.title = "Remove queued follow-up";
    remove.setAttribute("aria-label", "Remove queued follow-up");
    remove.textContent = "\u00d7";
    remove.addEventListener("click", () => {
      queuedTurns.splice(index, 1);
      renderTurnQueue();
    });
    item.append(text, remove);
    $turnQueueList.appendChild(item);
  });
}

function clearTurnQueue() {
  queuedTurns.length = 0;
  renderTurnQueue();
}

function dispatchCapturedTurn({ text, selection }) {
  appendUserMessage(text);
  wsSend({ type: "user_message", text, selection });
  setAgentStatus("working", "Working...");
  beginTurn();
}

function drainTurnQueue() {
  if (turnInFlight || !wsReady || queuedTurns.length === 0) return false;
  const next = queuedTurns.shift();
  renderTurnQueue();
  dispatchCapturedTurn(next);
  return true;
}

async function sendUserTurn(text) {
  // Snapshot selection immediately before the message. This closes the
  // debounce window where Excel has moved but context_update still carries
  // the previous range. Detaching explicitly sends null for this turn.
  if (submitPending || !text) return false;
  if (!wsReady) {
    appendEvent("Not connected to daemon.");
    return false;
  }
  submitPending = true;
  setComposerDisabled(true);
  try {
    const selection = attachSelection ? await captureSelection() : null;
    if (!wsReady) {
      appendEvent("Disconnected before the message could be sent.");
      return false;
    }
    const turn = { text, selection };
    if (turnInFlight || queuedTurns.length > 0) {
      queuedTurns.push(turn);
      renderTurnQueue();
      drainTurnQueue();
    } else {
      dispatchCapturedTurn(turn);
    }
    // The detach choice applies to one turn; the next turn starts attached.
    attachSelection = true;
    refreshSelectionChip();
    return true;
  } finally {
    submitPending = false;
    setComposerDisabled(false);
    $send.textContent = turnInFlight ? "Queue" : "Send";
  }
}

// ---------------------------------------------------------------------------
// Composer
// ---------------------------------------------------------------------------
$composer.addEventListener("submit", async (e) => {
  e.preventDefault();
  const text = $input.value.trim();
  if (await sendUserTurn(text)) $input.value = "";
});

$input.addEventListener("keydown", (e) => {
  // isComposing / keyCode 229: Enter that confirms an IME candidate (Chinese,
  // Japanese…) must not send the half-typed message.
  if (e.key === "Enter" && !e.shiftKey && !e.isComposing && e.keyCode !== 229) {
    e.preventDefault();
    $composer.dispatchEvent(new Event("submit"));
  }
});

// ---------------------------------------------------------------------------
// Theme — match Excel's theme (light/dark) via Office.context.officeTheme.
// If Office doesn't expose a theme (older build), the CSS
// `prefers-color-scheme: dark` media query already handles the OS-level
// preference, so this is purely an override for when Excel's theme differs
// from the OS.
// ---------------------------------------------------------------------------
function applyOfficeTheme() {
  try {
    const t = Office.context && Office.context.officeTheme;
    const bg = t && t.bodyBackgroundColor;
    if (!bg) return;
    const m = /^#?([0-9a-f]{6})$/i.exec(bg);
    if (!m) return;
    const hex = m[1];
    const r = parseInt(hex.slice(0, 2), 16);
    const g = parseInt(hex.slice(2, 4), 16);
    const b = parseInt(hex.slice(4, 6), 16);
    const luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
    document.documentElement.dataset.theme = luminance < 0.5 ? "dark" : "light";
  } catch {
    /* fall through; CSS media query handles OS-level preference. */
  }
}

// ---------------------------------------------------------------------------
// Boot
// ---------------------------------------------------------------------------
Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    setConnectionStatus("err", `Unsupported host: ${info.host}`);
    return;
  }

  applyOfficeTheme();

  try {
    activeDocUrl = Office.context.document.url || null;
  } catch {
    /* ignore */
  }
  refreshMismatchIndicator();

  // Wire selection-change event.
  Excel.run(async (context) => {
    context.workbook.onSelectionChanged.add(onSelectionChanged);
    await context.sync();
  }).catch((err) => console.warn("Could not attach Excel selection handler:", err));

  // Capture once on boot.
  captureSelection();
  changeTracker.start();

  // Initialize presets UI.
  initPresets();

  // Show the welcome card on first launch (per browser-storage, so it
  // re-shows in a fresh profile or if the user clears storage).
  maybeShowOnboarding();

  // Fetch the bridge token before opening the WS — the bridge will close any
  // connection that doesn't present it.
  fetchBridgeToken().then(() => wsConnect());
});

// ---------------------------------------------------------------------------
// First-run onboarding card
// ---------------------------------------------------------------------------
const ONBOARDING_KEY = "claude-code-office-onboarding-seen-v1";
function maybeShowOnboarding() {
  try {
    if (localStorage.getItem(ONBOARDING_KEY)) return;
  } catch {
    /* ignore */
  }
  const hostLabel = "Excel";
  const card = document.createElement("div");
  card.className = "onboarding-card";
  card.innerHTML = `
    <div class="onboarding-head">
      <strong>Welcome to Excel Assistant</strong>
      <button type="button" class="onboarding-dismiss" title="Dismiss">×</button>
    </div>
    <div class="onboarding-body">
      <p>The agent reads and edits this ${hostLabel} document, plus any folders or files you add as context.</p>
      <ol>
        <li><strong>Pick a workspace folder</strong> in the <a href="#" data-onboarding-jump="setup">Setup tab</a> — that's the folder Claude treats as its working directory.</li>
        <li><strong>Try a preset</strong> — the chips above the chat input are one-click prompts. "Summarize this document" is a good first try.</li>
        <li><strong>Add context files</strong> in Setup if you want Claude to consider background material (notes, prior drafts, references).</li>
      </ol>
    </div>
  `;
  const messagesEl = document.getElementById("messages");
  if (messagesEl) messagesEl.insertBefore(card, messagesEl.firstChild);
  else document.body.insertBefore(card, document.body.firstChild);
  card.querySelector(".onboarding-dismiss").addEventListener("click", () => {
    try {
      localStorage.setItem(ONBOARDING_KEY, "1");
    } catch {
      /* ignore */
    }
    card.remove();
  });
  card.querySelector("[data-onboarding-jump='setup']").addEventListener("click", (e) => {
    e.preventDefault();
    setActiveTab("setup");
  });
}

// ===========================================================================
// Tabs
// ===========================================================================
function setActiveTab(tabName) {
  document.body.dataset.activeTab = tabName;
  document.querySelectorAll(".tab-content").forEach((el) => {
    el.hidden = el.dataset.tab !== tabName;
  });
  document.querySelectorAll(".tab").forEach((btn) => {
    btn.setAttribute("aria-current", btn.dataset.tab === tabName ? "page" : "false");
  });
  if (tabName === "setup") {
    loadContext();
    loadWorkspaceSection();
  } else if (tabName === "backups") {
    loadRecoveryHistory();
  }
}

document.querySelectorAll(".tab").forEach((btn) => {
  btn.addEventListener("click", () => setActiveTab(btn.dataset.tab));
});

// ===========================================================================
// Presets — saved prompts that the user can pin to quick-chips or browse
// in the Library tab.
// ===========================================================================
// Workbook backups ----------------------------------------------------------
const $backupList = document.getElementById("backup-list");
const $backupStatus = document.getElementById("backup-status");
const $backupSearch = document.getElementById("backup-search");
const $backupRefresh = document.getElementById("backup-refresh");
const $backupClear = document.getElementById("backup-clear");
let recoverySnapshots = [];
let recoveryBusy = false;

const RECOVERY_OPERATION_LABELS = {
  excel_set_cell_range: "Edit cells",
  excel_clear_cell_range: "Clear cells",
  excel_copy_to: "Copy cells",
  excel_set_format: "Format cells",
  excel_sort_range: "Sort range",
  excel_resize_range: "Resize rows or columns",
  excel_modify_sheet_structure: "Insert or delete rows or columns",
  excel_modify_workbook_structure: "Change worksheet structure",
  restore: "Restore backup",
};

function recoveryOperationLabel(operation) {
  if (RECOVERY_OPERATION_LABELS[operation]) return RECOVERY_OPERATION_LABELS[operation];
  const plain = String(operation || "Workbook change")
    .replace(/^excel_/, "")
    .replaceAll("_", " ");
  return plain.charAt(0).toUpperCase() + plain.slice(1);
}

function showRecoveryStatus(message, { error = false } = {}) {
  if (!$backupStatus) return;
  $backupStatus.hidden = !message;
  $backupStatus.classList.toggle("error", error);
  $backupStatus.textContent = message || "";
}

function setRecoveryBusy(busy) {
  recoveryBusy = busy;
  if ($backupRefresh) $backupRefresh.disabled = busy;
  if ($backupClear) $backupClear.disabled = busy || recoverySnapshots.length === 0;
  $backupList?.querySelectorAll("button").forEach((button) => {
    button.disabled = busy;
  });
}

function renderRecoveryHistory() {
  if (!$backupList) return;
  $backupList.innerHTML = "";
  const query = ($backupSearch?.value || "").trim().toLowerCase();
  const visible = recoverySnapshots.filter((snapshot) => {
    if (!query) return true;
    return [
      recoveryOperationLabel(snapshot.operation),
      ...(snapshot.addresses || []),
      ...(snapshot.kinds || []),
    ]
      .join(" ")
      .toLowerCase()
      .includes(query);
  });

  if (visible.length === 0) {
    const empty = document.createElement("div");
    empty.className = "backup-empty";
    empty.textContent = recoverySnapshots.length
      ? "No backups match this search."
      : "No backups yet. A backup appears here before the assistant changes the workbook.";
    $backupList.appendChild(empty);
    if ($backupClear) $backupClear.disabled = recoverySnapshots.length === 0 || recoveryBusy;
    return;
  }

  for (const snapshot of visible) {
    const item = document.createElement("div");
    item.className = "backup-item";

    const titleRow = document.createElement("div");
    titleRow.className = "backup-title-row";
    const title = document.createElement("div");
    title.className = "backup-title";
    title.textContent = recoveryOperationLabel(snapshot.operation);
    const count = document.createElement("span");
    count.className = "backup-count";
    count.textContent = `${Number(snapshot.changedCount || 0).toLocaleString()} cells`;
    titleRow.append(title, count);

    const addresses = document.createElement("div");
    addresses.className = "backup-addresses";
    addresses.textContent = (snapshot.addresses || []).join(", ") || "Workbook structure";

    const meta = document.createElement("div");
    meta.className = "backup-meta";
    const created = new Date(snapshot.createdAt);
    meta.textContent = Number.isNaN(created.getTime()) ? "" : created.toLocaleString();

    const actions = document.createElement("div");
    actions.className = "backup-actions";
    const restore = document.createElement("button");
    restore.type = "button";
    restore.textContent = "Restore";
    restore.disabled = recoveryBusy;
    restore.addEventListener("click", async () => {
      if (turnInFlight || submitPending) {
        showRecoveryStatus("Stop the current assistant turn before restoring a backup.", {
          error: true,
        });
        return;
      }
      await runRecoveryAction("restore", snapshot.id);
    });

    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "Delete";
    remove.disabled = recoveryBusy;
    let confirmDelete = false;
    remove.addEventListener("click", async () => {
      if (!confirmDelete) {
        confirmDelete = true;
        remove.textContent = "Delete?";
        remove.classList.add("backup-delete-confirm");
        return;
      }
      await runRecoveryAction("delete", snapshot.id);
    });

    actions.append(restore, remove);
    item.append(titleRow, addresses, meta, actions);
    $backupList.appendChild(item);
  }
  if ($backupClear) $backupClear.disabled = recoveryBusy;
}

async function loadRecoveryHistory(message = "") {
  if (!$backupList || recoveryBusy) return;
  setRecoveryBusy(true);
  showRecoveryStatus(message || "Loading backups...");
  try {
    const result = await workbookHistory({ action: "list", limit: 120 });
    recoverySnapshots = result.snapshots || [];
    showRecoveryStatus(message);
  } catch (error) {
    showRecoveryStatus(error?.message || String(error), { error: true });
  } finally {
    setRecoveryBusy(false);
    renderRecoveryHistory();
  }
}

async function runRecoveryAction(action, snapshotId = null) {
  if (recoveryBusy) return;
  setRecoveryBusy(true);
  showRecoveryStatus(action === "restore" ? "Restoring backup..." : "Deleting backup...");
  try {
    const result = await workbookHistory({ action, snapshot_id: snapshotId });
    const listed = await workbookHistory({ action: "list", limit: 120 });
    recoverySnapshots = listed.snapshots || [];
    if (action === "restore") {
      const targets = result.addresses?.join(", ") || "the workbook";
      showRecoveryStatus(`Restored ${targets}. A reverse backup was created.`);
    } else {
      showRecoveryStatus("Backup deleted.");
    }
  } catch (error) {
    showRecoveryStatus(error?.message || String(error), { error: true });
  } finally {
    setRecoveryBusy(false);
    renderRecoveryHistory();
  }
}

$backupRefresh?.addEventListener("click", () => loadRecoveryHistory());
$backupSearch?.addEventListener("input", renderRecoveryHistory);
$backupClear?.addEventListener("click", async () => {
  if ($backupClear.dataset.confirm !== "true") {
    $backupClear.dataset.confirm = "true";
    $backupClear.textContent = "Clear all?";
    $backupClear.classList.add("backup-delete-confirm");
    return;
  }
  if (recoveryBusy) return;
  setRecoveryBusy(true);
  showRecoveryStatus("Clearing backups...");
  try {
    const result = await workbookHistory({ action: "clear" });
    recoverySnapshots = [];
    showRecoveryStatus(`Cleared ${result.removed || 0} backups.`);
  } catch (error) {
    showRecoveryStatus(error?.message || String(error), { error: true });
  } finally {
    $backupClear.dataset.confirm = "false";
    $backupClear.textContent = "Clear all";
    $backupClear.classList.remove("backup-delete-confirm");
    setRecoveryBusy(false);
    renderRecoveryHistory();
  }
});

const PRESETS_KEY = "claude-code-office-presets-v1:" + HOST;

function defaultPresets() {
  return defaultExcelPresets();
}

function defaultExcelPresets() {
  return [
    {
      id: uuid(),
      title: "Summarize this sheet",
      category: "Summarize",
      prompt:
        "List the worksheets, then read the active sheet's used range and give me a tight summary — what the data is, columns, row count, anything notable. Don't change anything.",
      pinned: true,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Explain the selected range",
      category: "Summarize",
      prompt:
        "Read my current selection and explain what it contains — the columns, the values, and any pattern or total worth noting. Don't change anything.",
      pinned: true,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Add a totals row",
      category: "Edit",
      prompt:
        "Add a labelled Total row beneath the data on the active sheet, using SUM formulas (not pre-computed numbers) for each numeric column. Read the range first; don't overwrite existing formulas.",
      pinned: false,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Check the data for problems",
      category: "Review",
      prompt:
        "Scan the active sheet for data problems — blank cells in a filled column, inconsistent formatting/casing, likely typos, duplicates. List what you find in chat with cell addresses; don't change anything yet.",
      pinned: false,
      auto_send: true,
    },
    {
      id: uuid(),
      title: "Find a value",
      category: "Edit",
      prompt: "Find every cell containing: ",
      pinned: false,
      auto_send: false,
    },
    {
      id: uuid(),
      title: "Answer using my context files",
      category: "Research",
      prompt: "Use the context files I've added to this workspace to answer: ",
      pinned: false,
      auto_send: false,
    },
  ];
}

function uuid() {
  if (crypto?.randomUUID) return crypto.randomUUID();
  return "p_" + Math.random().toString(36).slice(2, 10);
}

let presets = [];

function loadPresets() {
  try {
    const raw = localStorage.getItem(PRESETS_KEY);
    if (raw) return JSON.parse(raw);
  } catch {
    /* ignore */
  }
  return null;
}

function savePresets() {
  try {
    localStorage.setItem(PRESETS_KEY, JSON.stringify(presets));
  } catch {
    /* ignore */
  }
}

function initPresets() {
  const existing = loadPresets();
  if (existing === null) {
    presets = defaultPresets();
    savePresets();
  } else {
    presets = existing;
  }
  renderLibrary();
  renderQuickChips();
}

// ---- Quick chips (pinned presets) -----------------------------------------
const $quickChips = document.getElementById("quick-chips");

function renderQuickChips() {
  const pinned = presets.filter((p) => p.pinned);
  if (pinned.length === 0) {
    $quickChips.hidden = true;
    $quickChips.innerHTML = "";
    return;
  }
  $quickChips.hidden = false;
  $quickChips.innerHTML = "";
  for (const p of pinned) {
    const chip = document.createElement("button");
    chip.type = "button";
    chip.className = "quick-chip";
    chip.textContent = p.title;
    chip.title = p.prompt.length > 200 ? p.prompt.slice(0, 197) + "…" : p.prompt;
    chip.addEventListener("click", () => usePreset(p));
    $quickChips.appendChild(chip);
  }
}

// ---- Library tab list -----------------------------------------------------
const $libraryList = document.getElementById("library-list");

function renderLibrary() {
  $libraryList.innerHTML = "";
  if (presets.length === 0) {
    const empty = document.createElement("div");
    empty.className = "library-empty";
    empty.textContent = 'No presets yet. Click "+ New preset" to add one.';
    $libraryList.appendChild(empty);
    return;
  }

  // Group by category. Uncategorized go under "Other".
  const groups = new Map();
  for (const p of presets) {
    const k = p.category && p.category.trim() ? p.category.trim() : "Other";
    if (!groups.has(k)) groups.set(k, []);
    groups.get(k).push(p);
  }

  for (const [category, items] of groups) {
    const heading = document.createElement("div");
    heading.className = "library-category";
    heading.textContent = category;
    $libraryList.appendChild(heading);
    for (const p of items) {
      $libraryList.appendChild(renderPresetRow(p));
    }
  }
}

function renderPresetRow(p) {
  const row = document.createElement("div");
  row.className = "preset-row";
  row.title = p.prompt.length > 300 ? p.prompt.slice(0, 297) + "…" : p.prompt;

  // Click row → use preset
  row.addEventListener("click", (e) => {
    if (e.target.closest(".preset-actions") || e.target.closest(".preset-pin")) return;
    usePreset(p);
  });

  const title = document.createElement("div");
  title.className = "preset-title";
  title.textContent = p.title;
  row.appendChild(title);

  const meta = document.createElement("div");
  meta.className = "preset-meta";
  if (p.auto_send) meta.textContent = "auto-send";
  row.appendChild(meta);

  const pin = document.createElement("button");
  pin.type = "button";
  pin.className = "preset-pin" + (p.pinned ? " pinned" : "");
  pin.textContent = p.pinned ? "📌" : "📍";
  pin.title = p.pinned ? "Pinned (click to unpin)" : "Pin to quick chips";
  pin.addEventListener("click", (e) => {
    e.stopPropagation();
    p.pinned = !p.pinned;
    savePresets();
    renderLibrary();
    renderQuickChips();
  });
  row.appendChild(pin);

  const actions = document.createElement("div");
  actions.className = "preset-actions";
  const editBtn = document.createElement("button");
  editBtn.className = "icon-btn";
  editBtn.type = "button";
  editBtn.title = "Edit";
  editBtn.textContent = "✎";
  editBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    openPresetModal(p);
  });
  actions.appendChild(editBtn);

  const delBtn = document.createElement("button");
  delBtn.className = "icon-btn";
  delBtn.type = "button";
  delBtn.title = "Delete";
  delBtn.textContent = "✕";
  delBtn.addEventListener("click", (e) => {
    e.stopPropagation();
    // No confirm dialog (Office.js taskpanes block it on some builds); the
    // user can re-add a deleted preset if needed.
    presets = presets.filter((x) => x.id !== p.id);
    savePresets();
    renderLibrary();
    renderQuickChips();
  });
  actions.appendChild(delBtn);

  row.appendChild(actions);
  return row;
}

// ---- Use preset (click handler) -------------------------------------------
async function usePreset(p) {
  if (p.auto_send) {
    await sendUserTurn(p.prompt);
  } else {
    $input.value = p.prompt;
    $input.focus();
    // Cursor at end so user can extend the prompt
    $input.setSelectionRange($input.value.length, $input.value.length);
  }
  setActiveTab("chat");
}

// ---- Preset editor modal --------------------------------------------------
const $presetModal = document.getElementById("preset-modal");
const $presetModalTitle = document.getElementById("preset-modal-title");
const $presetTitle = document.getElementById("preset-title");
const $presetPrompt = document.getElementById("preset-prompt");
const $presetCategory = document.getElementById("preset-category");
const $presetAutoSend = document.getElementById("preset-auto-send");
const $presetPinned = document.getElementById("preset-pinned");
const $presetSave = document.getElementById("preset-save");
const $presetCancel = document.getElementById("preset-cancel");
const $presetModalClose = document.getElementById("preset-modal-close");

let editingPresetId = null;

function openPresetModal(p) {
  if (p) {
    editingPresetId = p.id;
    $presetModalTitle.textContent = "Edit preset";
    $presetTitle.value = p.title;
    $presetPrompt.value = p.prompt;
    $presetCategory.value = p.category || "";
    $presetAutoSend.checked = !!p.auto_send;
    $presetPinned.checked = !!p.pinned;
  } else {
    editingPresetId = null;
    $presetModalTitle.textContent = "New preset";
    $presetTitle.value = "";
    $presetPrompt.value = "";
    $presetCategory.value = "";
    $presetAutoSend.checked = false;
    $presetPinned.checked = false;
  }
  $presetModal.hidden = false;
  setTimeout(() => $presetTitle.focus(), 0);
}

function closePresetModal() {
  $presetModal.hidden = true;
  editingPresetId = null;
}

$presetCancel.addEventListener("click", closePresetModal);
$presetModalClose.addEventListener("click", closePresetModal);
$presetModal.addEventListener("click", (e) => {
  if (e.target === $presetModal) closePresetModal();
});

$presetSave.addEventListener("click", () => {
  const title = $presetTitle.value.trim();
  const prompt = $presetPrompt.value;
  if (!title) {
    $presetTitle.focus();
    return;
  }
  if (!prompt.trim()) {
    $presetPrompt.focus();
    return;
  }

  const data = {
    title,
    prompt,
    category: $presetCategory.value.trim(),
    auto_send: $presetAutoSend.checked,
    pinned: $presetPinned.checked,
  };

  if (editingPresetId) {
    const idx = presets.findIndex((p) => p.id === editingPresetId);
    if (idx !== -1) presets[idx] = { ...presets[idx], ...data };
  } else {
    presets.push({ id: uuid(), ...data });
  }

  savePresets();
  renderLibrary();
  renderQuickChips();
  closePresetModal();
});

document.getElementById("add-preset").addEventListener("click", () => openPresetModal(null));

// ===========================================================================
// Setup tab — Context files (folders or individual files saved to the
// workspace's CLAUDE.md, loaded by Claude Code each session).
// ===========================================================================
let contextCache = null;
let contextCacheCwd = null;
let contextLoading = null;
let contextLoadToken = Symbol("context-load");
const $contextList = document.getElementById("context-list");

function resetContextForWorkspace(cwd) {
  contextLoadToken = Symbol("context-load");
  contextCache = null;
  contextCacheCwd = cwd ?? null;
  contextLoading = null;
}

async function removeContextEntryAt(idx) {
  if (!contextCache) return;
  contextCache = contextCache.filter((_, i) => i !== idx);
  await saveContext();
}

async function loadContext(force = false) {
  const requestCwd = currentWorkspaceCwd;
  if (contextCache && contextCacheCwd === requestCwd && !force) {
    renderContext();
    return;
  }
  if (contextLoading?.cwd === requestCwd && !force) return contextLoading.promise;
  const token = Symbol("context-load");
  contextLoadToken = token;
  const promise = (async () => {
    try {
      $contextList.innerHTML = '<div class="references-loading">Loading…</div>';
      const r = await sendRequest("get_context");
      if (
        contextLoadToken !== token ||
        currentWorkspaceCwd !== requestCwd ||
        (r.cwd && currentWorkspaceCwd && r.cwd !== currentWorkspaceCwd)
      ) {
        return;
      }
      contextCache = Array.isArray(r.entries) ? r.entries : [];
      contextCacheCwd = r.cwd ?? requestCwd;
      renderContext();
    } catch (e) {
      if (contextLoadToken === token && currentWorkspaceCwd === requestCwd) {
        showListMessage($contextList, "references-empty", `Could not load: ${e.message}`);
      }
    } finally {
      if (contextLoadToken === token) contextLoading = null;
    }
  })();
  contextLoading = { cwd: requestCwd, promise };
  return promise;
}

function renderContext() {
  $contextList.innerHTML = "";
  if (!contextCache || contextCache.length === 0) {
    const empty = document.createElement("div");
    empty.className = "references-empty";
    empty.textContent = "No context files yet — add a folder or file to give Claude background.";
    $contextList.appendChild(empty);
    return;
  }
  contextCache.forEach((e, i) => {
    const row = document.createElement("div");
    row.className = "reference-row";
    row.title = e.path;

    const info = document.createElement("div");
    info.className = "reference-info";
    const pathEl = document.createElement("div");
    pathEl.className = "reference-path";
    const tag = document.createElement("span");
    tag.className = "kind-tag";
    tag.textContent = e.kind || "?";
    pathEl.appendChild(tag);
    pathEl.appendChild(document.createTextNode(e.path));
    info.appendChild(pathEl);
    if (e.description) {
      const desc = document.createElement("div");
      desc.className = "reference-description";
      desc.textContent = e.description;
      info.appendChild(desc);
    }
    row.appendChild(info);

    const remove = document.createElement("button");
    remove.className = "reference-remove";
    remove.type = "button";
    remove.title = "Remove";
    remove.textContent = "✕";
    remove.addEventListener("click", () => removeContextEntryAt(i));
    row.appendChild(remove);

    $contextList.appendChild(row);
  });
}

// In-pane error notice for the Context-files section. window.alert() works
// on Mac Office but is flaky/blocked in web Office; show a dismissible
// inline message in the Setup tab instead. textContent — never innerHTML —
// so a daemon-supplied path/error string can't inject markup.
const $contextError = document.getElementById("context-error");
function showContextError(message) {
  if (!$contextError) {
    console.warn("[context]", message);
    return;
  }
  $contextError.textContent = message;
  $contextError.hidden = false;
}
function clearContextError() {
  if ($contextError) {
    $contextError.textContent = "";
    $contextError.hidden = true;
  }
}

async function saveContext() {
  if (!contextCache) return false;
  const targetCwd = contextCacheCwd ?? currentWorkspaceCwd;
  if (!targetCwd || targetCwd !== currentWorkspaceCwd) return false;
  const entries = contextCache.map((entry) => ({ ...entry }));
  clearContextError();
  try {
    const r = await sendRequest("set_context", { entries, expected_cwd: targetCwd });
    if (!r.ok) throw new Error(r.error || "Could not save context files");
    if (currentWorkspaceCwd !== targetCwd || contextCacheCwd !== targetCwd) return false;
    if (r.errors && r.errors.length > 0) {
      const lines = r.errors.map((e) => `${e.path} — ${e.error}`).join("; ");
      showContextError(`Some entries could not be saved: ${lines}`);
    }
    contextCache = Array.isArray(r.saved) ? r.saved : contextCache;
    renderContext();
    return true;
  } catch (e) {
    if (currentWorkspaceCwd === targetCwd) showContextError(`Could not save: ${e.message}`);
    return false;
  }
}

// ---- Add-folder modal (shared between guidelines + samples) ----------------
const $addFolderModal = document.getElementById("add-folder-modal");
const $addFolderModalTitle = document.getElementById("add-folder-modal-title");
const $addFolderPath = document.getElementById("add-folder-path");
const $addFolderDescription = document.getElementById("add-folder-description");
const $addFolderError = document.getElementById("add-folder-error");
const $addFolderSave = document.getElementById("add-folder-save");
const $addFolderCancel = document.getElementById("add-folder-cancel");
const $addFolderModalClose = document.getElementById("add-folder-modal-close");
const $addFolderBrowseFile = document.getElementById("add-folder-browse-file");
const $addFolderBrowseFolder = document.getElementById("add-folder-browse-folder");
let addFolderTargetCwd = null;

function openAddFolderModal(prefillPath = "", kind = null) {
  addFolderTargetCwd = currentWorkspaceCwd;
  $addFolderModalTitle.textContent =
    kind === "file" ? "Add file" : kind === "folder" ? "Add folder" : "Add folder or file";
  $addFolderPath.value = prefillPath;
  $addFolderDescription.value = "";
  $addFolderError.hidden = true;
  $addFolderModal.hidden = false;
  // If the path is already chosen (the common flow — user picked first),
  // jump straight to the description field.
  setTimeout(() => (prefillPath ? $addFolderDescription : $addFolderPath).focus(), 0);
}
function closeAddFolderModal() {
  $addFolderModal.hidden = true;
}

// Two explicit entry points: pick first (single-mode dialog, reliable on
// every OS), then the modal just collects an optional description.
async function addContextEntry(includeFiles) {
  let picked;
  try {
    picked = await pickPathNative({
      start_path: currentWorkspaceCwd || null,
      include_files: includeFiles,
    });
  } catch (e) {
    console.error("[picker]", e);
    return;
  }
  if (!picked) return; // canceled — don't open an empty modal
  openAddFolderModal(picked.path, includeFiles ? "file" : "folder");
}
document
  .querySelectorAll(".add-context-folder")
  .forEach((b) => b.addEventListener("click", () => addContextEntry(false)));
document
  .querySelectorAll(".add-context-file")
  .forEach((b) => b.addEventListener("click", () => addContextEntry(true)));
$addFolderCancel.addEventListener("click", closeAddFolderModal);
$addFolderModalClose.addEventListener("click", closeAddFolderModal);
$addFolderModal.addEventListener("click", (e) => {
  if (e.target === $addFolderModal) closeAddFolderModal();
});

$addFolderSave.addEventListener("click", async () => {
  const path = $addFolderPath.value.trim();
  const description = $addFolderDescription.value.trim();
  const targetCwd = addFolderTargetCwd;
  $addFolderError.hidden = true;
  if (!path) {
    $addFolderError.textContent = "Path is required.";
    $addFolderError.hidden = false;
    return;
  }
  if (!targetCwd || currentWorkspaceCwd !== targetCwd) {
    $addFolderError.textContent =
      "The workspace changed. Close this dialog and add the file again.";
    $addFolderError.hidden = false;
    return;
  }
  if (!contextCache || contextCacheCwd !== currentWorkspaceCwd) await loadContext();
  if (currentWorkspaceCwd !== targetCwd || contextCacheCwd !== targetCwd) {
    $addFolderError.textContent =
      "The workspace changed. Close this dialog and add the file again.";
    $addFolderError.hidden = false;
    return;
  }
  if (!contextCache) contextCache = [];
  contextCache = [...contextCache, { path, description }];
  const ok = await saveContext();
  if (ok) closeAddFolderModal();
});

// ---- Native folder/file picker ---------------------------------------------
// Forwards the pick request through the daemon to the Electron main process,
// which shows a real macOS NSOpenPanel. This is the same panel every native
// Mac app uses, so it can navigate to Google Drive, iCloud, "Shared with me",
// recent items, sidebar shortcuts — none of which a synthetic in-page browser
// can reach. Resolves to `{ path, kind }` or `null` if the user cancelled.
async function pickPathNative({ start_path = null, include_files = false, title = null } = {}) {
  const r = await sendRequest("pick_path", {
    default_path: start_path,
    include_files,
    title,
  });
  if (!r.ok) throw new Error(r.error || "Picker failed");
  if (r.canceled) return null;
  return { path: r.path, kind: r.kind };
}

// Two single-mode pickers. A combined file+folder native dialog can't
// exist on Windows (it degrades to folder-only, hiding files), so the
// user explicitly chooses which kind to browse for.
async function browseInto(includeFiles) {
  const startPath = $addFolderPath.value.trim() || currentWorkspaceCwd || null;
  try {
    const picked = await pickPathNative({ start_path: startPath, include_files: includeFiles });
    if (!picked) return;
    $addFolderPath.value = picked.path;
    $addFolderDescription.focus();
  } catch (e) {
    console.error("[picker]", e);
  }
}
$addFolderBrowseFile.addEventListener("click", () => browseInto(true));
$addFolderBrowseFolder.addEventListener("click", () => browseInto(false));

$addFolderPath.addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    e.preventDefault();
    $addFolderSave.click();
  }
});
$addFolderDescription.addEventListener("keydown", (e) => {
  if (e.key === "Enter") {
    e.preventDefault();
    $addFolderSave.click();
  }
});

// ===========================================================================
// Workspace selection — managed in the Setup tab's Workspace section.
// The topbar chip is read-only status (current workspace name + mismatch warning);
// clicking it navigates to the Setup tab where the actual switching UI lives.
// ===========================================================================
let currentWorkspaceCwd = null;

// Replace a list container with a single-line empty-state message. Uses
// textContent so an attacker-controlled error string can't inject markup.
function showListMessage(container, klass, message) {
  container.innerHTML = "";
  const div = document.createElement("div");
  div.className = klass;
  div.textContent = message;
  container.appendChild(div);
}

const $workspaceChip = document.getElementById("workspace-chip");
const $workspaceFolder = document.getElementById("workspace-folder");
const $workspaceCwd = document.getElementById("workspace-cwd");
const $addWorkspace = document.getElementById("add-workspace");
const $workspaceError = document.getElementById("workspace-error");
const $workspaceWarning = document.getElementById("workspace-warning");

// The workspace is simply the folder the open document lives in. We follow
// it automatically. An explicit "Change workspace" pick overrides that and
// sticks until the active document moves to a *different* folder — tracked
// via lastFollowedDocDir so a re-render/reconnect doesn't yank the user's
// deliberate choice back.
let lastFollowedDocDir = null;

function setWorkspaceDisplay(cwd) {
  if (currentWorkspaceCwd !== cwd) resetContextForWorkspace(cwd);
  currentWorkspaceCwd = cwd;
  const name = cwd ? cwd.split(/[\\/]/).filter(Boolean).pop() : "(no workspace)";
  $workspaceFolder.textContent = name;
  $workspaceFolder.title = cwd || "";
  if ($workspaceCwd) {
    $workspaceCwd.textContent = cwd || "(no workspace)";
    $workspaceCwd.title = cwd || "";
  }
  refreshMismatchIndicator();
}

// Decide whether the workspace chip should show a warning. True when an active
// doc exists and its filesystem location is NOT inside the current workspace
// folder — the agent's filesystem tools won't see the doc's siblings.
function refreshMismatchIndicator() {
  let mismatch = false;
  if (activeDocUrl && currentWorkspaceCwd) {
    const docDir = docDirFromActiveUrl(activeDocUrl);
    // Mismatch if docDir isn't underneath the workspace cwd. Cloud-hosted
    // docs (docDir === "") fall through as not-a-mismatch — there's no
    // filesystem location to compare, so the chip stays neutral.
    mismatch = !!docDir && !isInOrUnder(docDir, currentWorkspaceCwd);
  }
  const baseTitle =
    "The agent reads source files (CLAUDE.md, notes, references) from the workspace folder — click to switch workspaces.";
  if (mismatch) {
    $workspaceChip.classList.add("mismatch");
    $workspaceWarning.hidden = false;
    $workspaceChip.title =
      baseTitle + " (⚠ The current workspace doesn't match the doc you're editing.)";
  } else {
    $workspaceChip.classList.remove("mismatch");
    $workspaceWarning.hidden = true;
    $workspaceChip.title = baseTitle;
  }
}

async function refreshWorkspaceFromDaemon() {
  try {
    const r = await sendRequest("get_cwd_state");
    if (r.current_cwd) setWorkspaceDisplay(r.current_cwd);
    await autoFollowDocWorkspace();
  } catch {
    /* ignore on initial boot */
  }
}

async function loadWorkspaceSection() {
  $workspaceError.hidden = true;
  await autoFollowDocWorkspace();
  try {
    const r = await sendRequest("get_cwd_state");
    if (r.current_cwd) setWorkspaceDisplay(r.current_cwd);
  } catch (e) {
    $workspaceError.textContent = `Could not load workspace: ${e.message}`;
    $workspaceError.hidden = false;
  }
}

// The workspace follows the open document's folder automatically. Workspace
// detection is deterministic now (the doc's own folder), so there's nothing
// to confirm — no banner, no setting. We switch when the doc's folder isn't
// the current workspace, UNLESS the user made an explicit pick for this same
// doc-folder (lastFollowedDocDir), in which case their choice stands until
// they open a document in a different folder.
async function autoFollowDocWorkspace() {
  if (!activeDocUrl) return;
  const docDir = docDirFromActiveUrl(activeDocUrl);
  if (!docDir) return; // cloud doc (no filesystem path) — leave workspace as-is
  if (docDir === lastFollowedDocDir) return; // already handled (incl. explicit override)
  if (currentWorkspaceCwd && isInOrUnder(docDir, currentWorkspaceCwd)) {
    lastFollowedDocDir = docDir; // doc already inside the workspace — fine
    return;
  }
  lastFollowedDocDir = docDir;
  await doSwitch(null, { autodetectFromDoc: true });
}

async function doSwitch(cwd, { autodetectFromDoc = false } = {}) {
  $workspaceError.hidden = true;
  try {
    const payload = autodetectFromDoc ? { autodetect_from_doc: activeDocUrl } : { cwd };
    const r = await sendRequest("set_cwd", payload);
    if (!r.ok) throw new Error(r.error || "switch failed");
    // An explicit pick (not the auto-follow) is a deliberate override: pin
    // it to the current doc's folder so autoFollowDocWorkspace won't yank
    // it back until the user opens a document in a different folder.
    if (!autodetectFromDoc) lastFollowedDocDir = docDirFromActiveUrl(activeDocUrl) || null;
    // The daemon emits cwd_changed which updates the chip via assistant_event.
    // Clear chat — visually distinguishing the new session from the old.
    $messages.innerHTML = "";
    assistantTurnElem = null;
    // Refresh the workspace section so the displayed cwd updates.
    loadWorkspaceSection();
  } catch (e) {
    $workspaceError.textContent = e.message;
    $workspaceError.hidden = false;
  }
}

// Clicking the topbar chip jumps to the Setup tab where the workspace UI lives.
$workspaceChip.addEventListener("click", () => setActiveTab("setup"));

// "Change workspace" — opens the native folder picker; on pick, switch to
// that folder (a deliberate override of the auto-followed doc folder).
$addWorkspace.addEventListener("click", async () => {
  try {
    const picked = await pickPathNative({ title: "Choose a workspace folder" });
    if (picked) doSwitch(picked.path);
  } catch (e) {
    console.error("[picker]", e);
  }
});
