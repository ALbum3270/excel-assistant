import { query } from "@anthropic-ai/claude-agent-sdk";
import { readFile, writeFile, mkdir, chmod } from "node:fs/promises";
import { randomBytes, createHash } from "node:crypto";
import { createServer } from "node:http";
import { fileURLToPath } from "node:url";
import { dirname, resolve, extname, join, sep } from "node:path";
import { homedir } from "node:os";
import { createBridge } from "./bridge.mjs";
import { createOfficeBridgeMcp } from "./office-tools.mjs";
import { resolveWorkspaceRoot, suggestWorkspaceRoot, ensureWorkspaceMarker } from "./workspace.mjs";
import { randomUUID } from "node:crypto";
import { getSessionId, saveSessionId, touchFolder, clearSessionId } from "./sessions.mjs";
import { readTranscript, locateSessionFile } from "./transcript.mjs";
import { diag } from "./diag.mjs";
import { getContextEntries, setContextEntries } from "./context.mjs";
import { ApprovalManager, needsApproval } from "./approval.mjs";
import { stat } from "node:fs/promises";
import { existsSync } from "node:fs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = resolve(__dirname, "..");

// Model provider settings (ANTHROPIC_BASE_URL, ANTHROPIC_AUTH_TOKEN, ...) apply
// only to this daemon's agent, not to the user's global Claude Code.
const ENV_FILE = join(PROJECT_ROOT, ".env");
if (existsSync(ENV_FILE)) {
  process.loadEnvFile(ENV_FILE);
  console.log(`[daemon] Loaded model provider settings from ${ENV_FILE}`);
}

// Draftspect deliberately uses 47833/47834. Another local Office add-in
// on this machine may bind 47823/47824, so a distinct pair avoids a port
// clash when both run at once. Keep WS_PORT = HTTP_PORT - 1 (the taskpane
// derives nothing; the WS port is referenced explicitly in taskpane.js
// and the index.html CSP — change all three together).
const WS_PORT = 47833;
const HTTP_PORT = 47834;
const HTTP_ORIGIN = `http://127.0.0.1:${HTTP_PORT}`;
const TOKEN_FILE = join(homedir(), ".claude", "office-addins", "bridge-token");

// Bridge token — random per-daemon-start. The taskpane fetches it from the
// HTTP server's /bridge-token endpoint (same-origin, CORS-restricted) and
// includes it in the first WS hello. Any WS that doesn't present this token
// (or comes from an unknown origin) is closed.
const BRIDGE_TOKEN = randomBytes(24).toString("hex");
{
  await mkdir(dirname(TOKEN_FILE), { recursive: true });
  await writeFile(TOKEN_FILE, BRIDGE_TOKEN, { mode: 0o600 });
  try {
    await chmod(TOKEN_FILE, 0o600);
  } catch {}
  console.log(`[daemon] Bridge token written to ${TOKEN_FILE}`);
}

// ---------------------------------------------------------------------------
// CLI args
// ---------------------------------------------------------------------------
const matterFolder = process.argv[2]
  ? resolve(process.argv[2])
  : process.env.MATTER_FOLDER
    ? resolve(process.env.MATTER_FOLDER)
    : process.cwd();

console.log(`[daemon] Workspace folder (agent cwd): ${matterFolder}`);

// ---------------------------------------------------------------------------
// HTTP server: serve the taskpane assets so Excel can load them.
// ---------------------------------------------------------------------------
const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "application/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".ico": "image/x-icon",
  ".json": "application/json; charset=utf-8",
  ".xml": "application/xml; charset=utf-8",
};

const taskpaneDir = join(PROJECT_ROOT, "taskpane");
// Office.js served locally so the pane works when appsforoffice.microsoft.com is unreachable.
const officeJsDir = join(PROJECT_ROOT, "node_modules", "@microsoft", "office-js", "dist");
// Browser ES modules the pane imports straight from node_modules.
const NPM_MODULES = {
  "/npm/marked.esm.js": join(PROJECT_ROOT, "node_modules", "marked", "lib", "marked.esm.js"),
  "/npm/purify.es.mjs": join(PROJECT_ROOT, "node_modules", "dompurify", "dist", "purify.es.mjs"),
};

// A branded, actionable error page. Office renders whatever the manifest's
// SourceLocation returns inside the task pane, so a bare "Not found" (the
// old body) left users staring at two unhelpful words. The common real
// causes are all recoverable; spell them out. Served only to document
// navigations (Accept: text/html) — asset fetches still get terse text so
// nothing tries to parse HTML as JS/CSS.
const htmlEscape = (s) =>
  String(s).replace(
    /[&<>"']/g,
    (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" })[c],
  );

function errorPageHtml(status, headline, requestedPath) {
  requestedPath = htmlEscape(requestedPath);
  return `<!doctype html><html><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Excel Assistant — ${status}</title>
<style>
  body{font:14px/1.5 -apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
       color:#1a1a1a;margin:0;padding:28px 24px;background:#fff}
  h1{font-size:18px;margin:0 0 4px}
  .sub{color:#666;margin:0 0 18px}
  ol{padding-left:20px;margin:0 0 18px} li{margin:6px 0}
  code{background:#f2f2f2;padding:1px 5px;border-radius:4px;font-size:12px}
  .foot{color:#999;font-size:12px;border-top:1px solid #eee;padding-top:12px}
</style></head><body>
<h1>Excel Assistant couldn't load this panel</h1>
<p class="sub">${headline}</p>
<p>This usually means one of:</p>
<ol>
  <li>The <strong>Excel Assistant tray app isn't running</strong> — start it, then reopen this panel.</li>
  <li>Excel cached an old add-in — <strong>fully quit the app (⌘Q / Alt+F4) and reopen it</strong> so it re-reads the add-in.</li>
  <li>This add-in's manifest points at a <strong>different port</strong> than the running Excel Assistant daemon (e.g. another add-in's daemon answered). Relaunch Excel Assistant, then quit &amp; reopen Excel.</li>
  <li>If it persists, reinstall the add-in from the Excel Assistant tray menu.</li>
</ol>
<p class="foot">Excel Assistant daemon on <code>127.0.0.1:${HTTP_PORT}</code> · requested <code>${requestedPath}</code> · ${status}</p>
</body></html>`;
}

function sendError(req, res, status, headline, requestedPath) {
  const wantsHtml = (req.headers.accept || "").includes("text/html");
  if (wantsHtml) {
    res.writeHead(status, { "Content-Type": "text/html; charset=utf-8" });
    res.end(errorPageHtml(status, headline, requestedPath));
  } else {
    res.writeHead(status).end(headline);
  }
}

const http = createServer(async (req, res) => {
  // Restrictive CORS: only the taskpane's same origin (HTTP_ORIGIN) gets the
  // Access-Control-Allow-Origin header. Other origins (malicious web pages
  // running in a regular browser tab) hit a no-CORS-header response and the
  // browser blocks them from reading it. Same-origin requests from the
  // taskpane itself don't go through CORS at all.
  const reqOrigin = req.headers.origin || "";
  if (reqOrigin === HTTP_ORIGIN) {
    res.setHeader("Access-Control-Allow-Origin", reqOrigin);
  }
  // No-cache: Word's webview likes to cache aggressively. During dev we want
  // every reload to pick up the latest taskpane JS/CSS/HTML.
  res.setHeader("Cache-Control", "no-store, no-cache, must-revalidate, max-age=0");
  res.setHeader("Pragma", "no-cache");
  res.setHeader("Expires", "0");
  if (req.method === "OPTIONS") {
    res.writeHead(204).end();
    return;
  }

  try {
    const urlPath = (req.url || "/").split("?")[0];

    // Token endpoint: serves the per-daemon bridge token to the taskpane.
    // Only same-origin (i.e. cross-origin requests get blocked by CORS).
    if (urlPath === "/bridge-token") {
      res.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
      res.end(BRIDGE_TOKEN);
      return;
    }

    if (urlPath.startsWith("/eval/")) {
      await handleEvalRequest(req, res, urlPath);
      return;
    }

    if (NPM_MODULES[urlPath]) {
      res.writeHead(200, { "Content-Type": MIME[".js"] });
      res.end(await readFile(NPM_MODULES[urlPath]));
      return;
    }

    const isOfficeJs = urlPath.startsWith("/office-js/");
    const baseDir = isOfficeJs ? officeJsDir : taskpaneDir;
    const relPath = isOfficeJs
      ? urlPath.slice("/office-js".length)
      : urlPath === "/"
        ? "/index.html"
        : urlPath;
    const fsPath = join(baseDir, relPath);
    // Containment check. `join` already normalizes `../`, so the obvious
    // traversal is blocked — but a bare startsWith(taskpaneDir) would also
    // accept a sibling like `<…>/taskpane-evil/x`. Require an exact match
    // OR a path under `taskpaneDir` + separator.
    if (fsPath !== baseDir && !fsPath.startsWith(baseDir + sep)) {
      res.writeHead(403).end("Forbidden");
      return;
    }
    const ext = extname(fsPath).toLowerCase();
    const mime = MIME[ext] || "application/octet-stream";
    if (ext === ".html") {
      // Office's webview caches the JS/CSS bundle aggressively and ignores
      // our no-store headers — so a taskpane code change wouldn't take
      // effect even after reopening the pane. Cache-bust the local asset
      // refs with the per-start bridge token: every daemon restart yields a
      // fresh URL, forcing a re-fetch. (The handler strips the query string
      // before resolving the file, so `?v=` doesn't affect routing.)
      let html = await readFile(fsPath, "utf8");
      html = html.replace(/(\/shared\/(?:taskpane\.js|styles\.css))"/g, `$1?v=${BRIDGE_TOKEN}"`);
      res.writeHead(200, { "Content-Type": mime });
      res.end(html);
    } else {
      const data = await readFile(fsPath);
      res.writeHead(200, { "Content-Type": mime });
      res.end(data);
    }
  } catch (err) {
    const reqPath = (req.url || "/").split("?")[0];
    if (err.code === "ENOENT") {
      sendError(
        req,
        res,
        404,
        "That page or file isn’t served by this Excel Assistant daemon.",
        reqPath,
      );
    } else {
      console.error("[http]", err);
      sendError(
        req,
        res,
        500,
        "The Excel Assistant daemon hit an internal error serving this page.",
        reqPath,
      );
    }
  }
});

http.listen(HTTP_PORT, "127.0.0.1", () => {
  console.log(`[daemon] HTTP server listening on http://127.0.0.1:${HTTP_PORT}/`);
});

// ---------------------------------------------------------------------------
// IPC channel back to the Electron main process. Used to ask main.mjs to
// show a native macOS open panel for folder/file picking — synthetic
// in-page modals can't navigate to Google Drive, iCloud, recent items, or
// any of the other sources macOS users expect in NSOpenPanel. The channel
// is fd 3 (added in main.mjs's spawn options), wired through Node IPC.
// ---------------------------------------------------------------------------
const pendingPicks = new Map(); // id -> { resolve, reject, timer }
const PICK_TIMEOUT_MS = 5 * 60_000; // 5 min; the user might leave the dialog open

if (process.send) {
  process.on("message", (msg) => {
    if (msg?.type !== "pick_path_result") return;
    const entry = pendingPicks.get(msg.id);
    if (!entry) return;
    pendingPicks.delete(msg.id);
    clearTimeout(entry.timer);
    entry.resolve(msg);
  });
}

function pickPathFromMain({ include_files, default_path, title, button_label }) {
  if (!process.send) {
    return Promise.reject(new Error("No IPC channel to Electron main process"));
  }
  const id = randomUUID();
  return new Promise((resolve, reject) => {
    const timer = setTimeout(() => {
      pendingPicks.delete(id);
      reject(new Error("Picker timed out"));
    }, PICK_TIMEOUT_MS);
    pendingPicks.set(id, { resolve, reject, timer });
    process.send({ type: "pick_path", id, include_files, default_path, title, button_label });
  });
}

// Word and Excel are fully independent surfaces: each host gets its own
// agent loop, session, message queue, transcript AND workspace. Nothing
// one host does ever touches the other (no shared "current session", no
// cross-host interrupt). Both maps are hoisted above createBridge so
// bridge handlers that fire during the top-level awaits below
// (preflightHttpMcpServers, etc.) don't hit a TDZ on these bindings.
//
//   sessions:        paneKey -> live session { key, host, cwd, sessionId,
//                    abortController, settled, generation }
//   workspaceByKey:  paneKey -> last-known cwd, so a pane that connects
//                    before its first message still resolves a workspace
//                    (and survives that pane's loop ending).
//
// paneKey identifies one open document (host + doc path), so two Word
// docs (or two workbooks) open at once each get their own independent
// loop/session/queue/workspace. The host is still carried for the tool
// family + per-host behavior. Persisted SDK transcript resume uses the
// same document identity independently from its workspace folder.
const sessions = new Map();
const workspaceByKey = new Map();
const sessionGenerationByKey = new Map();
const replayTokenByKey = new Map();

function invalidateReplay(key) {
  const token = Symbol("replay");
  replayTokenByKey.set(key, token);
  return token;
}

function cancelPaneSession(key) {
  sessionGenerationByKey.set(key, Symbol("session"));
  invalidateReplay(key);
  const pending = startQueue.get(key);
  if (pending) pending.latest = null;
  // An obsolete initialization may still be awaiting I/O. A new generation
  // gets its own runner instead of waiting for that obsolete start to finish.
  startQueue.delete(key);
  const live = sessionFor(key);
  sessions.delete(key);
  live?.abortController.abort();
  bridge.clearUserMessages(key);
  return live;
}
// Panes whose workspace the user set explicitly (Change workspace). Their
// workspace must NOT be re-derived from the document folder on reconnect.
const explicitWorkspaceKeys = new Set();
// Per-pane sticky model choice (set by the taskpane composer). The SDK
// model is fixed per agent loop, so a change while a loop is live triggers
// a resuming restart (see the set_model handler).
const modelByKey = new Map();
const ALLOWED_MODELS = new Set(["haiku", "sonnet", "opus"]);

// The model alias to pass to query(). Draftspect always pins an explicit
// model (cost-control UI) — never silently inherits the CLI default, which
// can be Opus. Falls back to "sonnet" if the taskpane hasn't said yet.
function modelArgFor(key) {
  const m = modelByKey.get(key);
  return ALLOWED_MODELS.has(m) ? m : "sonnet";
}

function sessionFor(key) {
  return sessions.get(key) ?? null;
}

function isCurrentSession(session) {
  return (
    sessionFor(session.key) === session &&
    sessionGenerationByKey.get(session.key) === session.generation &&
    !session.abortController.signal.aborted
  );
}

function retireCurrentSession(session, { clearMessages = false } = {}) {
  if (sessionFor(session.key) !== session) return false;
  session.settled = true;
  sessions.delete(session.key);
  session.abortController.abort();
  if (clearMessages) bridge.clearUserMessages(session.key);
  return true;
}

// This pane's workspace: its live session's cwd, else the last cwd we
// recorded for it, else the launch default.
function cwdForKey(key) {
  return sessionFor(key)?.cwd ?? workspaceByKey.get(key) ?? matterFolder;
}

// The bridge key is host + NUL + document URL (or anon:<pane id> for an
// unsaved document). Persist conversations by that document portion; cwd is
// only the workspace used for files and context.
function documentKeyForPane(key) {
  if (typeof key !== "string") return null;
  const separator = key.indexOf("\0");
  return separator === -1 ? key : key.slice(separator + 1);
}

// Resolve the session id to replay for this pane. Prefer the pane's live
// session id (set once the SDK reports init); otherwise the id persisted
// for (host, document) — covers the window before the SDK has re-inited.
async function resolveReplaySessionId(key, host, cwd) {
  const live = sessionFor(key);
  if (live?.sessionId) return live.sessionId;
  if (!host || !cwd) return null;
  try {
    return (await getSessionId(host, documentKeyForPane(key))) ?? null;
  } catch {
    return null;
  }
}

// Reconstruct this pane's prior conversation from its .jsonl and push it
// to THAT pane only. Sent on every taskpane hello and after a workspace
// switch (cwd_changed). Empty events => fresh chat, no divider.
async function sendTranscriptReplayTo(key, host, cwd, token = invalidateReplay(key)) {
  try {
    const sessionId = await resolveReplaySessionId(key, host, cwd);
    const { events, truncated } = sessionId
      ? await readTranscript(sessionId, { maxEvents: 200 })
      : { events: [], truncated: false };
    if (replayTokenByKey.get(key) !== token || cwdForKey(key) !== cwd) return;
    bridge.sendToTaskpane(
      {
        type: "transcript_replay",
        session_id: sessionId ?? null,
        truncated,
        events,
      },
      key,
    );
  } catch (err) {
    console.warn("[daemon] transcript replay failed:", err?.message ?? err);
  }
}

// Serialize/coalesce starts within a pane generation. Cancellation retires
// the runner; every awaited initialization checks its generation before
// publishing events or starting the SDK. Stop and stream failures recover
// lazily on the next user message, rather than starting an idle SDK loop.
const startQueue = new Map(); // key -> { running: boolean, latest: req | null }

function scheduleSessionStart(cwd, sessionId, key, host, reason, { replay = true } = {}) {
  let q = startQueue.get(key);
  if (!q) {
    q = { running: false, latest: null };
    startQueue.set(key, q);
  }
  q.latest = {
    cwd,
    resumeId: sessionId,
    host: host ?? null,
    reason,
    replay,
    generation: sessionGenerationByKey.get(key),
  };
  if (q.running) return; // the active runner will pick up `latest`
  q.running = true;
  setImmediate(() => runStartQueue(key));
}

async function runStartQueue(key) {
  const q = startQueue.get(key);
  if (!q) return;
  // startSessionForFolder returns once setup is done (it kicks the agent
  // loop off detached), so awaiting it serializes only the abort+create
  // step — exactly the part that must not interleave.
  while (startQueue.get(key) === q && q.latest) {
    const req = q.latest;
    q.latest = null;
    if (req.generation !== sessionGenerationByKey.get(key)) continue;
    try {
      await startSessionForFolder(req.cwd, req.resumeId, {
        key,
        host: req.host,
        replay: req.replay,
        generation: req.generation,
      });
    } catch (err) {
      console.error(`[daemon] ${req.reason} session start failed:`, err?.message ?? err);
    }
  }
  q.running = false;
  if (startQueue.get(key) === q) startQueue.delete(key);
}

// Called on every taskpane hello — for ANY open document, many panes
// possibly connected at once. Connecting a pane must NOT start the agent
// loop: connect-driven starts amplified the old connect/disconnect
// ping-pong and burn an unasked turn. We only re-render this pane's own
// transcript. The loop starts lazily on the first user message
// (onUserMessage → ensureLoopForMessage).
// Pending workspace resolution per pane. The bridge sends `welcome` without
// waiting for onPaneConnect, so a message that arrives right after connect
// must wait for this, or it starts in the daemon's default folder.
const workspaceResolving = new Map(); // key -> Promise<void>

async function awaitWorkspaceResolution(key) {
  // Failed mutations report their own error; readers use the last valid
  // workspace rather than inheriting another request's rejection.
  while (workspaceResolving.has(key)) await workspaceResolving.get(key).catch(() => {});
}

// Serialize directory resolution, explicit switches and New chat's history
// reset. Readers wait for directory state, never for transcript I/O.
function updateWorkspace(key, action) {
  const previous = workspaceResolving.get(key);
  const work = Promise.resolve(previous)
    .catch(() => {})
    .then(action);
  const tracked = work.finally(() => {
    if (workspaceResolving.get(key) === tracked) workspaceResolving.delete(key);
  });
  workspaceResolving.set(key, tracked);
  return tracked;
}

function onPaneConnect(key, host, doc) {
  if (!key) return undefined;
  const token = invalidateReplay(key);
  return updateWorkspace(key, () => resolvePaneWorkspace(key, host, doc)).then(async () => {
    await sendTranscriptReplayTo(key, host, cwdForKey(key), token);
    approvalManager.replay(key);
  });
}

async function resolvePaneWorkspace(key, host, doc) {
  // Resolve this pane's workspace from the open document's own folder,
  // server-side and immediately — deterministic, no loop start. Without
  // this, cwdForKey() falls back to matterFolder (the daemon's launch
  // cwd, often the repo root) until a message or the taskpane's set_cwd
  // round-trip lands, so get_context / replay read the WRONG folder's
  // CLAUDE.md (the symptom: context files bleeding across documents).
  // Skip if the user explicitly pinned this pane's workspace, or a live
  // session already owns the cwd.
  if (doc && !explicitWorkspaceKeys.has(key) && !sessionFor(key)) {
    try {
      const folder = await resolveWorkspaceRoot(doc);
      if (
        folder &&
        !explicitWorkspaceKeys.has(key) &&
        !sessionFor(key) &&
        bridge.isTaskpaneConnected(key)
      ) {
        workspaceByKey.set(key, folder);
      }
    } catch {
      /* unresolvable (cloud/unsaved) — keep the fallback */
    }
  }
  const cwd = cwdForKey(key);
  diag(`hello → replay key=${key} cwd=${cwd} (no loop start on connect)`);
}

// Called when a pane's WebSocket closes for good (the bridge already
// freed its own pane/queue state). Prune the daemon-side per-key Maps so
// they don't grow for the life of the daemon — but ONLY when no live
// session owns this key. A session is deliberately kept alive across a
// transient disconnect (it resumes on reconnect with the same stable
// key); its own `finally` clears `sessions` when its loop actually ends.
function onPaneClose(key) {
  if (!key || sessionFor(key)) return;
  cancelPaneSession(key);
  workspaceByKey.delete(key);
  explicitWorkspaceKeys.delete(key);
  modelByKey.delete(key);
  startQueue.delete(key);
  approvalManager.clearKey(key);
}

// Called when a user message arrives from a pane, BEFORE it's queued.
// Each pane has its OWN loop — independent of every other pane. If this
// pane's loop is already live, do nothing (its userMessageStream will
// consume the message). Otherwise start it, resuming this workbook's
// conversation. Deferred via setImmediate so it lands after the message
// is queued and after any in-flight finally; the new loop then drains
// this pane's queue. No other pane's loop is ever touched.
// Drop this pane's conversation; the next user message lazily starts a fresh
// session (ensureLoopForMessage finds no saved id). Cancellation also retires
// queued starts and makes the aborted loop's catch/finally harmless.
function startNewConversation(key, host) {
  cancelPaneSession(key);
  const generation = sessionGenerationByKey.get(key);
  const token = replayTokenByKey.get(key);
  bridge.sendAssistantEvent({ event: "turn_complete", interrupted: true }, key);
  return updateWorkspace(key, async () => {
    const cwd = cwdForKey(key);
    await clearSessionId(host, documentKeyForPane(key));
    if (sessionGenerationByKey.get(key) !== generation) return;
    if (replayTokenByKey.get(key) === token) {
      bridge.sendToTaskpane(
        { type: "transcript_replay", session_id: null, truncated: false, events: [] },
        key,
      );
    }
    console.log(`[daemon] New conversation for ${cwd} (${host})`);
  });
}

async function ensureLoopForMessage(key, host) {
  if (!key) return;
  invalidateReplay(key);
  const generation = sessionGenerationByKey.get(key);
  await awaitWorkspaceResolution(key);
  if (sessionGenerationByKey.get(key) !== generation) return;
  const cwd = cwdForKey(key);
  const live = sessionFor(key);
  if (live && !live.settled) {
    return; // this pane's loop is live and will consume the message
  }
  let resumeId = null;
  try {
    resumeId = await getSessionId(host, documentKeyForPane(key));
  } catch {
    /* fresh session if lookup fails */
  }
  if (sessionGenerationByKey.get(key) !== generation || sessionFor(key)) return;
  diag(`message → ensure loop key=${key} cwd=${cwd} resume=${resumeId ?? "(new)"}`);
  // replay:false — the pane already shows the chat (incl. the message
  // that just triggered this). An empty transcript_replay here (a
  // brand-new session has no .jsonl yet) would wipe the user's prompt.
  scheduleSessionStart(cwd, resumeId, key, host, "user message", { replay: false });
}

// ---------------------------------------------------------------------------
// WebSocket bridge.
// ---------------------------------------------------------------------------
let bridge;
const approvalManager = new ApprovalManager({
  sendEvent: (event, key) => bridge?.sendAssistantEvent(event, key),
});

bridge = createBridge({
  port: WS_PORT,
  token: BRIDGE_TOKEN,
  allowedOrigins: [HTTP_ORIGIN],
  onHello: (key, host, doc) => onPaneConnect(key, host, doc),
  onUserMessage: (key, host) => ensureLoopForMessage(key, host),
  onClose: (key) => onPaneClose(key),
  extraHandlers: {
    pick_path: async (msg, reply) => {
      try {
        const result = await pickPathFromMain({
          include_files: !!msg.include_files,
          default_path: msg.default_path || null,
          title: msg.title || null,
          button_label: msg.button_label || null,
        });
        reply({ type: "pick_path_result", request_id: msg.request_id, ...result });
      } catch (e) {
        reply({
          type: "pick_path_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    set_cwd: async (msg, reply, key, host) => {
      try {
        await updateWorkspace(key, async () => {
          let cwd;
          let explicitPick = false;
          if (msg.autodetect_from_doc) {
            const detected = await resolveWorkspaceRoot(msg.autodetect_from_doc);
            if (!detected)
              throw new Error("Could not auto-detect a workspace folder from that doc path");
            cwd = detected;
          } else if (msg.cwd) {
            cwd = msg.cwd;
            explicitPick = true;
          } else {
            throw new Error("set_cwd requires `cwd` or `autodetect_from_doc`");
          }
          // Remember an explicit pick so a later reconnect doesn't re-derive
          // this pane's workspace from the doc folder; an autodetect switch
          // clears that pin (the doc folder is authoritative again).
          const resolved = await switchFolder(cwd, key, host);
          if (explicitPick) explicitWorkspaceKeys.add(key);
          else explicitWorkspaceKeys.delete(key);
          // Drop a CLAUDE.md marker on explicit user pick so the next open of
          // any doc in this folder auto-detects silently.
          let markerCreated = false;
          if (explicitPick) {
            try {
              markerCreated = await ensureWorkspaceMarker(resolved);
            } catch (e) {
              console.warn(`[daemon] could not create CLAUDE.md in ${resolved}: ${e.message}`);
            }
          }
          reply({
            type: "set_cwd_result",
            ok: true,
            cwd: resolved,
            marker_created: markerCreated,
            request_id: msg.request_id,
          });
        });
      } catch (e) {
        reply({ type: "set_cwd_result", ok: false, error: e.message, request_id: msg.request_id });
      }
    },
    suggest_workspace: async (msg, reply) => {
      try {
        const suggestion = await suggestWorkspaceRoot(msg.doc_path || null);
        reply({
          type: "suggest_workspace_result",
          ok: true,
          suggestion,
          request_id: msg.request_id,
        });
      } catch (e) {
        reply({
          type: "suggest_workspace_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    get_cwd_state: async (msg, reply, key) => {
      await awaitWorkspaceResolution(key);
      reply({
        type: "get_cwd_state_result",
        ok: true,
        // This pane's own workspace, resolvable even before its first
        // message (lazy start ⇒ no session yet).
        current_cwd: cwdForKey(key),
        request_id: msg.request_id,
      });
    },
    stop_agent: async (msg, reply, key) => {
      // Also cancel starts still waiting for workspace/session lookup.
      // Abort releases the old bridge waiter immediately. The next user
      // message starts lazily; no idle SDK restart is needed after Stop.
      cancelPaneSession(key);
      bridge.sendAssistantEvent({ event: "turn_complete", interrupted: true }, key);
      reply({ type: "stop_agent_result", ok: true, request_id: msg.request_id });
    },
    get_models: async (msg, reply) => {
      // Real model ids behind each tier when a non-Anthropic provider is
      // configured in .env; null means the tier uses Claude as-is.
      reply({
        type: "get_models_result",
        ok: true,
        models: {
          haiku: process.env.ANTHROPIC_DEFAULT_HAIKU_MODEL || null,
          sonnet: process.env.ANTHROPIC_DEFAULT_SONNET_MODEL || null,
          opus: process.env.ANTHROPIC_DEFAULT_OPUS_MODEL || null,
        },
        request_id: msg.request_id,
      });
    },
    new_session: async (msg, reply, key, host) => {
      try {
        await startNewConversation(key, host);
        reply({ type: "new_session_result", ok: true, request_id: msg.request_id });
      } catch (e) {
        reply({
          type: "new_session_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    get_context: async (msg, reply, key) => {
      try {
        await awaitWorkspaceResolution(key);
        const cwd = cwdForKey(key);
        const entries = cwd ? await getContextEntries(cwd) : [];
        reply({ type: "get_context_result", ok: true, cwd, entries, request_id: msg.request_id });
      } catch (e) {
        reply({
          type: "get_context_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
    set_approval: async (msg, reply, key) => {
      approvalManager.setEnabled(key, Boolean(msg.enabled));
      reply({ type: "set_approval_result", ok: true, enabled: Boolean(msg.enabled), request_id: msg.request_id });
    },
    approval_response: async (msg, reply, key) => {
      const accepted = approvalManager.respond(key, msg.approval_request_id, msg.decision);
      reply({
        type: "approval_response_result",
        ok: accepted,
        decision: msg.decision,
        request_id: msg.request_id,
        ...(accepted ? {} : { error: "This approval request is no longer pending." }),
      });
    },
    set_model: async (msg, reply, key, host) => {
      // A pane that reconnects mid-evaluation (e.g. after machine sleep)
      // re-sends its sticky UI model; the evaluation's model must win.
      const evaluation = evalObservers.get(key);
      if (evaluation) {
        const requested = String(msg.model || "").trim();
        if (ALLOWED_MODELS.has(requested)) evaluation.restoreModel = requested;
        reply({
          type: "set_model_result",
          ok: true,
          model: evaluation.tier,
          request_id: msg.request_id,
        });
        return;
      }
      const requested = String(msg.model || "").trim();
      const model = ALLOWED_MODELS.has(requested) ? requested : "sonnet";
      const prev = modelByKey.get(key);
      modelByKey.set(key, model);
      reply({ type: "set_model_result", ok: true, model, request_id: msg.request_id });
      // Only relaunch if the model actually changed for an already-known
      // pane. On the initial connect `prev` is undefined (no live session
      // yet — the lazy first-message start will read modelByKey), so we
      // just record it. restartSession itself no-ops when no session.
      if (prev !== undefined && prev !== model) {
        console.log(`[daemon] model changed → ${model} (key bound); restarting loop`);
        restartSession(key, host, { reason: "model_changed" }).catch((err) =>
          console.warn("[daemon] restart after model change failed:", err.message),
        );
      }
    },
    set_context: async (msg, reply, key, host) => {
      try {
        await updateWorkspace(key, async () => {
          const cwd = cwdForKey(key);
          if (!cwd) throw new Error("No workspace selected");
          if (msg.expected_cwd && !samePath(resolve(msg.expected_cwd), resolve(cwd))) {
            throw new Error("Workspace changed before context files could be saved");
          }
          const { saved, errors } = await setContextEntries(cwd, msg.entries || []);
          reply({
            type: "set_context_result",
            ok: errors.length === 0,
            cwd,
            saved,
            errors,
            request_id: msg.request_id,
          });
          // Restart THIS pane's loop so its agent re-reads CLAUDE.md and
          // picks up the updated context block on the next turn. Every
          // other pane is unaffected.
          restartSession(key, host, { reason: "context_changed" }).catch((err) =>
            console.warn("[daemon] restart failed:", err.message),
          );
        });
      } catch (e) {
        reply({
          type: "set_context_result",
          ok: false,
          error: e.message,
          request_id: msg.request_id,
        });
      }
    },
  },
});

// ---------------------------------------------------------------------------
// Evaluation hooks. An external runner (evals/) opens a workbook whose task
// pane auto-opens, then POSTs a prompt to /eval/run; the prompt is injected as
// if typed in that pane and the call resolves when the turn completes.
// ---------------------------------------------------------------------------
const evalObservers = new Map(); // paneKey -> { text, tools, finish }

for (const method of ["sendAssistantEvent", "sendAssistantText"]) {
  const original = bridge[method];
  bridge[method] = (payload, key) => {
    const observer = key && evalObservers.get(key);
    if (observer) {
      if (method === "sendAssistantText") observer.text += payload;
      else if (payload.event === "tool_use_announce") observer.tools.push(payload.tool);
      else if (payload.event === "session_init") {
        observer.model = payload.model;
        observer.sessionId = payload.session_id ?? null;
      } else if (payload.event === "turn_complete" && !payload.interrupted)
        observer.finish({
          status: payload.subtype === "success" ? "completed" : payload.subtype || "completed",
          error: payload.error,
          usage: payload.usage,
          numTurns: payload.num_turns,
          costUsd: payload.total_cost_usd,
        });
      else if (payload.event === "error" || payload.event === "auth_error")
        observer.finish({ status: "error", error: payload.error });
    }
    return original(payload, key);
  };
}

const samePath = (a, b) =>
  String(a || "")
    .replaceAll("/", "\\")
    .toLowerCase() ===
  String(b || "")
    .replaceAll("/", "\\")
    .toLowerCase();

async function runEvalPrompt({ doc, prompt, model, timeoutMs = 15 * 60_000, paneWaitMs = 90_000 }) {
  if (!doc || !prompt) return { status: "bad_request", error: "doc and prompt are required" };
  const waitUntil = Date.now() + paneWaitMs;
  let pane;
  while (!(pane = bridge.listPanes().find((p) => samePath(p.activeDoc, doc)))) {
    if (Date.now() > waitUntil)
      return { status: "no_pane", error: `No task pane connected for ${doc}` };
    await new Promise((r) => setTimeout(r, 1000));
  }
  const { key, host } = pane;
  if (evalObservers.has(key))
    return { status: "busy", error: "An evaluation is already running in this pane" };
  const previousModel = modelByKey.get(key);
  const evaluationModel = model && ALLOWED_MODELS.has(model) ? model : modelArgFor(key);
  const started = Date.now();
  const startedMono = performance.now();
  // A heartbeat that arrives far later than scheduled means the process was
  // suspended (machine sleep) or starved; such attempts are infra failures,
  // not model latency.
  const HEARTBEAT_MS = 5000;
  const STALL_MS = 30_000;
  const clock = { lastWall: started, lastMono: startedMono, maxWallGapMs: 0, maxMonoGapMs: 0 };
  return new Promise((resolve) => {
    let timer = null;
    let heartbeat = null;
    const observer = {
      tier: evaluationModel,
      restoreModel: previousModel,
      text: "",
      tools: [],
      finished: false,
      finish: (result) => {
        if (observer.finished) return;
        observer.finished = true;
        clearTimeout(timer);
        clearInterval(heartbeat);
        if (evalObservers.get(key) === observer) evalObservers.delete(key);
        // Do not leave the long-lived evaluation loop waiting for another
        // message under the temporary model. The next ordinary message will
        // start lazily with the pane's restored choice.
        if (sessionFor(key)) cancelPaneSession(key);
        if (observer.restoreModel === undefined) modelByKey.delete(key);
        else modelByKey.set(key, observer.restoreModel);
        const finished = Date.now();
        const wallGap = Math.max(clock.maxWallGapMs, finished - clock.lastWall);
        const monoGap = Math.max(clock.maxMonoGapMs, performance.now() - clock.lastMono);
        const summary = {
          ...result,
          model: observer.model ?? null,
          sessionId: observer.sessionId ?? null,
          text: observer.text,
          tools: observer.tools,
          startedAt: new Date(started).toISOString(),
          finishedAt: new Date(finished).toISOString(),
          durationMs: finished - started,
          monotonicMs: Math.round(performance.now() - startedMono),
          maxWallGapMs: Math.round(wallGap),
          maxMonotonicGapMs: Math.round(monoGap),
          stalled: wallGap > STALL_MS || monoGap > STALL_MS,
        };
        resolve(
          (async () => ({
            ...summary,
            transcriptPath: summary.sessionId ? await locateSessionFile(summary.sessionId) : null,
          }))(),
        );
      },
    };
    // Publish the ownership lock before the first await. A reconnect while
    // New chat clears history must not replace the evaluation's model.
    evalObservers.set(key, observer);
    modelByKey.set(key, evaluationModel);
    heartbeat = setInterval(() => {
      const wall = Date.now();
      const mono = performance.now();
      clock.maxWallGapMs = Math.max(clock.maxWallGapMs, wall - clock.lastWall);
      clock.maxMonoGapMs = Math.max(clock.maxMonoGapMs, mono - clock.lastMono);
      clock.lastWall = wall;
      clock.lastMono = mono;
    }, HEARTBEAT_MS);
    timer = setTimeout(() => {
      cancelPaneSession(key);
      bridge.sendAssistantEvent({ event: "turn_complete", interrupted: true }, key);
      observer.finish({ status: "timeout", timeoutMs });
    }, timeoutMs);
    (async () => {
      try {
        await startNewConversation(key, host);
        if (evalObservers.get(key) !== observer) return;
        bridge.sendAssistantEvent({ event: "info", message: `Evaluation prompt:\n${prompt}` }, key);
        ensureLoopForMessage(key, host).catch((err) => {
          if (evalObservers.get(key) === observer) {
            bridge.clearUserMessages(key);
            bridge.sendAssistantEvent({ event: "error", error: err.message }, key);
          }
        });
        bridge.pushUserMessage(prompt, key);
      } catch (err) {
        observer.finish({ status: "error", error: err?.message ?? String(err) });
      }
    })();
  });
}

async function runEvalTool({ doc, name, args = {}, paneWaitMs = 90_000 }) {
  if (!doc || !name) return { status: "bad_request", error: "doc and name are required" };
  const waitUntil = Date.now() + paneWaitMs;
  let pane;
  while (!(pane = bridge.listPanes().find((candidate) => samePath(candidate.activeDoc, doc)))) {
    if (Date.now() > waitUntil) {
      return { status: "no_pane", error: `No task pane connected for ${doc}` };
    }
    await new Promise((resolve) => setTimeout(resolve, 250));
  }
  try {
    return {
      status: "ok",
      result: await bridge.callTaskpaneTool(name, args, pane.key),
    };
  } catch (error) {
    return {
      status: "tool_error",
      error: error?.message ?? String(error),
      ...(error?.code ? { code: error.code } : {}),
      ...(error?.commitStatus ? { commitStatus: error.commitStatus } : {}),
    };
  }
}

// Everything that changes agent behavior without changing the repo commit,
// so a run manifest can refuse to mix configurations.
async function evalInfo() {
  const sdkPackage = join(
    PROJECT_ROOT,
    "node_modules",
    "@anthropic-ai",
    "claude-agent-sdk",
    "package.json",
  );
  const baseUrl = process.env.ANTHROPIC_BASE_URL;
  const configuredEnv = Object.fromEntries(
    Object.entries(agentConfig.env ?? {}).sort(([a], [b]) => a.localeCompare(b)),
  );
  return {
    node: process.version,
    sdkVersion: JSON.parse(await readFile(sdkPackage, "utf8")).version,
    provider: baseUrl ? new URL(baseUrl).host : "anthropic",
    models: {
      haiku: process.env.ANTHROPIC_DEFAULT_HAIKU_MODEL || "haiku",
      sonnet: process.env.ANTHROPIC_DEFAULT_SONNET_MODEL || "sonnet",
      opus: process.env.ANTHROPIC_DEFAULT_OPUS_MODEL || "opus",
    },
    systemPromptSha256: createHash("sha256")
      .update(await buildSystemPromptAppend())
      .digest("hex"),
    mcpServers: Object.keys(userMcpServers).sort(),
    plugins: agentPlugins.map((p) => p.path),
    skills: agentConfig.skills ?? null,
    builtinTools: agentConfig.builtinTools,
    disallowedTools: agentConfig.disallowedTools,
    settingSources: agentConfig.settingSources,
    env: {
      keys: Object.keys(configuredEnv),
      sha256: createHash("sha256").update(JSON.stringify(configuredEnv)).digest("hex"),
    },
  };
}

async function readJsonBody(req) {
  const chunks = [];
  for await (const chunk of req) chunks.push(chunk);
  return JSON.parse(Buffer.concat(chunks).toString("utf8") || "{}");
}

async function handleEvalRequest(req, res, urlPath) {
  const reply = (status, body) => {
    res.writeHead(status, { "Content-Type": "application/json; charset=utf-8" });
    res.end(JSON.stringify(body));
  };
  // Browsers can't read the token file, so a web page can't drive the agent.
  if (req.headers["x-bridge-token"] !== BRIDGE_TOKEN) return reply(401, { error: "unauthorized" });
  if (req.method === "GET" && urlPath === "/eval/panes") return reply(200, bridge.listPanes());
  if (req.method === "GET" && urlPath === "/eval/info") return reply(200, await evalInfo());
  if (req.method === "POST" && urlPath === "/eval/tool")
    return reply(200, await runEvalTool(await readJsonBody(req)));
  if (req.method === "POST" && urlPath === "/eval/run")
    return reply(200, await runEvalPrompt(await readJsonBody(req)));
  return reply(404, { error: "not found" });
}

// ---------------------------------------------------------------------------
// Office-bridge MCP server (in-process; forwards tool calls to the taskpane).
// Built per session so the registered tool family matches the connected
// host (see startSessionForFolder / the host re-narrow on hello).
// ---------------------------------------------------------------------------

// ---------------------------------------------------------------------------
// Load the user's MCP servers from ~/.claude.json — the same file the
// interactive `claude` CLI uses. The Agent SDK doesn't read this file by
// default (it reads `~/.claude/settings.json` instead), so without this step
// the user's configured servers (visio, etc.) would be invisible to the
// daemon.
// ---------------------------------------------------------------------------
async function loadUserMcpServers() {
  const configPath = join(homedir(), ".claude.json");
  try {
    const raw = await readFile(configPath, "utf8");
    const parsed = JSON.parse(raw);
    const servers = parsed?.mcpServers ?? {};
    return servers;
  } catch (err) {
    if (err.code === "ENOENT") return {};
    console.warn(`[daemon] Could not load MCP servers from ${configPath}:`, err.message);
    return {};
  }
}

// Project-scoped agent config (MCP servers such as the COM-based Excel server,
// skill plugins, which Claude Code features the agent gets) lives beside the
// daemon so it neither leaks into nor inherits from the user's global Claude
// Code config. Defaults measured to cut the first request from ~63k to ~13k
// tokens (docs/optimization-analysis.md, section 28).
const DEFAULT_AGENT_CONFIG = {
  mcpServers: {},
  plugins: [],
  skills: undefined,
  builtinTools: ["Read", "Glob", "Grep", "Skill", "ToolSearch", "WebSearch", "WebFetch"],
  disallowedTools: ["mcp__thepexcel-excel__excel_vba"],
  settingSources: ["project"],
  inheritUserMcpServers: false,
  env: { ENABLE_TOOL_SEARCH: "true" },
};

async function loadAgentConfig() {
  const configPath = join(PROJECT_ROOT, "agent.config.json");
  try {
    return { ...DEFAULT_AGENT_CONFIG, ...JSON.parse(await readFile(configPath, "utf8")) };
  } catch (err) {
    if (err.code !== "ENOENT") {
      console.warn(`[daemon] Could not load ${configPath}:`, err.message);
    }
    return DEFAULT_AGENT_CONFIG;
  }
}

const agentConfig = await loadAgentConfig();
// Read by the spawned Claude Code process (e.g. ENABLE_TOOL_SEARCH defers
// large MCP tool catalogs behind ToolSearch).
Object.assign(process.env, agentConfig.env);
const globalMcpServers = agentConfig.inheritUserMcpServers ? await loadUserMcpServers() : {};
const userMcpServers = { ...globalMcpServers, ...agentConfig.mcpServers };
const agentPlugins = agentConfig.plugins;
for (const [source, servers] of [
  ["~/.claude.json", globalMcpServers],
  ["agent.config.json", agentConfig.mcpServers],
]) {
  const names = Object.keys(servers);
  if (names.length > 0) {
    console.log(
      `[daemon] Loaded ${names.length} MCP server(s) from ${source}: ${names.join(", ")}`,
    );
  }
}
if (agentPlugins.length > 0) {
  console.log(
    `[daemon] Loaded ${agentPlugins.length} plugin(s): ${agentPlugins.map((p) => p.path).join(", ")}`,
  );
}

// Preflight HTTP MCP servers. The SDK will silently drop any server whose
// initial handshake fails, with no retry for the lifetime of the session
// (see memory: sdk-silently-drops-failed-mcp). We can't fix that here, but
// we can surface the failure in the daemon log so it's obvious why a tool
// is missing — "restart the daemon when the server is back up" instead of
// "no idea why Visio doesn't work".
async function preflightHttpMcpServers(servers) {
  const entries = Object.entries(servers).filter(([, cfg]) => cfg?.type === "http" && cfg.url);
  await Promise.all(
    entries.map(async ([name, cfg]) => {
      const body = JSON.stringify({
        jsonrpc: "2.0",
        id: 1,
        method: "initialize",
        params: {
          protocolVersion: "2024-11-05",
          capabilities: {},
          clientInfo: { name: "claude-code-office-preflight", version: "0.1" },
        },
      });
      try {
        const res = await fetch(cfg.url, {
          method: "POST",
          headers: {
            "Content-Type": "application/json",
            Accept: "application/json, text/event-stream",
          },
          body,
          signal: AbortSignal.timeout(5000),
        });
        if (res.ok) {
          console.log(`[daemon] MCP preflight: ${name} reachable at ${cfg.url}`);
        } else {
          console.warn(
            `[daemon] MCP preflight: ${name} returned HTTP ${res.status} — tools may be missing until daemon restart`,
          );
        }
      } catch (err) {
        const reason =
          err.name === "TimeoutError" ? "timeout after 5s" : err.cause?.code || err.message;
        console.warn(
          `[daemon] MCP preflight: ${name} unreachable (${reason}) — tools will be missing until daemon restart`,
        );
      }
    }),
  );
}
await preflightHttpMcpServers(userMcpServers);

// ---------------------------------------------------------------------------
// System prompt: Claude Code default + Office-specific append.
// ---------------------------------------------------------------------------
// Shared base + the Excel section. Re-read fresh on every session start so
// edits take effect when a session restarts (no daemon restart required).
async function buildSystemPromptAppend() {
  const base = await readFile(join(__dirname, "system-prompt.md"), "utf8");
  const excel = await readFile(join(__dirname, "system-prompt-excel.md"), "utf8");
  return [base, excel].join("\n");
}

// ---------------------------------------------------------------------------
// Permission handler. Hard rule: filesystem tools must never WRITE to Office
// files. The workbook the user has open is held with unsaved changes; a
// filesystem write clobbers their work and can corrupt the file. Reading is
// allowed.
// ---------------------------------------------------------------------------
// Path/extension test for the Office-managed file types.
const OFFICE_FILE_EXT = /\.(docx?|xlsx?|docm|xlsm)\b/i;

// Bash is allowed to *read* Office files (the agent legitimately does
// `unzip -p draft.docx word/document.xml`, `cat`, `git log -- report.docx`,
// `ls *.docx`). We only deny commands that mutate one in place: a shell
// redirection whose target is an Office path, or a destructive verb used
// against one. This is an accident guard, not a security boundary — it
// stops the obvious foot-gun (`cp draft.docx active.docx`,
// `rm -rf workspace/*.docx`, `echo x > active.xlsx`, `sed -i … report.docx`),
// not a determined evasion (base64 payloads, here-docs, etc.).
const REDIR_TO_OFFICE = />>?\s*['"]?[^'"|;&\s]*\.(docx?|xlsx?|docm|xlsm)\b/i;
const DESTRUCTIVE_VERB = /\b(rm|mv|cp|dd|shred|truncate|install|tee|ln)\b/i;
const SED_IN_PLACE = /\bsed\b[^|;&]*\s-i\b/i;

function bashMutatesOfficeFile(cmd) {
  if (REDIR_TO_OFFICE.test(cmd)) return true;
  if (!OFFICE_FILE_EXT.test(cmd)) return false;
  return DESTRUCTIVE_VERB.test(cmd) || SED_IN_PLACE.test(cmd);
}

function denyWithOfficeMessage() {
  return {
    behavior: "deny",
    message:
      "Refusing to write/move/delete an Office file via filesystem tools. These files are " +
      "managed by Office and may have unsaved changes; a filesystem mutation can corrupt the " +
      "active workbook. Use the excel_* tools to change workbook contents.",
  };
}

// ---------------------------------------------------------------------------
// Approve before apply (optional, per pane, off by default). The idea comes
// from MS-Excel-AI-plugin's staged change-sets and ExcelLLMAddin's
// approve-before-apply; here each workbook-changing call waits in canUseTool
// for the user's decision, so the model still sees every real result.
// ---------------------------------------------------------------------------
async function permissionFor(key, session, host, toolName, input) {
  // Evaluations run unattended even if the user turned approvals on in the pane.
  if (approvalManager.isEnabled(key) && !evalObservers.has(key) && needsApproval(toolName, input)) {
    const decision = await approvalManager.request(key, session, toolName, input);
    if (decision === "approve_turn" && session) session.approveRestOfTurn = true;
    if (decision === "reject") {
      return {
        behavior: "deny",
        message: "The user rejected this workbook change. Do not retry it; ask what they would like instead.",
      };
    }
    if (decision === "timeout") {
      return {
        behavior: "deny",
        message: "Approval for this workbook change timed out. The user did not reject it; ask them to try again when the task pane is connected.",
      };
    }
    if (decision === "cancelled") {
      return {
        behavior: "deny",
        message: "Approval for this workbook change was cancelled because the turn or pane session ended.",
      };
    }
  }
  return customPermissionHandler(toolName, input, { host });
}

function customPermissionHandler(toolName, input, { host = null } = {}) {
  if (toolName === "Write" || toolName === "Edit" || toolName === "MultiEdit") {
    const path = input?.file_path ?? input?.path;
    if (typeof path === "string" && OFFICE_FILE_EXT.test(path)) {
      return Promise.resolve(denyWithOfficeMessage());
    }
  }
  if (toolName === "Bash") {
    const cmd = input?.command;
    if (typeof cmd === "string" && bashMutatesOfficeFile(cmd)) {
      return Promise.resolve(denyWithOfficeMessage());
    }
  }
  if (host === "excel" && toolName === "mcp__thepexcel-excel__excel_vba") {
    return Promise.resolve({
      behavior: "deny",
      message:
        "VBA execution is disabled in this Excel session. Use the mcp__office__ tools and excel_bash, or explain that an explicitly requested macro cannot be installed while VBA is disabled.",
    });
  }
  if (
    host === "excel" &&
    toolName === "mcp__thepexcel-excel__excel_range" &&
    input?.action === "write_py"
  ) {
    return Promise.resolve({
      behavior: "deny",
      message:
        "Python in Excel is not available in this session. Use mcp__office__excel_bash for computation and mcp__office__excel_set_cell_range for workbook writes.",
    });
  }
  // Everything else: auto-approve.
  return Promise.resolve({ behavior: "allow", updatedInput: input ?? {} });
}

// ---------------------------------------------------------------------------
// Async iterable that pulls user messages from the bridge and yields them
// to the Agent SDK in the SDKUserMessage shape.
//
// We also prepend a context header to each turn so the agent always knows the
// active doc and selection without having to call office_get_doc_info first.
// ---------------------------------------------------------------------------
async function* userMessageStream(key, session) {
  while (true) {
    let msg;
    try {
      msg = await bridge.nextUserMessage(key, { signal: session?.abortController.signal });
    } catch {
      // Bridge rejected the waiter — session was aborted. Exit cleanly so
      // the underlying query() iterator can shut down without a stray error.
      return;
    }
    if (session && !isCurrentSession(session)) return;
    const { text, context } = msg;
    // The Agent SDK only treats a turn as a slash command (built-in or a
    // custom .claude/commands/*.md) when the message *starts with* "/".
    // Our per-turn context header (Host:/Doc:/Selection:) would otherwise
    // push the "/" off the front and the command would be sent to the
    // model as prose. So for a slash command, send it bare (leading
    // whitespace trimmed so detection works) and skip the header — the
    // command template is self-contained; it can call office_* tools if
    // it needs doc context.
    const trimmed = typeof text === "string" ? text.trimStart() : text;
    const isSlashCommand = typeof trimmed === "string" && trimmed.startsWith("/");
    const automaticContext = isSlashCommand ? "" : await autoContext(key, session, context);
    const header = [renderContextHeader(context), automaticContext]
      .filter(Boolean)
      .join("\n\n");
    const content = isSlashCommand ? trimmed : header ? `${header}\n\n${text}` : text;
    // Per-turn tracking so a slash command that produces no assistant text
    // or tool call (terminal-only built-ins like /help, /context, /clear)
    // doesn't look like a dead chat. Reset at the start of every turn.
    if (session) {
      session.slashCommandPending = isSlashCommand;
      session.turnProducedOutput = false;
      session.turnOpen = true;
      session.approveRestOfTurn = false;
    }
    yield {
      type: "user",
      message: { role: "user", content },
      parent_tool_use_id: null,
    };
  }
}

// Workbook overview, the selection with nearby rows, and recent workbook edits
// since the last turn, read by the pane (pi-for-excel readers). The overview is
// repeated only when it changed within this session. Never blocks a turn for long.
const AUTO_CONTEXT_TIMEOUT_MS = 4000;

async function autoContext(key, session, ctx) {
  if (ctx.host !== "excel") return "";
  let snapshot;
  try {
    const signals = [AbortSignal.timeout(AUTO_CONTEXT_TIMEOUT_MS)];
    if (session?.abortController) signals.push(session.abortController.signal);
    snapshot = await bridge.callTaskpaneTool(
      "excel_context_snapshot",
      { selectionAddress: ctx.selection?.address ?? null },
      key,
      { signal: AbortSignal.any(signals) },
    );
  } catch (err) {
    diag(`[auto-context] skipped: ${err?.message ?? err}`);
    return "";
  }
  const sections = [];
  if (snapshot?.workbook && snapshot.workbook !== session?.lastWorkbookContext) {
    sections.push(snapshot.workbook);
    if (session) session.lastWorkbookContext = snapshot.workbook;
  }
  if (snapshot?.selection) sections.push(snapshot.selection);
  if (snapshot?.changes) sections.push(snapshot.changes);
  return sections.length ? `[Auto-context]\n${sections.join("\n\n")}` : "";
}

function renderContextHeader(ctx) {
  const parts = [];
  if (ctx.host === "excel") parts.push("Host: Excel");
  if (ctx.activeDoc) parts.push(`Doc: ${ctx.activeDoc}`);
  if (ctx.selection) {
    const s = ctx.selection;
    if (s.text) {
      const preview = s.text.length > 80 ? s.text.slice(0, 77) + "..." : s.text;
      parts.push(`Selection: "${preview}"`);
    } else if (s.para_id) {
      parts.push(`Cursor in paragraph ${s.para_id}`);
    }
  }
  // Only surface track-changes mode when it's NOT the default ("always"). The
  // system prompt says "always" by default; only deviation needs signaling.
  if (ctx.trackChangesMode && ctx.trackChangesMode !== "always") {
    parts.push(`Track changes: ${ctx.trackChangesMode}`);
  }
  return parts.length ? `[${parts.join(" · ")}]` : "";
}

// Which model provider the agent talks to, for user-facing error messages.
function providerLabel() {
  const base = process.env.ANTHROPIC_BASE_URL;
  if (!base) return "Claude";
  try {
    return new URL(base).host;
  } catch {
    return base;
  }
}

function providerRecoveryHint() {
  return process.env.ANTHROPIC_BASE_URL
    ? "Check the API key, balance and model names in .env, then restart the daemon."
    : "Wait for your Claude limit to reset, or set ANTHROPIC_API_KEY to use an API key.";
}

// ---------------------------------------------------------------------------
// Session management. Each open document (paneKey) runs its own
// independent query() loop with its own cwd; switching workspace or
// reloading config restarts only that pane's loop. Histories live in
// ~/.claude/projects/<hash>/*.jsonl per the SDK's normal persistence.
// `sessions` / `workspaceByKey` are declared at the top of the module
// (above createBridge) so bridge handlers that fire during top-level
// awaits (e.g. preflightHttpMcpServers, which can block for 5s) don't
// hit a TDZ on those bindings.
// ---------------------------------------------------------------------------

async function startSessionForFolder(
  cwd,
  resumeSessionId = null,
  { key = null, host = null, replay = true, generation = sessionGenerationByKey.get(key) } = {},
) {
  if (generation !== sessionGenerationByKey.get(key)) return null;
  // Only when actually superseding a live session for THIS pane: abort it
  // and drain its queue (same-pane restart: workspace switch, config
  // reload, post-Stop/stream-end resume). No other pane's loop is ever
  // involved. On a FRESH start there is no prior session and the first
  // user message has already been enqueued for this pane (it's what
  // triggered the lazy start) — clearing here would drop it, leaving
  // userMessageStream awaiting forever and the SDK with no first input
  // (no init, taskpane stuck on "Working…").
  const prior = sessionFor(key);
  if (prior) {
    prior.abortController.abort();
    bridge.clearUserMessages(key);
  }

  const abortController = new AbortController();
  const session = {
    key,
    cwd,
    sessionId: resumeSessionId,
    abortController,
    settled: false,
    host,
    generation,
  };
  sessions.set(key, session);
  workspaceByKey.set(key, cwd);

  // Register only this host's tool family, routed to THIS pane. A session
  // is created lazily on the first user message (onUserMessage →
  // ensureLoopForMessage), so `host` is normally "word" or "excel".
  let officeMcp;

  // Recents bookkeeping only; a failure here must not block the session.
  try {
    await touchFolder(cwd);
  } catch (err) {
    console.warn(`[daemon] could not record workspace ${cwd}: ${err?.message ?? err}`);
  }
  if (!isCurrentSession(session)) {
    session.settled = true;
    return null;
  }
  console.log(
    `[daemon] Starting session for ${cwd}` +
      (resumeSessionId ? ` (resuming ${resumeSessionId.slice(0, 8)}…)` : " (new session)"),
  );
  bridge.sendAssistantEvent({ event: "cwd_changed", cwd, resumed: !!resumeSessionId }, key);
  // Structured readiness signal to the Electron shell over the IPC
  // channel — the session loop is up. Lets main.mjs flip the tray to
  // "Ready" without sniffing our stdout for a log substring. No-op when
  // run via `npm run dev` (no IPC channel).
  if (process.send) {
    try {
      process.send({ type: "daemon_ready", cwd });
    } catch {
      /* channel gone */
    }
  }
  // Replay this host's transcript for the new workspace so its panel
  // reflects the workspace you just switched to (not the previous chat).
  // Skipped when this start was triggered by the user's own message
  // (replay:false): the pane already shows that message, and a fresh
  // session's empty replay would erase it.
  if (replay) sendTranscriptReplayTo(key, host, cwd).catch(() => {});

  // Re-read the drafting setup append fresh each session start.
  let append;
  try {
    append = await buildSystemPromptAppend();
    if (!isCurrentSession(session)) {
      session.settled = true;
      return null;
    }
    officeMcp = createOfficeBridgeMcp(bridge, host, key, {
      signal: abortController.signal,
    });
  } catch (err) {
    // Without cleanup the half-built session stays registered as live, so
    // ensureLoopForMessage would treat it as a consumer and every later
    // message would wait forever. Drop it and end the pending turn instead.
    if (!isCurrentSession(session)) return null;
    retireCurrentSession(session, { clearMessages: true });
    const reason = `Could not start the agent: ${err?.message ?? err}`;
    console.error(`[daemon] ${reason}`);
    bridge.sendAssistantEvent({ event: "error", error: reason }, key);
    throw err;
  }

  // Tracked per turn (not per loop): the loop serves many turns, and a later
  // turn can end without a `result` even after earlier turns succeeded. The
  // SDK normally ends each turn with a `result`; on a usage-limit / quota hit
  // (and some transport failures) the stream just ends with no result and no
  // thrown error — leaving the taskpane pinned to "Working…".
  session.turnOpen = false;
  session.sawAnyResult = false;
  // Best-effort usage-limit detection from the SDK CLI's stderr. The exact
  // phrasing varies by SDK version and limit kind (per-minute / daily /
  // weekly); match broadly.
  let rateLimitHint = null;
  const RATE_LIMIT_RE =
    /(usage limit|rate limit|daily limit|weekly limit|quota|too many requests|429|limit reached|limit will reset|resets? at|upgrade to|out of (?:credits|quota))/i;

  // Fire-and-forget; index.mjs keeps running while the agent loop iterates.
  (async () => {
    try {
      for await (const msg of query({
        prompt: userMessageStream(key, session),
        options: {
          cwd,
          systemPrompt: {
            type: "preset",
            preset: "claude_code",
            append,
            // Re-read the prompt files on every session start, including resumes.
            snapshot: false,
          },
          mcpServers: { ...userMcpServers, office: officeMcp },
          plugins: agentPlugins,
          ...(agentConfig.skills ? { skills: agentConfig.skills } : {}),
          tools: agentConfig.builtinTools,
          disallowedTools: agentConfig.disallowedTools,
          settingSources: agentConfig.settingSources,
          canUseTool: (toolName, input) => permissionFor(key, session, host, toolName, input),
          includePartialMessages: true,
          // User-chosen model (composer dropdown); always explicit.
          model: modelArgFor(key),
          abortController,
          // Surface the SDK CLI's stderr (MCP connect failures, internal
          // warnings, etc.) in our daemon log. Also sniff it for
          // usage-limit phrasing so we can show the user a clear message
          // instead of a silently stuck "Working…".
          stderr: (data) => {
            for (const line of String(data).split(/\r?\n/)) {
              if (!line.trim()) continue;
              console.error(`[sdk] ${line}`);
              if (!rateLimitHint && RATE_LIMIT_RE.test(line)) rateLimitHint = line.trim();
            }
          },
          ...(resumeSessionId ? { resume: resumeSessionId } : {}),
        },
      })) {
        if (!isCurrentSession(session)) break;
        if (msg.type === "result") {
          session.turnOpen = false;
          session.sawAnyResult = true;
        }
        handleAgentMessage(msg, session);
      }
      // Include accepted input that the SDK has not consumed yet. Report
      // one failure terminal event; a subsequent message starts a fresh
      // consumer. A result from an earlier turn cannot hide this failure.
      const turnUnfinished =
        session.turnOpen || bridge.hasPendingUserMessages(key) || !session.sawAnyResult;
      if (isCurrentSession(session) && turnUnfinished) {
        const friendly = rateLimitHint
          ? `${providerLabel()} usage limit reached. ${rateLimitHint}`
          : `The agent stopped unexpectedly — usually a usage limit, quota or error at the model provider (${providerLabel()}). ${providerRecoveryHint()}`;
        // Retire before notifying the UI. A user can retry immediately from
        // an event handler; that retry must see no old live session.
        retireCurrentSession(session, { clearMessages: true });
        bridge.sendAssistantEvent(
          { event: "error", subtype: "stream_ended", error: friendly },
          key,
        );
      }
    } catch (err) {
      if (sessionFor(key) === session) {
        retireCurrentSession(session, { clearMessages: true });
        console.error("[daemon] Agent loop crashed:", err);
        // Detect auth failures and surface them as a distinct event so the
        // taskpane can show a recoverable banner ("sign in to Claude Code")
        // instead of just dumping the SDK's raw error. Matched generously:
        // SDK error messages have varied across versions.
        const msgText = String(err?.message ?? err);
        const isAuth =
          /\b(authentication|unauthorized|credential|api[- ]?key|sign[- ]?in|401)\b/i.test(
            msgText,
          ) || /OAUTH/i.test(msgText);
        if (isAuth) {
          bridge.sendAssistantEvent({ event: "auth_error", error: msgText }, key);
        } else {
          bridge.sendAssistantEvent({ event: "error", error: msgText }, key);
        }
      }
    } finally {
      session.settled = true;
      if (sessionFor(key) === session) sessions.delete(key);
      // The SDK can end its output while still awaiting input. Release
      // only this loop's waiter, including on normal completion.
      abortController.abort();
    }
  })();

  return session;
}

async function switchFolder(rawCwd, key, host = null) {
  const cwd = resolve(rawCwd);
  // Validate the path is a directory.
  const s = await stat(cwd);
  if (!s.isDirectory()) throw new Error(`Not a directory: ${cwd}`);
  // Switch ONLY the requesting pane to the target folder. Its workbook
  // conversation remains the same; every other pane stays untouched.
  const resumeId = host ? await getSessionId(host, documentKeyForPane(key)) : null;
  cancelPaneSession(key);
  workspaceByKey.set(key, cwd);
  bridge.sendAssistantEvent({ event: "turn_complete", interrupted: true }, key);
  bridge.sendAssistantEvent({ event: "cwd_changed", cwd, resumed: !!resumeId }, key);
  // A folder switch only changes context/history. Start the SDK when the
  // next user message arrives, just as on an initial pane connection.
  sendTranscriptReplayTo(key, host, cwd).catch(() => {});
  return cwd;
}

// Re-launch one pane's loop (same cwd, resuming via session_id) so that
// changes to CLAUDE.md, the drafting setup, or other config loaded at
// session-init take effect without losing conversation history. No-op if
// that pane has no live loop.
async function restartSession(key, host, { reason = "config_changed" } = {}) {
  const s = sessionFor(key);
  if (!s) return;
  const { cwd, sessionId } = s;
  cancelPaneSession(key);
  console.log(`[daemon] Restarting session for ${cwd} (reason: ${reason})`);
  bridge.sendAssistantEvent({ event: "config_reloaded", reason }, key);
  // Funnel through the serialized per-key queue (not a direct
  // startSessionForFolder) so a config/model restart coalesces with any
  // concurrent post-Stop / first-message start instead of racing it.
  scheduleSessionStart(cwd, sessionId, key, host, reason, { replay: false });
}

function handleAgentMessage(msg, session) {
  switch (msg.type) {
    case "system": {
      if (msg.subtype === "init") {
        console.log(
          `[agent] init session ${msg.session_id} (model: ${msg.model}; plugins: ${(msg.plugins ?? []).map((p) => p.name).join(", ") || "none"}; skills: ${msg.skills?.length ?? 0})`,
        );
        try {
          const tn = Array.isArray(msg.tools) ? msg.tools : Object.keys(msg.tools ?? {});
          diag(
            `init tools (${tn.length}):`,
            tn.filter((t) => /office|excel|mcp__/.test(String(t))).join(", ") ||
              "(no office/excel/mcp tools in init list!)",
          );
        } catch (e) {
          diag("init tools introspection failed:", e?.message);
        }
        bridge.sendAssistantEvent(
          {
            event: "session_init",
            session_id: msg.session_id,
            model: msg.model,
          },
          session?.key,
        );
        // Record this session_id for THIS workbook so a reconnect resumes
        // it independently of the workspace folder.
        if (session && msg.session_id && msg.session_id !== session.sessionId) {
          session.sessionId = msg.session_id;
          saveSessionId(
            session.host,
            documentKeyForPane(session.key),
            session.cwd,
            msg.session_id,
          ).catch((err) => console.warn("[daemon] Could not save session id:", err.message));
        }
      } else {
        console.log(`[agent] system/${msg.subtype}`);
      }
      break;
    }
    case "stream_event": {
      // includePartialMessages stream events. Forward text_delta to the
      // taskpane as assistant_text. Other event types (content_block_start,
      // content_block_stop, message_start/stop) we currently ignore.
      const delta = msg.event?.delta;
      if (delta?.type === "text_delta" && typeof delta.text === "string" && delta.text.length > 0) {
        if (session) session.turnProducedOutput = true;
        bridge.sendAssistantText(delta.text, session?.key);
      }
      break;
    }
    case "assistant": {
      // The complete assistant message arrives after streaming. We skip
      // text blocks (already streamed as deltas) and only forward tool_use
      // announces — those are atomic and not streamed.
      const blocks = msg.message?.content ?? [];
      for (const block of blocks) {
        if (block.type === "tool_use") {
          diag("model called tool:", block.name);
          if (session) session.turnProducedOutput = true;
          bridge.sendAssistantEvent(
            {
              event: "tool_use_announce",
              tool: block.name,
              input: block.input,
            },
            session?.key,
          );
        }
      }
      break;
    }
    case "result": {
      console.log(
        `[agent] turn complete (${msg.subtype}) slash=${!!session?.slashCommandPending} output=${!!session?.turnProducedOutput}`,
      );
      // A slash command that emitted neither assistant text nor a tool call
      // (terminal-only built-ins: /help, /context, /clear, …) would
      // otherwise complete silently and look broken. Surface a note.
      if (session?.slashCommandPending && !session.turnProducedOutput) {
        console.log("[agent] slash command produced no output — sending info note");
        bridge.sendAssistantEvent(
          {
            event: "info",
            message:
              "Command ran, but produced no chat output. Some built-in commands (e.g. /help, /context, /clear) are terminal-only and don't return anything here — custom .claude/commands and content-producing commands do.",
          },
          session?.key,
        );
      }
      if (session) {
        session.slashCommandPending = false;
        session.turnProducedOutput = false;
      }
      bridge.sendAssistantEvent(
        {
          event: "turn_complete",
          subtype: msg.subtype,
          ...(msg.subtype !== "success"
            ? {
                error:
                  msg.errors?.join("\n") ||
                  `The agent ended this request with ${msg.subtype || "an error"}.`,
              }
            : {}),
          usage: msg.usage,
          num_turns: msg.num_turns,
          total_cost_usd: msg.total_cost_usd,
        },
        session?.key,
      );
      break;
    }
    case "user": {
      // tool_result messages — we don't forward them; the bridge handles them.
      break;
    }
    default:
      // Other event types (api_retry, hook events, etc.) — not surfaced.
      break;
  }
}

// ---------------------------------------------------------------------------
// Kick off.
//
// No eager session: an agent session needs a host, and the host is only
// known once a taskpane connects and says hello (→ ensureSessionFor
// ActivePane). Starting both-tool sessions pre-hello is exactly what
// caused the per-host tool/transcript churn. The bridge buffers any
// user_message until a session's loop consumes it, so nothing is lost.
//
// Tell the Electron shell we're up now (servers listening) so the tray
// flips to "Ready" without waiting for a pane — independent of, and
// idempotent with, the daemon_ready that startSessionForFolder also
// emits on the first real session.
// ---------------------------------------------------------------------------
console.log(`[daemon] Ready; waiting for a taskpane. Default workspace: ${matterFolder}`);
if (process.send) {
  try {
    process.send({ type: "daemon_ready" });
  } catch {
    /* no IPC channel (npm run dev) */
  }
}

// Keep process alive even when nothing is happening.
process.on("SIGINT", () => {
  console.log("\n[daemon] Shutting down");
  process.exit(0);
});
