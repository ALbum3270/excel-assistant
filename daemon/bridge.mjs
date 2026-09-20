import { WebSocketServer } from "ws";
import { randomUUID } from "node:crypto";
import { createWorkbookExecution, canonicalWorkbookId } from "./workbook-execution.mjs";
import { needsApproval } from "./approval.mjs";

// 60s, not 30s. Some Office.js ops legitimately run long on big docs —
// office_clear_highlights / a wide office_replace_section over a large
// document, batched office_highlight, a whole-sheet excel_read_range.
// 30s was tripping on real-sized documents. (Per-tool timeouts would be
// finer-grained but aren't worth the complexity yet.)
const TOOL_TIMEOUT_MS = 60_000;

function normalizeHost(h) {
  return h === "word" || h === "excel" ? h : null;
}

// A pane is keyed by (host, document), not just host: a user can have two
// Word docs (or two workbooks) open at once, each with its own task pane.
// Keying by host alone made them fight over a single slot — the same
// connect/disconnect ping-pong the host split originally fixed, just one
// level down. The host stays embedded in the key so the tool family and
// per-host behavior are still derivable from it.
const PANE_SEP = "\u0000";
function makePaneKey(host, doc, paneId) {
  let d;
  if (doc && String(doc).trim()) {
    d = String(doc);
  } else if (paneId && String(paneId).trim()) {
    // Unsaved / cloud doc with no filesystem path: key by the taskpane's
    // stable per-instance id (sent in hello) so a reconnect REUSES this
    // pane's key instead of minting a fresh random one every ~1.5s while
    // the daemon is down — which leaked per-pane state unbounded.
    d = "anon:" + String(paneId);
  } else {
    d = "anon:" + randomUUID();
  }
  return host + PANE_SEP + d;
}
function hostOfKey(key) {
  return typeof key === "string" ? key.split(PANE_SEP, 1)[0] : null;
}

export function createBridge({
  port,
  extraHandlers = {},
  token,
  allowedOrigins = [],
  onHello,
  onUserMessage,
  onClose,
  onListening,
  onListenError,
  execution = createWorkbookExecution(),
}) {
  if (!token) throw new Error("createBridge requires a token");
  const wss = new WebSocketServer({
    port,
    // Bind to loopback only. The `ws` library defaults to 0.0.0.0 when
    // given just `port`, which exposes the bridge to the local network.
    // This is a local-only IPC channel (taskpane <-> daemon on the same
    // machine), so it should never be reachable off-host regardless of
    // the token + origin gates below. Matches the HTTP server in
    // index.mjs, which already binds 127.0.0.1.
    host: "127.0.0.1",
    // First gate: only allow upgrades from known origins (the taskpane's
    // origin = our own HTTP server's). Browsers honor the Origin header on
    // WebSocket upgrades; rejecting unknown origins blocks malicious local
    // web pages from driving the agent.
    verifyClient: (info, cb) => {
      const origin = info.req.headers.origin || "";
      if (allowedOrigins.length === 0 || allowedOrigins.includes(origin)) {
        cb(true);
      } else {
        console.warn(`[bridge] Rejecting WS upgrade from origin: ${origin || "(none)"}`);
        cb(false, 403, "Forbidden origin");
      }
    },
  });

  // One pane per (host, document). A connection is "unbound" until its
  // hello names a host + active doc; a reconnect with the SAME key
  // replaces the prior pane (Office reloaded that document's panel), any
  // DIFFERENT key — different host OR a different document of the same
  // host — coexists.
  const panes = new Map(); // paneKey -> { ws, context }

  function emptyContext(host = null, activeDoc = null) {
    return { activeDoc, selection: null, trackChangesMode: "always", host };
  }

  function paneWs(key) {
    const p = panes.get(key);
    return p && p.ws.readyState === p.ws.OPEN ? p.ws : null;
  }

  // Tool calls awaiting their tool_result. Keyed by id; each remembers the
  // ws it was dispatched on so a stale pane's disconnect only rejects its
  // own in-flight calls.
  const pendingTools = new Map();

  // Per-pane queue of pending user messages. Waiters are {resolve, reject}
  // so we can reject pending awaits when that pane's session is aborted
  // (otherwise the suspended generator from the aborted session would
  // consume the next user message, starving the new session).
  const queues = new Map(); // paneKey -> { queue: [], waiters: [] }
  function queueFor(key) {
    let q = queues.get(key);
    if (!q) {
      q = { queue: [], waiters: [] };
      queues.set(key, q);
    }
    return q;
  }

  function pushUserMessage(text, key) {
    const ctx = panes.get(key)?.context ?? emptyContext(hostOfKey(key));
    const payload = { text, context: { ...ctx, host: ctx.host ?? hostOfKey(key) } };
    const q = queueFor(key);
    if (q.waiters.length > 0) {
      q.waiters.shift().resolve(payload);
    } else {
      q.queue.push(payload);
    }
  }

  function nextUserMessage(key, { signal } = {}) {
    const aborted = () => Object.assign(new Error("Session aborted"), { name: "AbortError" });
    if (signal?.aborted) return Promise.reject(aborted());
    const q = queueFor(key);
    if (q.queue.length > 0) return Promise.resolve(q.queue.shift());
    return new Promise((resolve, reject) => {
      const cleanup = () => signal?.removeEventListener("abort", onAbort);
      const waiter = {
        resolve: (value) => {
          cleanup();
          resolve(value);
        },
        reject: (error) => {
          cleanup();
          reject(error);
        },
      };
      const onAbort = () => {
        const index = q.waiters.indexOf(waiter);
        if (index !== -1) q.waiters.splice(index, 1);
        waiter.reject(aborted());
      };
      q.waiters.push(waiter);
      signal?.addEventListener("abort", onAbort, { once: true });
    });
  }

  // Reject all pending waiters and drop queued messages for ONE pane.
  // Called before that pane's session swaps so the new session starts
  // clean with no zombie consumer — and, crucially, without touching any
  // OTHER pane's queue (a message just enqueued for the pane we're
  // switching TO must survive).
  function clearUserMessages(key) {
    const q = queues.get(key);
    if (!q) return;
    while (q.waiters.length > 0) {
      q.waiters.shift().reject(new Error("Session aborted"));
    }
    q.queue.length = 0;
  }

  // Merge inbound doc/selection/track-changes fields into a pane's
  // context. `hello` resets doc+selection (a fresh taskpane connection —
  // absent fields mean null), while `context_update` patches (only fields
  // present on the message change). track-changes is "set only when
  // valid" in both modes. host is fixed at bind time, never re-derived.
  function mergeContext(ctx, msg, { reset = false } = {}) {
    if (reset) {
      ctx.activeDoc = msg.active_doc ?? null;
      ctx.selection = msg.selection ?? null;
    } else {
      if (msg.active_doc !== undefined) ctx.activeDoc = msg.active_doc;
      if (msg.selection !== undefined) ctx.selection = msg.selection;
    }
    if (typeof msg.track_changes_mode === "string") {
      ctx.trackChangesMode = msg.track_changes_mode;
    }
  }

  function send(obj, key) {
    const ws = paneWs(key);
    if (!ws) {
      console.warn(`[bridge] No pane for ${key ?? "?"}; dropping`, obj.type);
      return;
    }
    ws.send(JSON.stringify(obj));
  }

  function workbookIdFor(key) {
    return canonicalWorkbookId(panes.get(key)?.context.activeDoc || key);
  }

  async function callTaskpaneTool(
    name,
    args,
    key = null,
    { signal, expectedRevision = args?.expected_revision } = {},
  ) {
    const workbookId = workbookIdFor(key);
    const write =
      needsApproval(`mcp__office__${name}`, args) &&
      (name !== "excel_workbook_history" || args?.action === "restore");
    const { expected_revision, ...toolArgs } = args ?? {};
    const completed = await execution.run(
      workbookId,
      { write, signal, expectedRevision, toolName: name, timeoutMs: TOOL_TIMEOUT_MS },
      (context) => {
        if (workbookIdFor(key) !== workbookId)
          throw Object.assign(
            new Error("The pane changed workbooks before dispatch; read its current context."),
            { commitStatus: "not_committed" },
          );
        return dispatchTaskpaneTool(name, toolArgs, key, { ...context, workbookId });
      },
    );
    return {
      ...(completed.result && typeof completed.result === "object"
        ? completed.result
        : { result: completed.result }),
      workbookRevision: completed.revision,
      ...(completed.uncertain ? { workbookUncertain: true } : {}),
    };
  }

  async function dispatchTaskpaneTool(
    name,
    args,
    key,
    { signal, revision, opId, workbookId, write },
  ) {
    const ws = paneWs(key);
    if (!ws) {
      throw Object.assign(new Error(`Cannot call tool ${name}: no taskpane for ${key ?? "?"}`), {
        commitStatus: "not_committed",
      });
    }
    if (signal?.aborted) {
      throw Object.assign(new Error(`Tool ${name} cancelled before dispatch`), {
        name: "AbortError",
        code: "TOOL_CANCELLED",
        commitStatus: "not_committed",
      });
    }
    const id = opId;
    // Capture the WS this call is dispatched on. On close we only reject
    // pending calls belonging to that specific WS — so a stale pane
    // disconnecting after a fresh one is active doesn't kill live work.
    const ownerWs = ws;
    const promise = new Promise((resolve, reject) => {
      const cancelPaneCall = () => {
        if (ownerWs.readyState === ownerWs.OPEN) {
          try {
            ownerWs.send(JSON.stringify({ type: "tool_cancel", id }));
          } catch {}
        }
      };
      const cleanup = () => signal?.removeEventListener("abort", onAbort);
      const onAbort = () => {
        cancelPaneCall();
        // The caller stops waiting through execution.run, but the shared queue
        // stays occupied until the pane confirms the handler actually ended.
      };
      pendingTools.set(id, {
        ws: ownerWs,
        key,
        name,
        resolve,
        reject,
        cleanup,
        workbookId,
        write,
        runtimeId: panes.get(key)?.runtimeId,
      });
      signal?.addEventListener("abort", onAbort, { once: true });
      if (signal?.aborted) onAbort();
    });
    ownerWs.send(
      JSON.stringify({ type: "tool_call", id, name, args, workbook_revision: revision }),
    );
    return await promise;
  }

  function sendAssistantText(delta, key) {
    send({ type: "assistant_text", delta }, key);
  }

  function sendAssistantEvent(event, key) {
    send({ type: "assistant_event", ...event }, key);
  }

  wss.on("connection", (ws, req) => {
    console.log(
      `[bridge] WS connection from ${req.socket.remoteAddress} (origin: ${req.headers.origin || "(none)"})`,
    );

    // Second gate: first message must be a hello with the correct token.
    // Until that arrives we don't trust the connection — any other message
    // type is rejected and the socket is closed. The taskpane fetches the
    // token from /bridge-token over the same-origin HTTP server before
    // opening the WS. A connection is bound to exactly one (host,
    // document) pane key for its lifetime.
    let authed = false;
    let boundKey = null;
    let boundHost = null;

    // True iff this ws is still the live pane for the key it bound to.
    // A superseded (same-key reconnect) ws stays open just long enough
    // to be force-closed; drop any stale frames it sends meanwhile.
    function isLivePane() {
      return authed && panes.get(boundKey)?.ws === ws;
    }

    ws.on("message", (raw) => {
      let msg;
      try {
        msg = JSON.parse(raw.toString());
      } catch (e) {
        console.error("[bridge] Bad JSON from taskpane:", e.message);
        return;
      }

      if (!authed) {
        if (msg.type !== "hello") {
          console.warn(`[bridge] Pre-auth message ${msg.type}; closing`);
          ws.close(4001, "Authenticate first");
          return;
        }
        if (msg.token !== token) {
          console.warn("[bridge] hello with invalid token; closing");
          ws.close(4002, "Invalid token");
          return;
        }
        const host = normalizeHost(msg.host);
        if (!host) {
          console.warn(`[bridge] hello without a valid host (${msg.host}); closing`);
          ws.close(4005, "hello must name host word|excel");
          return;
        }
        boundHost = host;
        boundKey = makePaneKey(host, msg.active_doc, msg.pane_id);
        // Replace only a prior pane with the SAME key (same document's
        // panel reloaded / reopened). A different document — even of the
        // same host — gets its own pane and coexists.
        const prior = panes.get(boundKey);
        if (prior && prior.ws !== ws && prior.ws.readyState === prior.ws.OPEN) {
          console.log(`[bridge] New pane authed for ${boundKey}; closing prior`);
          try {
            prior.ws.close(4003, "Replaced by new pane");
          } catch {}
        }
        authed = true;
        panes.set(boundKey, {
          ws,
          context: emptyContext(host, msg.active_doc ?? null),
          runtimeId: msg.runtime_id,
        });
      } else if (!isLivePane() && msg.type !== "tool_result" && msg.type !== "tool_settled") {
        // Superseded by a newer same-key pane; its close is in flight.
        // Drop stale frames rather than letting them mutate context or
        // queue a user message for a pane that's going away.
        console.warn(`[bridge] Post-auth ${msg.type} from superseded pane ${boundKey}; ignoring`);
        return;
      }

      const pane = panes.get(boundKey);

      switch (msg.type) {
        case "hello": {
          mergeContext(pane.context, msg, { reset: true });
          ws.send(
            JSON.stringify({
              type: "welcome",
              session_id: randomUUID(),
              server_version: "0.1.0",
            }),
          );
          console.log(
            `[bridge] hello received; host: ${boundHost}; active doc: ${pane.context.activeDoc}`,
          );
          // Let the daemon replay this pane's transcript (and anything
          // else it wants on connect). Fire-and-forget; never block.
          if (onHello) {
            Promise.resolve(onHello(boundKey, boundHost, pane.context.activeDoc)).catch((err) =>
              console.warn("[bridge] onHello failed:", err?.message ?? err),
            );
          }
          break;
        }
        case "user_message": {
          console.log(`[bridge] user_message (${boundHost}):`, msg.text);
          // The taskpane captures selection immediately before submit. Merge
          // that per-turn snapshot before queueing so the payload cannot use
          // an older debounced context_update.
          if (Object.prototype.hasOwnProperty.call(msg, "selection")) {
            mergeContext(pane.context, { selection: msg.selection });
          }
          // Give the daemon a chance to (lazily) start/swap the agent
          // loop for this pane BEFORE the message is queued, so the right
          // loop consumes it. Fire-and-forget; ordering holds because
          // session start is deferred to setImmediate, after this push.
          if (onUserMessage) {
            Promise.resolve(onUserMessage(boundKey, boundHost, pane.context.activeDoc)).catch(
              (err) => console.warn("[bridge] onUserMessage failed:", err?.message ?? err),
            );
          }
          pushUserMessage(msg.text, boundKey);
          break;
        }
        case "context_update": {
          mergeContext(pane.context, msg);
          break;
        }
        case "tool_settled":
        case "tool_result": {
          const pending = pendingTools.get(msg.id);
          if (!pending) {
            if (msg.type === "tool_result")
              console.warn("[bridge] tool_result for unknown id:", msg.id);
            return;
          }
          const resumedRuntime =
            pending.runtimeId && pending.runtimeId === panes.get(boundKey)?.runtimeId;
          if ((pending.ws !== ws && !resumedRuntime) || pending.key !== boundKey) {
            console.warn("[bridge] tool_result owner mismatch for id:", msg.id);
            return;
          }
          pendingTools.delete(msg.id);
          pending.cleanup?.();
          execution.settled(pending.workbookId, msg.id);
          if (msg.type === "tool_settled") {
            pending.reject(
              Object.assign(
                new Error(
                  "The cancelled Excel operation finished; inspect the workbook before retrying.",
                ),
                { code: "TOOL_CANCELLED", commitStatus: "unknown", executionSettled: true },
              ),
            );
          } else if (msg.ok) {
            pending.resolve(msg.result);
          } else {
            const detail =
              msg.error && typeof msg.error === "object"
                ? msg.error
                : { message: msg.error ?? "Unknown tool error" };
            pending.reject(
              Object.assign(new Error(detail.message ?? "Unknown tool error"), {
                ...(detail.code ? { code: detail.code } : {}),
                ...(detail.commitStatus ? { commitStatus: detail.commitStatus } : {}),
                ...(detail.recovery ? { recovery: detail.recovery } : {}),
                executionSettled: true,
              }),
            );
          }
          break;
        }
        case "ping": {
          ws.send(JSON.stringify({ type: "pong" }));
          break;
        }
        default: {
          // Custom handlers registered by the daemon (e.g. settings /
          // refs). They receive the message, a `reply(obj)` shortcut that
          // sends back over this same WS, and the bound pane key + host so
          // per-pane actions (set_cwd, stop_agent, …) act on the right
          // document.
          const handler = extraHandlers[msg.type];
          if (handler) {
            const reply = (obj) => ws.send(JSON.stringify(obj));
            Promise.resolve()
              .then(() => handler(msg, reply, boundKey, boundHost))
              .catch((err) => {
                console.error(`[bridge] handler for ${msg.type} threw:`, err);
                reply({
                  type: msg.type + "_result",
                  ok: false,
                  error: err.message ?? String(err),
                  request_id: msg.request_id,
                });
              });
          } else {
            console.warn("[bridge] Unknown message type:", msg.type);
          }
        }
      }
    });

    ws.on("close", () => {
      console.log(`[bridge] Taskpane disconnected${boundKey ? ` (${boundHost})` : ""}`);
      // Only clear the pane mapping if THIS ws still owns it (a superseded
      // ws closing must not evict the fresh same-key pane).
      if (boundKey && panes.get(boundKey)?.ws === ws) {
        panes.delete(boundKey);
        // Let the daemon prune per-key state for a pane that's gone.
        // Deliberately NOT touching this pane's queue or its live
        // session here: a transient WS blip auto-reconnects with the
        // SAME key (stable now), and a suspended userMessageStream is
        // awaiting the existing queue — dropping it would resurface the
        // "stuck on Working…" bug. The daemon's onClose only prunes the
        // lightweight Maps for keys with no live session.
        if (onClose) {
          try {
            onClose(boundKey, boundHost);
          } catch (err) {
            console.warn("[bridge] onClose failed:", err?.message ?? err);
          }
        }
      }
      // Reject only the pending tool calls dispatched on THIS ws. A
      // different (newer) WS may have its own in-flight calls; leave them.
      for (const [id, p] of pendingTools) {
        if (p.ws === ws) {
          p.cleanup?.();
          p.reject(
            Object.assign(
              new Error(
                "Taskpane disconnected while a tool call was in flight; workbook state may be unknown",
              ),
              {
                code: "TASKPANE_DISCONNECTED",
                commitStatus: "unknown",
              },
            ),
          );
          if (p.write)
            p.orphaned = true; // A late completion can safely clear the write block.
          else pendingTools.delete(id);
        }
      }
    });

    ws.on("error", (err) => {
      console.error("[bridge] WS error:", err.message);
    });
  });

  // WebSocketServer binds asynchronously, so this has to wait for the event:
  // logging it at construction time announced a listener that a port clash
  // then quietly denied.
  wss.on("listening", () => {
    console.log(
      `[bridge] WebSocket server listening on ws://127.0.0.1:${port} (origin allowlist: ${allowedOrigins.join(", ") || "<empty>"})`,
    );
    onListening?.();
  });
  wss.on("error", (error) => {
    if (onListenError) onListenError(error);
    else console.error(`[bridge] WebSocket server error: ${error.message}`);
  });

  return {
    execution,
    workbookIdFor,
    acknowledgeWorkbook: (key) => {
      const workbookId = workbookIdFor(key);
      const revision = execution.acknowledge(workbookId);
      for (const [id, pending] of pendingTools)
        if (pending.orphaned && pending.workbookId === workbookId) pendingTools.delete(id);
      return revision;
    },
    nextUserMessage,
    hasPendingUserMessages: (key) => (queues.get(key)?.queue.length ?? 0) > 0,
    pushUserMessage,
    clearUserMessages,
    // Connected panes with the document each one has open (used by /eval).
    listPanes: () =>
      [...panes.entries()]
        .filter(([key]) => !!paneWs(key))
        .map(([key, p]) => ({ key, host: p.context.host, activeDoc: p.context.activeDoc })),
    callTaskpaneTool,
    sendAssistantText,
    sendAssistantEvent,
    sendToTaskpane: send,
    getContext: (key) => ({ ...(panes.get(key)?.context ?? emptyContext(hostOfKey(key))) }),
    isTaskpaneConnected: (key) =>
      key ? !!paneWs(key) : [...panes.keys()].some((k) => !!paneWs(k)),
    // Test/observability surface: the bound address (null until listening)
    // and a clean shutdown. Used by the loopback-bind unit test.
    address: () => wss.address(),
    close: () => new Promise((resolve) => wss.close(resolve)),
  };
}
