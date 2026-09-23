// The pane's connection to the daemon: a WebSocket that presents the per-daemon
// bridge token, reconnects after a drop (re-fetching the token, which rotates
// when the daemon restarts), and a request/response helper keyed by request_id.

const REQUEST_TIMEOUT_MS = 10_000;
// Native picking may take minutes; allow the daemon's five-minute picker
// deadline to report its own result before the client expires.
const LONG_REQUESTS = { pick_path: 310_000 };

function newId() {
  return globalThis.crypto?.randomUUID?.() ?? "r_" + Math.random().toString(36).slice(2, 12);
}

export function createBridge({
  wsUrl = "ws://127.0.0.1:47833",
  tokenUrl = "/bridge-token",
  fetchImpl = globalThis.fetch?.bind(globalThis),
  WebSocketImpl = globalThis.WebSocket,
  retryMs = 1500,
  onOpen = () => {},
  onClose = () => {},
  onMessage = () => {},
} = {}) {
  let ws = null;
  let ready = false;
  let token = null;
  const pending = new Map();

  async function fetchToken() {
    try {
      const res = await fetchImpl(tokenUrl, { credentials: "same-origin" });
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      token = (await res.text()).trim();
    } catch (error) {
      console.warn("Failed to fetch bridge token:", error);
      token = null;
    }
  }

  function open() {
    ws = new WebSocketImpl(wsUrl);
    ws.onopen = () => {
      ready = true;
      onOpen();
    };
    ws.onclose = () => {
      ready = false;
      for (const request of pending.values()) {
        clearTimeout(request.timer);
        request.reject(new Error("Disconnected from daemon"));
      }
      pending.clear();
      onClose();
      setTimeout(() => fetchToken().then(open), retryMs);
    };
    ws.onerror = (error) => console.warn("WS error:", error);
    ws.onmessage = (event) => {
      let message;
      try {
        message = JSON.parse(event.data);
      } catch {
        return;
      }
      // Responses to request() end in "_result" and carry the request id.
      if (
        typeof message.type === "string" &&
        message.type.endsWith("_result") &&
        message.request_id
      ) {
        const request = pending.get(message.request_id);
        if (request) {
          pending.delete(message.request_id);
          clearTimeout(request.timer);
          request.resolve(message);
          return;
        }
      }
      onMessage(message);
    };
  }

  function send(message) {
    if (!ready) return false;
    ws.send(JSON.stringify(message));
    return true;
  }

  function request(type, payload = {}) {
    return new Promise((resolve, reject) => {
      if (!ready) {
        reject(new Error("Not connected to daemon"));
        return;
      }
      const request_id = newId();
      const timer = setTimeout(() => {
        if (pending.delete(request_id)) reject(new Error(`Request "${type}" timed out`));
      }, LONG_REQUESTS[type] ?? REQUEST_TIMEOUT_MS);
      pending.set(request_id, { resolve, reject, timer });
      send({ type, request_id, ...payload });
    });
  }

  return {
    // The bridge closes any connection that does not present the token.
    connect: () => fetchToken().then(open),
    send,
    request,
    get ready() {
      return ready;
    },
    get token() {
      return token;
    },
  };
}
