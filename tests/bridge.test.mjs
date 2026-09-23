// Unit test for daemon/bridge.mjs network binding.
//
// Regression guard for the WS loopback-bind security fix: the `ws`
// library binds 0.0.0.0 when constructed with only `port`, exposing the
// agent bridge to the local network. createBridge must pass
// host: "127.0.0.1" so the bridge is loopback-only.

import { test } from "node:test";
import assert from "node:assert/strict";
import { once } from "node:events";
import WebSocket from "ws";

import { createBridge } from "../daemon/bridge.mjs";

// Wait for the ws server to finish binding (it binds asynchronously).
async function waitForListening(bridge, timeoutMs = 3000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    if (bridge.address()) return;
    await new Promise((r) => setTimeout(r, 20));
  }
  throw new Error("bridge did not start listening within timeout");
}

test("WS bridge binds to loopback (127.0.0.1), not all interfaces", async () => {
  // port: 0 → ephemeral free port, so the test never collides with a
  // running daemon or another test.
  const bridge = createBridge({ port: 0, token: "test-token", allowedOrigins: [] });
  try {
    await waitForListening(bridge);
    const addr = bridge.address();
    assert.ok(addr && typeof addr === "object", "address() should return the bound socket info");
    assert.equal(addr.address, "127.0.0.1", "WS bridge must bind loopback, not 0.0.0.0");
  } finally {
    await bridge.close();
  }
});

test("createBridge requires a token", () => {
  assert.throws(() => createBridge({ port: 0 }), /requires a token/);
});

test("user message carries its submit-time selection snapshot", async () => {
  const bridge = createBridge({ port: 0, token: "test-token", allowedOrigins: [] });
  let client;
  try {
    await waitForListening(bridge);
    client = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
    await once(client, "open");
    const welcome = once(client, "message");
    client.send(
      JSON.stringify({
        type: "hello",
        token: "test-token",
        host: "excel",
        active_doc: "book.xlsx",
        selection: { address: "Sheet1!A1", text: "old" },
      }),
    );
    await welcome;

    client.send(
      JSON.stringify({
        type: "user_message",
        text: "use the current cell",
        selection: { address: "Sheet1!B2", text: "new" },
      }),
    );
    const first = await bridge.nextUserMessage("excel\0book.xlsx");
    assert.deepEqual(first.context.selection, { address: "Sheet1!B2", text: "new" });

    client.send(
      JSON.stringify({ type: "user_message", text: "without selection", selection: null }),
    );
    const second = await bridge.nextUserMessage("excel\0book.xlsx");
    assert.equal(second.context.selection, null);
  } finally {
    client?.terminate();
    await bridge.close();
  }
});

test("aborting a tool call notifies the owning pane and rejects with unknown commit state", async () => {
  const bridge = createBridge({ port: 0, token: "test-token", allowedOrigins: [] });
  let client;
  try {
    await waitForListening(bridge);
    client = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
    await once(client, "open");
    const welcome = once(client, "message");
    client.send(
      JSON.stringify({
        type: "hello",
        token: "test-token",
        host: "excel",
        active_doc: "book.xlsx",
      }),
    );
    await welcome;

    const controller = new AbortController();
    const toolCallMessage = once(client, "message");
    const pending = bridge.callTaskpaneTool(
      "excel_set_cell_range",
      { range: "A1" },
      "excel\0book.xlsx",
      { signal: controller.signal },
    );
    const [toolCallRaw] = await toolCallMessage;
    const toolCall = JSON.parse(toolCallRaw.toString());
    assert.equal(toolCall.type, "tool_call");

    const cancelMessage = once(client, "message");
    controller.abort();
    await assert.rejects(pending, (error) => {
      assert.equal(error.code, "TOOL_CANCELLED");
      assert.equal(error.commitStatus, "unknown");
      return true;
    });
    const [cancelRaw] = await cancelMessage;
    assert.deepEqual(JSON.parse(cancelRaw.toString()), { type: "tool_cancel", id: toolCall.id });
  } finally {
    client?.terminate();
    await bridge.close();
  }
});

test("structured taskpane errors preserve code and commit state", async () => {
  const bridge = createBridge({ port: 0, token: "test-token", allowedOrigins: [] });
  let client;
  try {
    await waitForListening(bridge);
    client = new WebSocket(`ws://127.0.0.1:${bridge.address().port}`);
    await once(client, "open");
    const welcome = once(client, "message");
    client.send(
      JSON.stringify({
        type: "hello",
        token: "test-token",
        host: "excel",
        active_doc: "book.xlsx",
      }),
    );
    await welcome;
    const message = once(client, "message");
    const pending = bridge.callTaskpaneTool(
      "excel_clear_cell_range",
      { range: "A1" },
      "excel\0book.xlsx",
    );
    const [raw] = await message;
    const call = JSON.parse(raw.toString());
    client.send(
      JSON.stringify({
        type: "tool_result",
        id: call.id,
        ok: false,
        error: {
          message: "write rejected",
          code: "OVERWRITE_BLOCKED",
          commitStatus: "not_committed",
          recovery: { status: "checkpoint_created", snapshotIds: ["before-write"] },
        },
      }),
    );
    await assert.rejects(pending, (error) => {
      assert.equal(error.message, "write rejected");
      assert.equal(error.code, "OVERWRITE_BLOCKED");
      assert.equal(error.commitStatus, "not_committed");
      assert.deepEqual(error.recovery, {
        status: "checkpoint_created",
        snapshotIds: ["before-write"],
      });
      return true;
    });
  } finally {
    client?.terminate();
    await bridge.close();
  }
});
