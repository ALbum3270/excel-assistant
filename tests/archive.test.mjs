import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import vm from "node:vm";
import { eventsFromLine } from "../daemon/transcript.mjs";

// index.mjs boots servers on import; evaluate only the archive functions.
const source = readFileSync(new URL("../daemon/index.mjs", import.meta.url), "utf8")
  .split(String.fromCharCode(13))
  .join("");
const start = source.indexOf("const ARCHIVE_MAX_BYTES");
const end = source.indexOf("async function ensureLoopForMessage", start);
assert.ok(start >= 0 && end > start, "archive source anchors must exist");

function archiveApi({ events, title = "Conversation" }) {
  const imported = [];
  const sandbox = {
    Buffer,
    JSON,
    Math,
    Date,
    getSessionRecord: async () => ({ title }),
    readTranscript: async () => ({ events, truncated: false }),
    documentKeyForPane: () => "doc",
    importArchivedSession: async (_host, _doc, archive) => {
      imported.push(archive);
      return "archive_1";
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(
    `${source.slice(start, end)}\nglobalThis.api = { exportConversation, importConversation };`,
    sandbox,
  );
  return { ...sandbox.api, imported };
}

test("a conversation with tool results exports and imports with each result paired to its call", async () => {
  // The events the transcript reader produces for a call and its receipt.
  const events = [
    ...eventsFromLine({ type: "user", message: { role: "user", content: "fill F" } }),
    ...eventsFromLine({
      type: "assistant",
      message: { role: "assistant", content: [{ type: "tool_use", id: "call_1", name: "mcp__office__excel_set_cell_range", input: { range: "F2" } }] },
    }),
    ...eventsFromLine({
      type: "user",
      message: { role: "user", content: [{ type: "tool_result", tool_use_id: "call_1", content: '{"success":true}' }] },
    }),
  ];
  const api = archiveApi({ events });
  const archive = await api.exportConversation("key", "excel", "session");
  const tool = archive.events.find((event) => event.kind === "tool");
  const result = archive.events.find((event) => event.kind === "tool_result");
  assert.equal(tool.id, "call_1");
  assert.equal(result.id, "call_1");

  await api.importConversation("key", "excel", JSON.parse(JSON.stringify(archive)));
  assert.equal(api.imported[0].events.length, 3);
});

test("an export of a long Chinese conversation fits the import limit in bytes", async () => {
  const events = Array.from({ length: 17 }, (_, index) => ({
    kind: index % 2 ? "assistant" : "user",
    text: "汉".repeat(100_000),
  }));
  const api = archiveApi({ events });
  const archive = await api.exportConversation("key", "excel", "session");
  const bytes = Buffer.byteLength(JSON.stringify(archive), "utf8");
  assert.ok(bytes <= 2_000_000, `export is ${bytes} bytes, over the 2,000,000-byte import limit`);
  assert.equal(archive.truncated, true);
  await api.importConversation("key", "excel", JSON.parse(JSON.stringify(archive)));
  assert.equal(api.imported.length, 1);
});
