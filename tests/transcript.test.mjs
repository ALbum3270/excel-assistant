// Unit tests for daemon/transcript.mjs — the SDK-session .jsonl → chat
// bubble reconstruction. Pure transform + a tmpdir round-trip; no SDK,
// no Office.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

import { stripContextHeader, eventsFromLine, readTranscript } from "../daemon/transcript.mjs";

test("stripContextHeader removes the injected per-turn header", () => {
  assert.equal(
    stripContextHeader('[Host: Word · Doc: /x/y.docx · Selection: "hi"]\n\nFix the typo.'),
    "Fix the typo.",
  );
  assert.equal(stripContextHeader("No header here."), "No header here.");
  // Only the leading bracketed line is stripped, not later bracketed text.
  assert.equal(stripContextHeader("[Doc: a]\n\nsee [Note: keep this]"), "see [Note: keep this]");
});

test("eventsFromLine: user string → user bubble (header stripped)", () => {
  const evs = eventsFromLine({
    type: "user",
    message: { role: "user", content: "[Doc: a]\n\nhello" },
  });
  assert.deepEqual(evs, [{ kind: "user", text: "hello" }]);
});

test("eventsFromLine: user array (tool_result) → skipped", () => {
  const evs = eventsFromLine({
    type: "user",
    message: { role: "user", content: [{ type: "tool_result", content: "ok" }] },
  });
  assert.deepEqual(evs, []);
});

test("eventsFromLine: assistant text + tool_use, thinking skipped", () => {
  const evs = eventsFromLine({
    type: "assistant",
    message: {
      role: "assistant",
      content: [
        { type: "thinking", thinking: "hmm" },
        { type: "text", text: "Done." },
        { type: "tool_use", name: "mcp__office__office_replace_text", input: { a: 1 } },
      ],
    },
  });
  assert.deepEqual(evs, [
    { kind: "assistant", text: "Done." },
    { kind: "tool", name: "mcp__office__office_replace_text", input: { a: 1 } },
  ]);
});

test("eventsFromLine: internal lines (queue-operation/ai-title) → skipped", () => {
  assert.deepEqual(eventsFromLine({ type: "queue-operation", foo: 1 }), []);
  assert.deepEqual(eventsFromLine({ type: "ai-title", title: "x" }), []);
  assert.deepEqual(eventsFromLine({}), []);
});

test("readTranscript returns [] for a missing session", async () => {
  const r = await readTranscript("definitely-not-a-real-session-id-xyz");
  assert.deepEqual(r, { events: [], truncated: false });
  assert.deepEqual(await readTranscript(null), { events: [], truncated: false });
});

test("readTranscript round-trips a real .jsonl under ~/.claude/projects + truncates", async () => {
  // transcript.mjs reads $HOME/.claude/projects; override HOME to a tmpdir.
  const fakeHome = await mkdtemp(join(tmpdir(), "cc-office-transcript-test-"));
  const prevHome = process.env.HOME;
  const prevUserProfile = process.env.USERPROFILE;
  process.env.HOME = fakeHome;
  process.env.USERPROFILE = fakeHome;
  try {
    // transcript.mjs caches PROJECTS_DIR at import time from the original
    // HOME, so re-import a fresh copy bound to the fake HOME.
    const mod = await import("../daemon/transcript.mjs?t=" + Date.now() + Math.random());
    const projDir = join(fakeHome, ".claude", "projects", "proj-1");
    await mkdir(projDir, { recursive: true });
    const sid = "session-abc";
    const lines = [
      JSON.stringify({ type: "queue-operation", x: 1 }),
      JSON.stringify({ type: "user", message: { role: "user", content: "[Doc: a]\n\nfirst" } }),
      JSON.stringify({
        type: "assistant",
        message: { role: "assistant", content: [{ type: "text", text: "reply one" }] },
      }),
      "not json — should be skipped",
      JSON.stringify({ type: "user", message: { role: "user", content: "second" } }),
      JSON.stringify({
        type: "assistant",
        message: { role: "assistant", content: [{ type: "tool_use", name: "Bash", input: {} }] },
      }),
    ];
    await writeFile(join(projDir, `${sid}.jsonl`), lines.join("\n") + "\n");

    const full = await mod.readTranscript(sid);
    assert.equal(full.truncated, false);
    assert.deepEqual(full.events, [
      { kind: "user", text: "first" },
      { kind: "assistant", text: "reply one" },
      { kind: "user", text: "second" },
      { kind: "tool", name: "Bash", input: {} },
    ]);

    // maxEvents ring buffer: keep only the last 2, flag truncated.
    const clipped = await mod.readTranscript(sid, { maxEvents: 2 });
    assert.equal(clipped.truncated, true);
    assert.equal(clipped.events.length, 2);
    assert.deepEqual(
      clipped.events.map((e) => e.kind),
      ["user", "tool"],
    );

    assert.equal(await mod.deleteTranscript(sid), true);
    assert.deepEqual(await mod.readTranscript(sid), { events: [], truncated: false });
  } finally {
    if (prevHome === undefined) delete process.env.HOME;
    else process.env.HOME = prevHome;
    if (prevUserProfile === undefined) delete process.env.USERPROFILE;
    else process.env.USERPROFILE = prevUserProfile;
    await rm(fakeHome, { recursive: true, force: true });
  }
});
