// Read a resumed SDK session's transcript back into the chat panel.
//
// The Agent SDK persists every session as newline-delimited JSON at
// ~/.claude/projects/<encoded-cwd>/<session_id>.jsonl. We reconstruct the
// user-visible bubble list from it so reopening the add-in (or restarting
// the daemon / Office) shows the prior conversation — and shows EXACTLY
// what the agent still remembers, since it's the same file the SDK resumes
// from.
//
// We locate the file by globbing for `<session_id>.jsonl` under every
// project dir, so we never depend on Claude Code's (undocumented) cwd ->
// directory-name encoding. Schema observed empirically (2026-05-16):
//
//   { type:"user",      message:{ role:"user",      content:<str|array> } }
//   { type:"assistant", message:{ role:"assistant", content:[ blocks ] } }
//   { type:"queue-operation"|"ai-title"|"attachment"|"last-prompt", ... }  -> skip
//
//   user.content as a string  -> a user-typed message (carries our injected
//                                 [Doc: … · Selection: …] header to strip)
//   user.content as an array  -> tool_result blocks (skip; not user-typed)
//   assistant blocks: text -> assistant bubble; tool_use -> tool announce;
//                     thinking -> skip (not shown live either)
//
// The .jsonl schema is the SDK's internal format. The daemon already
// depends on it for live messages, so this is the same coupling, not new.
// Every per-line parse is defensive (bad lines skipped).

import { readdir, stat, unlink } from "node:fs/promises";
import { createReadStream } from "node:fs";
import { createInterface } from "node:readline";
import { homedir } from "node:os";
import { join } from "node:path";

const PROJECTS_DIR = join(homedir(), ".claude", "projects");

// Strip the per-turn context header the daemon prepends to user messages
// (renderContextHeader: a single bracketed line containing Host:/Doc:/
// Selection:/Track changes:, then a blank line, then the user's actual
// text).
function stripContextHeader(text) {
  return text.replace(
    /^\[(?:Host:|Doc:|Selection:|Cursor in paragraph|Track changes:)[^\n]*\]\n\n?/,
    "",
  );
}

// Find <sessionId>.jsonl under any project dir. Returns the newest match by
// mtime, or null. No dependency on the cwd->dirname encoding scheme.
async function locateSessionFile(sessionId) {
  let dirs;
  try {
    dirs = await readdir(PROJECTS_DIR, { withFileTypes: true });
  } catch {
    return null;
  }
  const candidates = [];
  for (const d of dirs) {
    if (!d.isDirectory()) continue;
    const p = join(PROJECTS_DIR, d.name, `${sessionId}.jsonl`);
    try {
      const s = await stat(p);
      if (s.isFile()) candidates.push({ path: p, mtime: s.mtimeMs });
    } catch {
      /* not in this dir */
    }
  }
  if (candidates.length === 0) return null;
  candidates.sort((a, b) => b.mtime - a.mtime);
  return candidates[0].path;
}

// A receipt is small; anything larger is a bulk read the panel re-renders from
// its own preview, so cap what the replay carries.
const MAX_RESULT_CHARS = 4000;

function toolResultText(content) {
  const text = Array.isArray(content)
    ? content
        .filter((part) => part && part.type === "text" && typeof part.text === "string")
        .map((part) => part.text)
        .join("")
    : typeof content === "string"
      ? content
      : "";
  return text.slice(0, MAX_RESULT_CHARS);
}

function eventsFromLine(obj) {
  const out = [];
  const m = obj && obj.message;
  if (!m || typeof m !== "object") return out; // queue-operation/ai-title/etc.
  const role = m.role;
  const content = m.content;

  if (obj.type === "user" || role === "user") {
    if (typeof content === "string") {
      const text = stripContextHeader(content).trim();
      if (text) out.push({ kind: "user", text });
    } else if (Array.isArray(content)) {
      // Arrays carry tool_result blocks. They are not user-typed text, but
      // they hold the write receipts — status, range, verification, revision,
      // backup — so a replayed conversation shows what each change did instead
      // of a bare argument dump.
      for (const b of content) {
        if (!b || b.type !== "tool_result" || !b.tool_use_id) continue;
        out.push({
          kind: "tool_result",
          id: b.tool_use_id,
          isError: Boolean(b.is_error),
          text: toolResultText(b.content),
        });
      }
      const text = content
        .filter((b) => b && b.type === "text" && typeof b.text === "string")
        .map((b) => b.text)
        .join("")
        .trim();
      if (text) out.push({ kind: "user", text });
    }
    return out;
  }

  if (obj.type === "assistant" || role === "assistant") {
    if (!Array.isArray(content)) return out;
    for (const b of content) {
      if (!b || typeof b !== "object") continue;
      if (b.type === "text" && typeof b.text === "string") {
        const t = b.text.trim();
        if (t) out.push({ kind: "assistant", text: t });
      } else if (b.type === "tool_use") {
        out.push({ kind: "tool", id: b.id ?? null, name: b.name || "", input: b.input ?? {} });
      }
      // b.type === "thinking" -> skip (not shown in the live UI either)
    }
  }
  return out;
}

/**
 * Reconstruct the chat transcript for a session.
 * @param {string} sessionId
 * @param {{ maxEvents?: number }} opts
 * @returns {Promise<{ events: Array, truncated: boolean }>}
 *   events kinds: {kind:"user",text} | {kind:"assistant",text} |
 *                 {kind:"tool",name,input}
 *   truncated: true if older events were dropped to honor maxEvents.
 */
export async function readTranscript(sessionId, { maxEvents = 200 } = {}) {
  if (!sessionId) return { events: [], truncated: false };
  const file = await locateSessionFile(sessionId);
  if (!file) return { events: [], truncated: false };

  // Ring buffer of the last maxEvents render events — bounds memory and
  // render time for long sessions.
  const ring = [];
  let total = 0;

  const rl = createInterface({
    input: createReadStream(file, { encoding: "utf8" }),
    crlfDelay: Infinity,
  });
  try {
    for await (const line of rl) {
      if (!line.trim()) continue;
      let obj;
      try {
        obj = JSON.parse(line);
      } catch {
        continue;
      }
      for (const ev of eventsFromLine(obj)) {
        ring.push(ev);
        total++;
        if (ring.length > maxEvents) ring.shift();
      }
    }
  } finally {
    rl.close();
  }

  return { events: ring, truncated: total > ring.length };
}

export async function deleteTranscript(sessionId) {
  if (!sessionId) return false;
  const file = await locateSessionFile(sessionId);
  if (!file) return false;
  await unlink(file);
  return true;
}

// Exposed for testing / reuse.
export { stripContextHeader, locateSessionFile, eventsFromLine };
