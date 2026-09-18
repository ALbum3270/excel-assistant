// Per-(host, workspace) Agent SDK session bookkeeping.
//
// The SDK persists each session's transcript to
// ~/.claude/projects/<hash>/<session_id>.jsonl. To resume via
// query({ options: { resume } }) we need that session_id. Word and Excel
// are independent surfaces — opening the same folder in Word vs Excel is
// a different conversation — so the session_id is keyed by **host AND
// cwd**, not cwd alone. (A single cwd-keyed session is why Excel used to
// replay Word's history.)
//
// The folder list itself (for the workspace switcher UI) stays
// cwd-deduped and host-agnostic — you pick a folder, not a folder+host.
//
// Stored at ~/.claude/office-addins/sessions.json (version 2):
//
//   {
//     "version": 2,
//     "folders": {
//       "/Users/x/Proj": {
//         "last_used": "2026-05-16T…Z",
//         "display_name": "Proj",
//         "sessions": { "word": "uuid…", "excel": "uuid…" }
//       }
//     }
//   }

import { readFile, writeFile, mkdir } from "node:fs/promises";
import { homedir } from "node:os";
import { join, dirname, basename } from "node:path";
import { isSystemHomeChild } from "./system-paths.mjs";

const FILE = join(homedir(), ".claude", "office-addins", "sessions.json");
const VERSION = 2;

// Don't persist or surface OS-managed $HOME children (e.g. ~/Library) as
// a recent workspace.
function isAllowedMatterPath(cwd) {
  return !isSystemHomeChild(cwd);
}

function normalizeHost(host) {
  return host === "word" || host === "excel" ? host : null;
}

async function readState() {
  let parsed;
  try {
    parsed = JSON.parse(await readFile(FILE, "utf8"));
  } catch (e) {
    if (e.code === "ENOENT") return { version: VERSION, folders: {} };
    throw e;
  }
  if (!parsed || typeof parsed.folders !== "object") {
    return { version: VERSION, folders: {} };
  }
  if (parsed.version === VERSION) return parsed;
  // Migrate v1 ({ session_id } per folder): the old id can't be attributed
  // to a host, so drop it (one conversation lost across the upgrade), but
  // keep the folder so it still shows in recents.
  const folders = {};
  for (const [cwd, info] of Object.entries(parsed.folders ?? {})) {
    folders[cwd] = {
      last_used: info?.last_used ?? new Date(0).toISOString(),
      display_name: info?.display_name ?? basename(cwd),
      sessions: {},
    };
  }
  return { version: VERSION, folders };
}

async function writeState(state) {
  await mkdir(dirname(FILE), { recursive: true });
  await writeFile(FILE, JSON.stringify(state, null, 2) + "\n", "utf8");
}

function ensureFolder(state, cwd) {
  let f = state.folders[cwd];
  if (!f) {
    f = { last_used: new Date(0).toISOString(), display_name: basename(cwd), sessions: {} };
    state.folders[cwd] = f;
  }
  if (typeof f.sessions !== "object" || f.sessions === null) f.sessions = {};
  return f;
}

// The resumable session_id for this host in this folder, or null.
export async function getSessionId(host, cwd) {
  const h = normalizeHost(host);
  if (!h) return null;
  const state = await readState();
  return state.folders[cwd]?.sessions?.[h] ?? null;
}

// Record the session_id for (host, cwd) and bump last_used.
export async function saveSessionId(host, cwd, sessionId) {
  const h = normalizeHost(host);
  if (!h) return;
  const state = await readState();
  const f = ensureFolder(state, cwd);
  f.sessions[h] = sessionId;
  f.last_used = new Date().toISOString();
  f.display_name = basename(cwd);
  await writeState(state);
}

// Forget the resumable session for (host, cwd) so the next message starts a
// fresh conversation. The transcript .jsonl stays on disk.
export async function clearSessionId(host, cwd) {
  const h = normalizeHost(host);
  if (!h) return;
  const state = await readState();
  if (!state.folders[cwd]?.sessions?.[h]) return;
  delete state.folders[cwd].sessions[h];
  await writeState(state);
}

// Touch a folder's bookkeeping (last_used / display_name) without changing
// any session id. Ensures the folder entry exists for the per-host session
// store. Skips OS-managed $HOME children so they never get persisted.
export async function touchFolder(cwd) {
  if (!isAllowedMatterPath(cwd)) return;
  const state = await readState();
  const f = ensureFolder(state, cwd);
  f.last_used = new Date().toISOString();
  f.display_name = basename(cwd);
  await writeState(state);
}
