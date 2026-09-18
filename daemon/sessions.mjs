// Per-(host, document) Agent SDK session bookkeeping.
//
// A workspace controls filesystem context. It is not a workbook identity:
// two workbooks can live in one folder, and one workbook can use a different
// explicitly selected workspace. The caller therefore supplies an opaque
// document key (normally the document URL from the bridge, or a pane id for
// an unsaved workbook) independently from cwd.
//
// Stored at ~/.claude/office-addins/sessions.json (version 3):
//
//   {
//     "version": 3,
//     "folders": {
//       "C:/Work": { "last_used": "…", "display_name": "Work" }
//     },
//     "conversations": {
//       "excel": {
//         "<sha256 of document key>": {
//           "session_id": "uuid…", "cwd": "C:/Work", "last_used": "…"
//         }
//       }
//     }
//   }

import { readFile, writeFile, mkdir, rename, unlink } from "node:fs/promises";
import { createHash, randomUUID } from "node:crypto";
import { homedir } from "node:os";
import { join, dirname, basename } from "node:path";
import { isSystemHomeChild } from "./system-paths.mjs";

const FILE = join(homedir(), ".claude", "office-addins", "sessions.json");
const VERSION = 3;
let mutationChain = Promise.resolve();

function emptyState() {
  return { version: VERSION, folders: {}, conversations: {} };
}

function isAllowedMatterPath(cwd) {
  return !isSystemHomeChild(cwd);
}

function normalizeHost(host) {
  return host === "word" || host === "excel" ? host : null;
}

function documentStorageKey(documentKey) {
  if (typeof documentKey !== "string" || !documentKey.trim()) return null;
  // Cloud document URLs can contain tenant-specific paths or query data.
  // Hash the opaque identity so sessions.json does not persist that URL.
  return createHash("sha256").update(documentKey).digest("hex");
}

async function readState() {
  let parsed;
  try {
    parsed = JSON.parse(await readFile(FILE, "utf8"));
  } catch (e) {
    if (e.code === "ENOENT") return emptyState();
    throw e;
  }
  if (!parsed || typeof parsed.folders !== "object" || parsed.folders === null) {
    return emptyState();
  }
  if (parsed.version === VERSION) {
    if (typeof parsed.conversations !== "object" || parsed.conversations === null) {
      parsed.conversations = {};
    }
    return parsed;
  }

  // v1 and v2 session ids were keyed by workspace (and, in v2, host).
  // No old entry can be safely assigned to one workbook when multiple files
  // share that folder, so retain recent folders and deliberately drop those
  // ambiguous resumable ids.
  const folders = {};
  for (const [cwd, info] of Object.entries(parsed.folders ?? {})) {
    folders[cwd] = {
      last_used: info?.last_used ?? new Date(0).toISOString(),
      display_name: info?.display_name ?? basename(cwd),
    };
  }
  return { version: VERSION, folders, conversations: {} };
}

async function writeState(state) {
  const directory = dirname(FILE);
  await mkdir(directory, { recursive: true });
  const temp = join(directory, `.sessions-${process.pid}-${randomUUID()}.tmp`);
  try {
    await writeFile(temp, JSON.stringify(state, null, 2) + "\n", {
      encoding: "utf8",
      mode: 0o600,
    });
    await rename(temp, FILE);
  } catch (error) {
    await unlink(temp).catch(() => {});
    throw error;
  }
}

function mutateState(mutator) {
  const operation = mutationChain.then(async () => {
    const state = await readState();
    if (await mutator(state)) await writeState(state);
  });
  mutationChain = operation.catch(() => {});
  return operation;
}

async function readCurrentState() {
  await mutationChain;
  return readState();
}

function ensureFolder(state, cwd) {
  let folder = state.folders[cwd];
  if (!folder) {
    folder = { last_used: new Date(0).toISOString(), display_name: basename(cwd) };
    state.folders[cwd] = folder;
  }
  delete folder.sessions;
  return folder;
}

function ensureHostConversations(state, host) {
  let conversations = state.conversations[host];
  if (!conversations || typeof conversations !== "object") {
    conversations = {};
    state.conversations[host] = conversations;
  }
  return conversations;
}

export async function getSessionId(host, documentKey) {
  const h = normalizeHost(host);
  const key = documentStorageKey(documentKey);
  if (!h || !key) return null;
  const state = await readCurrentState();
  return state.conversations?.[h]?.[key]?.session_id ?? null;
}

export async function saveSessionId(host, documentKey, cwd, sessionId) {
  const h = normalizeHost(host);
  const key = documentStorageKey(documentKey);
  if (!h || !key || typeof sessionId !== "string" || !sessionId) return;
  await mutateState((state) => {
    const now = new Date().toISOString();
    ensureHostConversations(state, h)[key] = { session_id: sessionId, cwd, last_used: now };
    if (cwd && isAllowedMatterPath(cwd)) {
      const folder = ensureFolder(state, cwd);
      folder.last_used = now;
      folder.display_name = basename(cwd);
    }
    return true;
  });
}

export async function clearSessionId(host, documentKey) {
  const h = normalizeHost(host);
  const key = documentStorageKey(documentKey);
  if (!h || !key) return;
  await mutateState((state) => {
    const conversations = state.conversations?.[h];
    if (!conversations?.[key]) return false;
    delete conversations[key];
    if (Object.keys(conversations).length === 0) delete state.conversations[h];
    return true;
  });
}

export async function touchFolder(cwd) {
  if (!cwd || !isAllowedMatterPath(cwd)) return;
  await mutateState((state) => {
    const folder = ensureFolder(state, cwd);
    folder.last_used = new Date().toISOString();
    folder.display_name = basename(cwd);
    return true;
  });
}
