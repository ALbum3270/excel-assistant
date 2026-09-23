// Per-(host, document) Agent SDK session bookkeeping.
//
// A workspace controls filesystem context. It is not a workbook identity:
// two workbooks can live in one folder, and one workbook can use a different
// explicitly selected workspace. The caller therefore supplies an opaque
// document key (normally the document URL from the bridge, or a pane id for
// an unsaved workbook) independently from cwd.
//
// Stored at ~/.claude/office-addins/sessions.json (version 4):
//
//   {
//     "version": 4,
//     "folders": {
//       "C:/Work": { "last_used": "…", "display_name": "Work" }
//     },
//     "conversations": {
//       "excel": {
//         "<sha256 of document key>": {
//           "active_session_id": "uuid…",
//           "sessions": [{ "session_id": "uuid…", "title": "Summarize Q3", "compatibility_key": "sha256…" }]
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
const VERSION = 4;
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

  // v3 was already keyed by host + workbook, but stored only the active
  // session. Preserve that session as the first history entry.
  if (parsed.version === 3) {
    const conversations = {};
    for (const host of ["word", "excel"]) {
      const oldHost = parsed.conversations?.[host];
      if (!oldHost || typeof oldHost !== "object") continue;
      for (const [key, old] of Object.entries(oldHost)) {
        if (!old?.session_id) continue;
        const lastUsed = old.last_used ?? new Date(0).toISOString();
        conversations[host] ??= {};
        conversations[host][key] = {
          active_session_id: old.session_id,
          sessions: [
            {
              session_id: old.session_id,
              cwd: old.cwd ?? null,
              title: "Previous conversation",
              created_at: lastUsed,
              last_used: lastUsed,
            },
          ],
        };
      }
    }
    parsed.version = VERSION;
    parsed.conversations = conversations;
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
  return state.conversations?.[h]?.[key]?.active_session_id ?? null;
}

export async function getSessionRecord(host, documentKey, sessionId = null) {
  const h = normalizeHost(host);
  const key = documentStorageKey(documentKey);
  if (!h || !key) return null;
  const state = await readCurrentState();
  const conversation = state.conversations?.[h]?.[key];
  const id = sessionId ?? conversation?.active_session_id;
  if (!id || !Array.isArray(conversation?.sessions)) return null;
  const saved = conversation.sessions.find((entry) => entry.session_id === id);
  return saved ? { ...saved } : null;
}

export async function saveSessionId(
  host,
  documentKey,
  cwd,
  sessionId,
  { title = null, compatibilityKey = null } = {},
) {
  const h = normalizeHost(host);
  const key = documentStorageKey(documentKey);
  if (!h || !key || typeof sessionId !== "string" || !sessionId) return;
  await mutateState((state) => {
    const now = new Date().toISOString();
    const conversations = ensureHostConversations(state, h);
    let conversation = conversations[key];
    if (!conversation || !Array.isArray(conversation.sessions)) {
      conversation = { active_session_id: null, sessions: [] };
      conversations[key] = conversation;
    }
    let saved = conversation.sessions.find((entry) => entry.session_id === sessionId);
    if (!saved) {
      saved = {
        session_id: sessionId,
        cwd: cwd ?? null,
        title: typeof title === "string" && title.trim() ? title.trim() : "New conversation",
        compatibility_key: compatibilityKey,
        created_at: now,
        last_used: now,
      };
      conversation.sessions.push(saved);
    } else {
      saved.cwd = cwd ?? saved.cwd ?? null;
      saved.last_used = now;
      if (compatibilityKey) saved.compatibility_key = compatibilityKey;
      if (
        typeof title === "string" &&
        title.trim() &&
        (!saved.title || saved.title === "New conversation")
      ) {
        saved.title = title.trim();
      }
    }
    conversation.active_session_id = sessionId;
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
    const conversation = state.conversations?.[h]?.[key];
    if (!conversation?.active_session_id) return false;
    conversation.active_session_id = null;
    return true;
  });
}

export async function listSessions(host, documentKey) {
  const h = normalizeHost(host);
  const key = documentStorageKey(documentKey);
  if (!h || !key) return { active_session_id: null, sessions: [] };
  const state = await readCurrentState();
  const conversation = state.conversations?.[h]?.[key];
  const sessions = Array.isArray(conversation?.sessions)
    ? conversation.sessions
        .filter((entry) => typeof entry?.session_id === "string" && entry.session_id)
        .map(({ archive_events, ...entry }) => ({
          ...entry,
          archived: Array.isArray(archive_events),
        }))
        .sort((a, b) => String(b.last_used ?? "").localeCompare(String(a.last_used ?? "")))
    : [];
  return {
    active_session_id: conversation?.active_session_id ?? null,
    sessions,
  };
}

export async function activateSession(host, documentKey, sessionId) {
  const h = normalizeHost(host);
  const key = documentStorageKey(documentKey);
  if (!h || !key || typeof sessionId !== "string" || !sessionId) return false;
  let activated = false;
  await mutateState((state) => {
    const conversation = state.conversations?.[h]?.[key];
    const saved = conversation?.sessions?.find((entry) => entry.session_id === sessionId);
    if (!saved) return false;
    saved.last_used = new Date().toISOString();
    conversation.active_session_id = sessionId;
    activated = true;
    return true;
  });
  return activated;
}

export async function deleteSession(host, documentKey, sessionId) {
  const h = normalizeHost(host);
  const key = documentStorageKey(documentKey);
  if (!h || !key || typeof sessionId !== "string" || !sessionId) return false;
  let deleted = false;
  await mutateState((state) => {
    const conversations = state.conversations?.[h];
    const conversation = conversations?.[key];
    if (!Array.isArray(conversation?.sessions)) return false;
    const index = conversation.sessions.findIndex((entry) => entry.session_id === sessionId);
    if (index === -1) return false;
    conversation.sessions.splice(index, 1);
    if (conversation.active_session_id === sessionId) conversation.active_session_id = null;
    if (conversation.sessions.length === 0) {
      delete conversations[key];
      if (Object.keys(conversations).length === 0) delete state.conversations[h];
    }
    deleted = true;
    return true;
  });
  return deleted;
}

export async function importArchivedSession(
  host,
  documentKey,
  { title, events, truncated = false },
) {
  const h = normalizeHost(host);
  const key = documentStorageKey(documentKey);
  if (!h || !key || !Array.isArray(events)) return null;
  const sessionId = `archive_${randomUUID()}`;
  await mutateState((state) => {
    const now = new Date().toISOString();
    const conversations = ensureHostConversations(state, h);
    let conversation = conversations[key];
    if (!conversation || !Array.isArray(conversation.sessions)) {
      conversation = { active_session_id: null, sessions: [] };
      conversations[key] = conversation;
    }
    conversation.sessions.push({
      session_id: sessionId,
      cwd: null,
      title:
        typeof title === "string" && title.trim()
          ? title.trim().slice(0, 120)
          : "Imported conversation",
      compatibility_key: null,
      archive_events: events,
      archive_truncated: Boolean(truncated),
      imported_at: now,
      created_at: now,
      last_used: now,
    });
    return true;
  });
  return sessionId;
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
