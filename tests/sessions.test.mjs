// Unit tests for daemon/sessions.mjs — per-(host, document) session-id
// bookkeeping. Tests override $HOME via the environment so they touch a
// fresh tmpdir instead of ~/.claude/. sessions.mjs reads homedir() at
// import time, so the override must happen before the import — we use
// dynamic import() inside the test to do that cleanly.

import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, writeFile, readFile, readdir, rm } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

async function withFakeHome(fn) {
  const fakeHome = await mkdtemp(join(tmpdir(), "cc-office-sess-test-"));
  const prevHome = process.env.HOME;
  const prevUserProfile = process.env.USERPROFILE;
  process.env.HOME = fakeHome;
  process.env.USERPROFILE = fakeHome;
  try {
    // Cache-bust the dynamic import so each test gets a freshly-evaluated
    // module reading the new $HOME.
    const cacheBuster = "?t=" + Date.now() + Math.random();
    const mod = await import("../daemon/sessions.mjs" + cacheBuster);
    await fn(mod, fakeHome);
  } finally {
    if (prevHome === undefined) delete process.env.HOME;
    else process.env.HOME = prevHome;
    if (prevUserProfile === undefined) delete process.env.USERPROFILE;
    else process.env.USERPROFILE = prevUserProfile;
    await rm(fakeHome, { recursive: true, force: true });
  }
}

async function readStateFile(fakeHome) {
  try {
    const raw = await readFile(join(fakeHome, ".claude", "office-addins", "sessions.json"), "utf8");
    return JSON.parse(raw);
  } catch (e) {
    if (e.code === "ENOENT") return null;
    throw e;
  }
}

test("getSessionId returns null when no state file exists", async () => {
  await withFakeHome(async ({ getSessionId }) => {
    assert.equal(await getSessionId("word", "doc-a"), null);
  });
});

test("save → get round-trip is keyed by host AND document", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }) => {
    await saveSessionId("word", "doc-a", "/tmp/folderA", "word-sid");
    await saveSessionId("excel", "doc-a", "/tmp/folderA", "excel-sid");
    await saveSessionId("excel", "doc-b", "/tmp/folderA", "other-book-sid");

    assert.equal(await getSessionId("word", "doc-a"), "word-sid");
    assert.equal(await getSessionId("excel", "doc-a"), "excel-sid");
    assert.equal(await getSessionId("excel", "doc-b"), "other-book-sid");
    assert.equal(await getSessionId("word", "doc-b"), null);
  });
});

test("two workbooks in one folder remain independent after a module restart", async () => {
  await withFakeHome(async ({ saveSessionId }) => {
    await saveSessionId("excel", "book-a", "/tmp/shared", "session-a");
    await saveSessionId("excel", "book-b", "/tmp/shared", "session-b");
    const restarted = await import("../daemon/sessions.mjs?restart=" + Date.now() + Math.random());
    assert.equal(await restarted.getSessionId("excel", "book-a"), "session-a");
    assert.equal(await restarted.getSessionId("excel", "book-b"), "session-b");
  });
});

test("a workbook keeps its conversation when its workspace changes", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }, fakeHome) => {
    await saveSessionId("excel", "doc-a", "/tmp/folderA", "first");
    await saveSessionId("excel", "doc-a", "/tmp/folderB", "continued");
    assert.equal(await getSessionId("excel", "doc-a"), "continued");
    const state = await readStateFile(fakeHome);
    const records = Object.values(state.conversations.excel);
    assert.equal(records.length, 1);
    assert.equal(records[0].cwd, "/tmp/folderB");
    assert.deepEqual(Object.keys(state.folders).sort(), ["/tmp/folderA", "/tmp/folderB"]);
  });
});

test("normalizeHost: a bad host is a no-op / null", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }) => {
    await saveSessionId("powerpoint", "doc", "/tmp/x", "nope");
    assert.equal(await getSessionId("powerpoint", "doc"), null);
    assert.equal(await getSessionId("word", "doc"), null);
  });
});

test("touchFolder does not clobber a saved session id", async () => {
  await withFakeHome(async ({ touchFolder, saveSessionId, getSessionId }) => {
    await touchFolder("/tmp/folderB");
    assert.equal(await getSessionId("word", "doc-b"), null);

    await saveSessionId("word", "doc-b", "/tmp/folderB", "w1");
    await touchFolder("/tmp/folderB");
    assert.equal(await getSessionId("word", "doc-b"), "w1");
  });
});

test("concurrent saves preserve every document and host", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }, fakeHome) => {
    await Promise.all([
      saveSessionId("excel", "doc-a", "/tmp/folderA", "a-excel"),
      saveSessionId("word", "doc-a", "/tmp/folderA", "a-word"),
      saveSessionId("excel", "doc-b", "/tmp/folderA", "b-excel"),
    ]);

    assert.equal(await getSessionId("excel", "doc-a"), "a-excel");
    assert.equal(await getSessionId("word", "doc-a"), "a-word");
    assert.equal(await getSessionId("excel", "doc-b"), "b-excel");
    const state = await readStateFile(fakeHome);
    assert.deepEqual(Object.keys(state.folders), ["/tmp/folderA"]);
  });
});

test("reads wait for an earlier unawaited mutation", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }) => {
    const saving = saveSessionId("excel", "doc-a", "/tmp/folderA", "new-id");
    assert.equal(await getSessionId("excel", "doc-a"), "new-id");
    await saving;
  });
});

test("concurrent touch and saves do not clobber session ids", async () => {
  await withFakeHome(async ({ touchFolder, saveSessionId, getSessionId }) => {
    await saveSessionId("excel", "doc-a", "/tmp/folderA", "old-id");
    await Promise.all([
      touchFolder("/tmp/folderA"),
      saveSessionId("excel", "doc-a", "/tmp/folderA", "new-id"),
      saveSessionId("excel", "doc-b", "/tmp/folderB", "other-id"),
    ]);
    assert.equal(await getSessionId("excel", "doc-a"), "new-id");
    assert.equal(await getSessionId("excel", "doc-b"), "other-id");
  });
});

test("clear is ordered with concurrent saves", async () => {
  await withFakeHome(async ({ saveSessionId, clearSessionId, getSessionId }) => {
    const saveExcel = saveSessionId("excel", "doc-a", "/tmp/folderA", "excel-id");
    const saveWord = saveSessionId("word", "doc-a", "/tmp/folderA", "word-id");
    const clearExcel = clearSessionId("excel", "doc-a");
    await Promise.all([saveExcel, saveWord, clearExcel]);
    assert.equal(await getSessionId("excel", "doc-a"), null);
    assert.equal(await getSessionId("word", "doc-a"), "word-id");
  });
});

test("atomic writes leave no temporary session files", async () => {
  await withFakeHome(async ({ saveSessionId }, fakeHome) => {
    await saveSessionId("excel", "doc-a", "/tmp/folderA", "id");
    const directory = join(fakeHome, ".claude", "office-addins");
    assert.deepEqual((await readdir(directory)).sort(), ["sessions.json"]);
  });
});

test("a failed atomic replace does not poison later mutations", async () => {
  await withFakeHome(async ({ saveSessionId, getSessionId }, fakeHome) => {
    const directory = join(fakeHome, ".claude", "office-addins");
    const target = join(directory, "sessions.json");
    await mkdir(target, { recursive: true });
    await assert.rejects(saveSessionId("excel", "doc-a", "/tmp/folderA", "blocked"));

    await rm(target, { recursive: true, force: true });
    await saveSessionId("excel", "doc-a", "/tmp/folderA", "recovered");
    assert.equal(await getSessionId("excel", "doc-a"), "recovered");
    assert.deepEqual((await readdir(directory)).sort(), ["sessions.json"]);
  });
});

test("touchFolder silently drops disallowed system home children (macOS only)", async () => {
  if (process.platform !== "darwin") return; // The deny-list is macOS-specific.
  await withFakeHome(async ({ touchFolder }, fakeHome) => {
    await touchFolder(join(fakeHome, "Library"));
    const state = await readStateFile(fakeHome);
    // touchFolder returns early for $HOME children — nothing persisted.
    assert.ok(
      state === null || !state.folders || Object.keys(state.folders).length === 0,
      "an OS-managed $HOME child must never be persisted",
    );
  });
});

test("migrates v1 and v2 state by keeping folders and dropping ambiguous sessions", async () => {
  await withFakeHome(async ({ getSessionId }, fakeHome) => {
    const dir = join(fakeHome, ".claude", "office-addins");
    await mkdir(dir, { recursive: true });
    const v1 = {
      version: 1,
      folders: {
        "/tmp/legacy": {
          session_id: "old-uuid",
          last_used: "2026-05-01T00:00:00.000Z",
          display_name: "legacy",
        },
      },
    };
    await writeFile(join(dir, "sessions.json"), JSON.stringify(v1));
    // The v1 id can't be attributed to a host, so it's intentionally not
    // resumable post-migration.
    assert.equal(await getSessionId("word", "doc-a"), null);
    assert.equal(await getSessionId("excel", "doc-a"), null);
  });
});

test("v2 folder-scoped ids are not assigned to an arbitrary workbook", async () => {
  await withFakeHome(async ({ getSessionId, saveSessionId }, fakeHome) => {
    const dir = join(fakeHome, ".claude", "office-addins");
    await mkdir(dir, { recursive: true });
    await writeFile(
      join(dir, "sessions.json"),
      JSON.stringify({
        version: 2,
        folders: {
          "/tmp/shared": {
            last_used: "2026-05-01T00:00:00.000Z",
            display_name: "shared",
            sessions: { excel: "ambiguous-old-id" },
          },
        },
      }),
    );
    assert.equal(await getSessionId("excel", "book-a"), null);
    assert.equal(await getSessionId("excel", "book-b"), null);
    await saveSessionId("excel", "book-a", "/tmp/shared", "new-book-a-id");
    const state = await readStateFile(fakeHome);
    assert.equal(state.version, 3);
    assert.deepEqual(
      Object.values(state.conversations.excel).map((record) => record.session_id),
      ["new-book-a-id"],
    );
    assert.equal("sessions" in state.folders["/tmp/shared"], false);
  });
});

test("sessions.json hashes document identities instead of persisting raw URLs", async () => {
  await withFakeHome(async ({ saveSessionId }, fakeHome) => {
    const sensitiveUrl = "https://tenant.example/workbook.xlsx?access_token=secret";
    await saveSessionId("excel", sensitiveUrl, "/tmp/shared", "session-id");
    const raw = await readFile(join(fakeHome, ".claude", "office-addins", "sessions.json"), "utf8");
    assert.equal(raw.includes(sensitiveUrl), false);
    assert.equal(raw.includes("access_token"), false);
    assert.equal(raw.includes("session-id"), true);
  });
});
