import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, readdir, writeFile, utimes } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";

const workspace = await mkdtemp(join(tmpdir(), "com-backup-"));
process.env.EXCEL_COM_BACKUP_DIR = join(workspace, "backups");
const { backupWorkbookFile } = await import("../daemon/com-backup.mjs");

async function workbook(name, content) {
  const path = join(workspace, name);
  await writeFile(path, content);
  return path;
}

test("a COM write copies the workbook once per saved state", async () => {
  const path = await workbook("book.xlsx", "one");
  const first = await backupWorkbookFile(path);
  assert.equal(first.status, "copied");

  const again = await backupWorkbookFile(path);
  assert.equal(again.status, "reused");
  assert.equal(again.file, first.file);

  await writeFile(path, "two");
  const third = await backupWorkbookFile(path, { now: () => new Date(Date.now() + 1000) });
  assert.equal(third.status, "copied");
  assert.notEqual(third.file, first.file);
});

test("backups are turned off by EXCEL_COM_BACKUP=off", async () => {
  const path = await workbook("off.xlsx", "one");
  process.env.EXCEL_COM_BACKUP = "off";
  try {
    const result = await backupWorkbookFile(path);
    assert.equal(result.status, "skipped");
    assert.match(result.reason, /turned off/);
  } finally {
    delete process.env.EXCEL_COM_BACKUP;
  }
});

test("only the ten newest copies of a workbook are kept", async () => {
  const path = await workbook("kept.xlsx", "0");
  for (let index = 0; index < 14; index += 1) {
    await writeFile(path, String(index));
    // Distinct file names come from the timestamp, so move it forward.
    await backupWorkbookFile(path, { now: () => new Date(Date.UTC(2026, 0, 1, 0, index)) });
  }
  const directory = join(process.env.EXCEL_COM_BACKUP_DIR, "kept.xlsx");
  await mkdir(directory, { recursive: true });
  assert.equal((await readdir(directory)).length, 10);
});

test("a missing workbook file fails instead of silently skipping", async () => {
  await assert.rejects(() => backupWorkbookFile(join(workspace, "gone.xlsx")), /ENOENT/);
  await utimes(await workbook("book.xlsx", "one"), new Date(), new Date());
});
