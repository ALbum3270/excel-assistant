import { test } from "node:test";
import assert from "node:assert/strict";
import { mkdtemp, mkdir, readdir, readFile, writeFile, utimes } from "node:fs/promises";
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
  const folders = (await readdir(process.env.EXCEL_COM_BACKUP_DIR)).filter((name) =>
    name.startsWith("kept.xlsx-"),
  );
  assert.equal(folders.length, 1);
  const directory = join(process.env.EXCEL_COM_BACKUP_DIR, folders[0]);
  assert.equal((await readdir(directory)).length, 10);
});

test("a missing workbook file fails instead of silently skipping", async () => {
  await assert.rejects(() => backupWorkbookFile(join(workspace, "gone.xlsx")), /ENOENT/);
  await utimes(await workbook("book.xlsx", "one"), new Date(), new Date());
});

test("workbooks with the same name in different folders keep separate backups", async () => {
  const left = join(workspace, "left");
  const right = join(workspace, "right");
  await mkdir(left, { recursive: true });
  await mkdir(right, { recursive: true });
  await writeFile(join(left, "Budget.xlsx"), "left");
  await writeFile(join(right, "Budget.xlsx"), "right");
  const at = () => new Date(Date.UTC(2026, 8, 21, 12, 0, 0, 0));
  const a = await backupWorkbookFile(join(left, "Budget.xlsx"), { now: at });
  const b = await backupWorkbookFile(join(right, "Budget.xlsx"), { now: at });
  assert.notEqual(a.file, b.file);
  assert.equal(
    await readFile(a.file, "utf8"),
    "left",
    "the first backup is not replaced by the second workbook",
  );
  assert.equal(await readFile(b.file, "utf8"), "right");
});

test("two saves in the same millisecond both keep their backup", async () => {
  const path = await workbook("same-ms.xlsx", "first");
  const at = () => new Date(Date.UTC(2026, 8, 21, 12, 0, 0, 0));
  const first = await backupWorkbookFile(path, { now: at });
  await writeFile(path, "second, longer");
  const second = await backupWorkbookFile(path, { now: at });
  assert.notEqual(first.file, second.file);
  assert.equal(await readFile(first.file, "utf8"), "first");
  assert.equal(await readFile(second.file, "utf8"), "second, longer");
});
