// File-level safety net for COM writes.
//
// The COM tools drive Excel through its own object model, so we never learn
// which cells they touched and cannot take a range snapshot the way the
// Office.js tools do. What we can do is copy the workbook file before the
// write. The copy is the workbook as it was last saved — edits still only in
// Excel's memory are not in it — so every caller has to say that plainly.
import { copyFile, mkdir, readdir, stat, unlink } from "node:fs/promises";
import { homedir } from "node:os";
import { basename, extname, join } from "node:path";

const MAX_BYTES = 200 * 1024 * 1024;
const KEEP_PER_WORKBOOK = 10;

// canonical workbook path -> the copy we already made of that exact file state
const lastBackup = new Map();

export function comBackupRoot() {
  return (
    process.env.EXCEL_COM_BACKUP_DIR ||
    join(homedir(), ".claude", "office-addins", "com-backups")
  );
}

export function comBackupEnabled() {
  return String(process.env.EXCEL_COM_BACKUP ?? "").toLowerCase() !== "off";
}

function folderName(workbookPath) {
  return basename(workbookPath).replace(/[^\p{L}\p{N}._-]+/gu, "_").slice(0, 80) || "workbook";
}

function stamp(date) {
  return date.toISOString().replace(/[-:]/g, "").replace(/\.\d+Z$/, "Z");
}

async function pruneOlder(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = entries.filter((entry) => entry.isFile()).map((entry) => entry.name).sort();
  for (const name of files.slice(0, Math.max(0, files.length - KEEP_PER_WORKBOOK))) {
    await unlink(join(directory, name)).catch(() => {});
  }
}

/**
 * Copy the workbook file before a COM write.
 *
 * Returns {status: "copied"|"reused"|"skipped", file?, savedAt?, reason?}.
 * "reused" means the file has not changed on disk since the previous copy, so
 * that copy still represents the current saved state.
 */
export async function backupWorkbookFile(workbookPath, { now = () => new Date() } = {}) {
  if (!comBackupEnabled()) {
    return { status: "skipped", reason: "COM backups are turned off (EXCEL_COM_BACKUP=off)." };
  }
  const info = await stat(workbookPath);
  if (info.size > MAX_BYTES) {
    return {
      status: "skipped",
      reason: `The workbook is ${Math.round(info.size / 1024 / 1024)} MB, over the ${MAX_BYTES / 1024 / 1024} MB backup limit.`,
    };
  }
  const previous = lastBackup.get(workbookPath);
  if (previous && previous.mtimeMs === info.mtimeMs && previous.size === info.size) {
    return { status: "reused", file: previous.file, savedAt: new Date(info.mtimeMs).toISOString() };
  }

  const directory = join(comBackupRoot(), folderName(workbookPath));
  await mkdir(directory, { recursive: true });
  const extension = extname(workbookPath) || ".xlsx";
  const file = join(directory, `${stamp(now())}${extension}`);
  await copyFile(workbookPath, file);
  lastBackup.set(workbookPath, { mtimeMs: info.mtimeMs, size: info.size, file });
  await pruneOlder(directory);
  return { status: "copied", file, savedAt: new Date(info.mtimeMs).toISOString() };
}
