// File-level safety net for COM writes.
//
// The COM tools drive Excel through its own object model, so we never learn
// which cells they touched and cannot take a range snapshot the way the
// Office.js tools do. What we can do is copy the workbook file before the
// write. The copy is the workbook as it was last saved — edits still only in
// Excel's memory are not in it — so every caller has to say that plainly.
import { createHash } from "node:crypto";
import { constants } from "node:fs";
import { access, copyFile, mkdir, readdir, stat, unlink } from "node:fs/promises";
import { homedir } from "node:os";
import { basename, extname, join } from "node:path";

const MAX_BYTES = 200 * 1024 * 1024;
const KEEP_PER_WORKBOOK = 10;

// canonical workbook path -> the copy we already made of that exact file state
const lastBackup = new Map();

export function comBackupRoot() {
  return (
    process.env.EXCEL_COM_BACKUP_DIR || join(homedir(), ".claude", "office-addins", "com-backups")
  );
}

export function comBackupEnabled() {
  return String(process.env.EXCEL_COM_BACKUP ?? "").toLowerCase() !== "off";
}

// One folder per workbook, not per file name: a\Budget.xlsx and b\Budget.xlsx
// used to share a folder, so each could prune or replace the other's copies.
// The readable name is for people; the path digest is what keeps them apart.
function folderName(workbookPath) {
  const readable =
    basename(workbookPath)
      .replace(/[^\p{L}\p{N}._-]+/gu, "_")
      .slice(0, 60) || "workbook";
  const digest = createHash("sha256").update(workbookPath.toLowerCase()).digest("hex").slice(0, 12);
  return `${readable}-${digest}`;
}

function stamp(date) {
  return date.toISOString().replace(/[-:]/g, "").replace(".", "-");
}

// Never replace an existing copy: two saves in one millisecond, or a clock that
// repeats, get a suffix instead of silently overwriting the earlier backup.
async function copyWithoutOverwrite(source, directory, name, extension) {
  for (let attempt = 0; attempt < 100; attempt += 1) {
    const file = join(directory, `${name}${attempt ? `-${attempt}` : ""}${extension}`);
    try {
      await copyFile(source, file, constants.COPYFILE_EXCL);
      return file;
    } catch (error) {
      if (error?.code !== "EEXIST") throw error;
    }
  }
  throw new Error(`Could not find a free backup name in ${directory}`);
}

async function exists(file) {
  try {
    await access(file);
    return true;
  } catch {
    return false;
  }
}

async function pruneOlder(directory) {
  const entries = await readdir(directory, { withFileTypes: true });
  const files = entries
    .filter((entry) => entry.isFile())
    .map((entry) => entry.name)
    .sort();
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
  if (
    previous &&
    previous.mtimeMs === info.mtimeMs &&
    previous.size === info.size &&
    (await exists(previous.file))
  ) {
    return { status: "reused", file: previous.file, savedAt: new Date(info.mtimeMs).toISOString() };
  }

  const directory = join(comBackupRoot(), folderName(workbookPath));
  await mkdir(directory, { recursive: true });
  const extension = extname(workbookPath) || ".xlsx";
  const file = await copyWithoutOverwrite(workbookPath, directory, stamp(now()), extension);
  lastBackup.set(workbookPath, { mtimeMs: info.mtimeMs, size: info.size, file });
  await pruneOlder(directory);
  return { status: "copied", file, savedAt: new Date(info.mtimeMs).toISOString() };
}
