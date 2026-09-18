// Auto-sideload the Excel Assistant add-in manifest into Excel.
//
// macOS: drop the manifest into the wef/ folder inside Excel's container.
// Excel scans that folder on launch and surfaces every manifest under
// Insert → Office Add-ins → SHARED FOLDER.
//
//   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/excel.xml
//
// Windows: copy manifests to a folder under %APPDATA%, then register
// each manifest under the WEF\Developer registry key. Office scans that
// key on launch and sideloads each listed manifest directly — no admin,
// no network share. (A Trusted Catalog would require the folder to be a
// real SMB share with a UNC Url; Office silently ignores catalogs whose
// Url is a local path, which is why the catalog approach never showed
// the add-in. The Developer key is what the official Office dev tooling
// uses for local sideloading.)
//
//   %APPDATA%\ExcelAssistant\manifests\excel.xml
//   HKCU\Software\Microsoft\Office\16.0\WEF\Developer
//     <full manifest path> = <full manifest path>  (REG_SZ)
//
// All operations are idempotent: re-running install is safe, uninstall
// silently ignores already-removed pieces.

import { copyFile, mkdir, unlink, readdir } from "node:fs/promises";
import { existsSync } from "node:fs";
import { join, dirname } from "node:path";
import { homedir } from "node:os";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = join(__dirname, "..");
const MANIFESTS_DIR = join(PROJECT_ROOT, "manifests");

// One row per supported Office host. `mac_container` is the bundle ID whose
// sandbox we drop the manifest into on macOS. `guid` is purely informational
// here (the manifest carries its own Id element).
const HOSTS = [{ host: "Excel", file: "excel.xml", mac_container: "com.microsoft.Excel" }];

// ---------------------------------------------------------------------------
// macOS
// ---------------------------------------------------------------------------
function macWefDir(container) {
  return join(homedir(), "Library", "Containers", container, "Data", "Documents", "wef");
}
function macWefPath(host) {
  return join(macWefDir(host.mac_container), host.file);
}

async function macInstall() {
  const installed = [];
  for (const host of HOSTS) {
    const wef = macWefDir(host.mac_container);
    await mkdir(wef, { recursive: true });
    const src = join(MANIFESTS_DIR, host.file);
    const dst = join(wef, host.file);
    await copyFile(src, dst);
    installed.push({ host: host.host, path: dst });
  }
  return { installed };
}

async function macUninstall() {
  const removed = [];
  for (const host of HOSTS) {
    const p = macWefPath(host);
    try {
      await unlink(p);
      removed.push({ host: host.host, path: p });
    } catch (e) {
      if (e.code !== "ENOENT") throw e;
    }
  }
  return { removed };
}

function macIsInstalled() {
  return HOSTS.every((host) => existsSync(macWefPath(host)));
}

// ---------------------------------------------------------------------------
// Windows
// ---------------------------------------------------------------------------
// Office scans this key on launch and sideloads each listed manifest
// directly — the mechanism the official Office dev tooling uses. No
// admin, no SMB share. The value NAME is the manifest's full path (what
// Office reads); we set the data to the same path for legibility.
const WIN_DEVELOPER_KEY = "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\Developer";

// Legacy: earlier builds (incorrectly) registered a Trusted Catalog with
// a LOCAL-path Url. Office silently ignores such catalogs — a catalog Url
// must be a UNC share — which is why the add-in never appeared. We delete
// this on install + uninstall so upgrading users don't keep a dead entry.
const WIN_LEGACY_CATALOG_KEY =
  "HKCU\\Software\\Microsoft\\Office\\16.0\\WEF\\TrustedCatalogs\\claude-code-office-trusted-catalog";

function winCatalogDir() {
  return join(homedir(), "AppData", "Roaming", "ExcelAssistant", "manifests");
}

// This project started as a fork of Draftspect; drop that add-in's
// registrations so Excel doesn't list both.
async function winRemoveDraftspectRegistrations() {
  const legacyDir = join(homedir(), "AppData", "Roaming", "Draftspect", "manifests");
  for (const file of ["word.xml", "excel.xml"]) {
    const path = join(legacyDir, file);
    try {
      await regDeleteValue(WIN_DEVELOPER_KEY, path);
    } catch {
      /* not registered */
    }
    try {
      await unlink(path);
    } catch {
      /* already gone */
    }
  }
}

function regCommand(args) {
  return new Promise((resolve, reject) => {
    const p = spawn("reg.exe", args, { windowsHide: true });
    let stderr = "";
    p.stderr.on("data", (d) => {
      stderr += d.toString();
    });
    p.on("exit", (code) => {
      if (code === 0) resolve();
      else reject(new Error(`reg.exe ${args.join(" ")} exited ${code}: ${stderr.trim()}`));
    });
    p.on("error", reject);
  });
}

async function regAddString(key, name, value) {
  return regCommand(["add", key, "/v", name, "/t", "REG_SZ", "/d", value, "/f"]);
}
async function regDeleteKey(key) {
  return regCommand(["delete", key, "/f"]);
}
async function regDeleteValue(key, name) {
  return regCommand(["delete", key, "/v", name, "/f"]);
}
function regQueryValue(key, name) {
  return new Promise((resolve) => {
    const p = spawn("reg.exe", ["query", key, "/v", name], { windowsHide: true });
    p.on("exit", (code) => resolve(code === 0));
    p.on("error", () => resolve(false));
  });
}

async function winInstall() {
  const catalogDir = winCatalogDir();
  await mkdir(catalogDir, { recursive: true });
  const copied = [];
  for (const host of HOSTS) {
    const dst = join(catalogDir, host.file);
    await copyFile(join(MANIFESTS_DIR, host.file), dst);
    // Register the manifest under WEF\Developer (value name = path).
    await regAddString(WIN_DEVELOPER_KEY, dst, dst);
    copied.push({ host: host.host, path: dst });
  }
  await winRemoveDraftspectRegistrations();
  // Drop the dead legacy local-path catalog if a prior build left one.
  try {
    await regDeleteKey(WIN_LEGACY_CATALOG_KEY);
  } catch {
    /* not present — fine */
  }
  return { installed: copied, catalog: catalogDir, registry: WIN_DEVELOPER_KEY };
}

async function winUninstall() {
  const removed = [];
  const catalogDir = winCatalogDir();
  await winRemoveDraftspectRegistrations();
  for (const host of HOSTS) {
    const dst = join(catalogDir, host.file);
    try {
      await regDeleteValue(WIN_DEVELOPER_KEY, dst);
    } catch {
      /* value not present */
    }
    try {
      await unlink(dst);
      removed.push(dst);
    } catch {
      /* file gone — fine */
    }
  }
  // Also clear the legacy catalog key if present.
  try {
    await regDeleteKey(WIN_LEGACY_CATALOG_KEY);
  } catch {
    /* not present */
  }
  return { removed, registry: WIN_DEVELOPER_KEY };
}

async function winIsInstalled() {
  const catalogDir = winCatalogDir();
  // Installed iff every host's manifest is registered under the key.
  for (const host of HOSTS) {
    if (!(await regQueryValue(WIN_DEVELOPER_KEY, join(catalogDir, host.file)))) {
      return false;
    }
  }
  return true;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------
export async function isAddinInstalled() {
  if (process.platform === "darwin") return macIsInstalled();
  if (process.platform === "win32") return await winIsInstalled();
  return false;
}

export async function installAddin() {
  if (process.platform === "darwin") return macInstall();
  if (process.platform === "win32") return winInstall();
  throw new Error(
    `Auto-sideload not supported on ${process.platform}; see README for manual instructions.`,
  );
}

export async function uninstallAddin() {
  if (process.platform === "darwin") return macUninstall();
  if (process.platform === "win32") return winUninstall();
  throw new Error(`Auto-sideload not supported on ${process.platform}.`);
}
