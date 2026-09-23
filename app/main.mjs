// Excel Assistant — Electron menu bar shell.
//
// Wraps the daemon as a child process, exposes a tray icon with status and
// controls, hides the dock icon (we're a background-only app), and restarts
// the daemon if it crashes. The daemon code itself is untouched.

import { app, Tray, Menu, shell, dialog, nativeImage, BrowserWindow } from "electron";
import { spawn } from "node:child_process";
import { readFile, stat } from "node:fs/promises";
import fs from "node:fs";
import { homedir } from "node:os";
import { join, dirname } from "node:path";
import { fileURLToPath } from "node:url";
import { isAddinInstalled, installAddin, uninstallAddin } from "./sideload.mjs";
import { isSystemHomeChild } from "../daemon/system-paths.mjs";

const __dirname = dirname(fileURLToPath(import.meta.url));
const PROJECT_ROOT = join(__dirname, "..");
const DAEMON_ENTRY = join(PROJECT_ROOT, "daemon", "index.mjs");
const SESSIONS_FILE = join(homedir(), ".claude", "office-addins", "sessions.json");
const LOG_FILE = join(homedir(), ".claude", "office-addins", "daemon.log");

// Don't launch the daemon with an OS-managed $HOME child as the initial
// cwd, even if a stale sessions.json says so. Canonical set lives in
// daemon/system-paths.mjs (isSystemHomeChild).

let tray = null;
let daemonProcess = null;
let logStream = null;
let daemonStatus = "starting"; // "starting" | "running" | "crashed" | "stopped"
let currentWorkspace = null;
let restartAttempts = 0;
const MAX_RESTART = 3;
// The restart counter is only zeroed once the daemon has stayed up for this
// long. Resetting on the first "running" signal would let a daemon that
// crashes seconds after every start restart forever — the cap would never
// bite. See PR "crash-cap + restart backoff".
const STABLE_AFTER_MS = 45_000;
let stableTimer = null;
// One pending automatic restart and one start in flight at most. A crash used
// to schedule a restart that nothing could cancel, so Stop pressed in the next
// second was undone, and a manual start in that second raced it for the ports.
let restartTimer = null;
let startInFlight = null;
// Bumped by Stop, so a start that was still looking up its workspace when the
// user stopped does not go on to spawn.
let lifecycle = 0;

function cancelPendingRestart() {
  if (restartTimer) clearTimeout(restartTimer);
  restartTimer = null;
}

// --------------------------------------------------------------------------
// Daemon lifecycle
// --------------------------------------------------------------------------
async function findInitialWorkspace() {
  try {
    const state = JSON.parse(await readFile(SESSIONS_FILE, "utf8"));
    const folders = Object.entries(state.folders || {}).filter(([cwd]) => !isSystemHomeChild(cwd));
    if (folders.length === 0) return null;
    folders.sort(([, a], [, b]) => (b.last_used || "").localeCompare(a.last_used || ""));
    return folders[0][0];
  } catch {
    return null;
  }
}

function openLogStream() {
  fs.mkdirSync(dirname(LOG_FILE), { recursive: true });
  logStream = fs.createWriteStream(LOG_FILE, { flags: "a" });
  logStream.write(`\n=== ${new Date().toISOString()} app start ===\n`);
}

function startDaemon() {
  cancelPendingRestart();
  if (daemonProcess) return Promise.resolve();
  if (!startInFlight) {
    const starting = spawnDaemon().finally(() => {
      if (startInFlight === starting) startInFlight = null;
    });
    startInFlight = starting;
  }
  return startInFlight;
}

async function spawnDaemon() {
  const generation = lifecycle;
  const workspace = await findInitialWorkspace();
  if (generation !== lifecycle) return;
  currentWorkspace = workspace;
  daemonStatus = "starting";
  updateTray();

  const args = [DAEMON_ENTRY];
  if (workspace) args.push(workspace);

  daemonProcess = spawn(process.execPath, args, {
    cwd: PROJECT_ROOT,
    // fd 3 is an IPC channel for control messages (e.g. the native folder
    // picker). The daemon uses process.send / process.on('message') over this
    // channel; stdout/stderr remain plain log streams.
    stdio: ["ignore", "pipe", "pipe", "ipc"],
    env: { ...process.env, ELECTRON_RUN_AS_NODE: "1" },
  });

  // Arm the stability timer for this process. If it's still the live daemon
  // after STABLE_AFTER_MS, the start "stuck" and the restart budget is
  // refilled. A crash before then clears the timer (see the exit handler), so
  // a fast crash-loop never refills the budget and the cap eventually bites.
  if (stableTimer) clearTimeout(stableTimer);
  const thisProc = daemonProcess;
  stableTimer = setTimeout(() => {
    if (daemonProcess === thisProc && restartAttempts > 0) {
      restartAttempts = 0;
      logStream?.write(
        `[app] daemon stable for ${STABLE_AFTER_MS / 1000}s — restart counter reset\n`,
      );
    }
  }, STABLE_AFTER_MS);

  daemonProcess.stdout.on("data", (chunk) => {
    logStream?.write(chunk);
    const text = chunk.toString();
    if (text.includes("Starting agent loop") || text.includes("Starting session for")) {
      daemonStatus = "running";
      // NB: do not reset restartAttempts here — a crash-looping daemon prints
      // this on every start. The counter is only cleared by the stability
      // timer once the process has actually stayed up (see startDaemon).
      updateTray();
    }
    const m = /Starting session for (.+?) \(/.exec(text);
    if (m) {
      currentWorkspace = m[1];
      updateTray();
    }
  });

  daemonProcess.stderr.on("data", (chunk) => {
    logStream?.write(chunk);
  });

  daemonProcess.on("message", handleDaemonMessage);

  daemonProcess.on("exit", (code, signal) => {
    const wasIntentional = signal === "SIGTERM" || daemonStatus === "stopped";
    daemonProcess = null;
    // The process is gone before it proved stable — cancel the pending reset
    // so a fast crash-loop keeps spending the restart budget.
    if (stableTimer) {
      clearTimeout(stableTimer);
      stableTimer = null;
    }
    if (wasIntentional) {
      daemonStatus = "stopped";
    } else {
      daemonStatus = "crashed";
      logStream?.write(`\n[app] daemon exited unexpectedly (code=${code}, signal=${signal})\n`);
      if (restartAttempts < MAX_RESTART) {
        restartAttempts++;
        logStream?.write(`[app] auto-restarting (attempt ${restartAttempts}/${MAX_RESTART})\n`);
        restartTimer = setTimeout(() => {
          restartTimer = null;
          if (daemonStatus !== "stopped") startDaemon();
        }, 1000);
      } else {
        logStream?.write(`[app] giving up after ${MAX_RESTART} restart attempts\n`);
      }
    }
    updateTray();
  });
}

function stopDaemon() {
  // Also stops a restart that a crash scheduled but has not run yet.
  cancelPendingRestart();
  lifecycle += 1;
  startInFlight = null;
  daemonStatus = "stopped";
  daemonProcess?.kill("SIGTERM");
  updateTray();
}

function restartDaemon() {
  cancelPendingRestart();
  if (daemonProcess) {
    const proc = daemonProcess;
    const generation = lifecycle;
    daemonStatus = "starting";
    updateTray();
    proc.once("exit", () => {
      if (generation !== lifecycle) return;
      restartAttempts = 0;
      startDaemon();
    });
    proc.kill("SIGTERM");
  } else {
    restartAttempts = 0;
    startDaemon();
  }
}

// --------------------------------------------------------------------------
// IPC: handle control requests from the daemon. Currently just the native
// folder/file picker — the taskpane forwards picker requests over WebSocket
// to the daemon, which proxies them up here so we can show a real macOS
// open panel (NSOpenPanel, with sidebar shortcuts, Google Drive, iCloud,
// recent items — everything a synthetic in-page modal can't do).
// --------------------------------------------------------------------------
async function handleDaemonMessage(msg) {
  if (!msg) return;
  // Structured readiness signal over the IPC channel — replaces sniffing
  // stdout for a log substring (fragile: log copy changes, line
  // chunking). The stdout heuristic below is kept as a fallback for any
  // path that doesn't emit this.
  if (msg.type === "daemon_ready") {
    daemonStatus = "running";
    // Not reset here for the same reason as the stdout heuristic: a
    // crash-looping daemon emits daemon_ready on every start. Only the
    // stability timer clears the counter.
    if (msg.cwd) currentWorkspace = msg.cwd;
    updateTray();
    return;
  }
  if (msg.type === "restart_daemon") {
    restartDaemon();
    return;
  }
  if (msg.type !== "pick_path") return;
  const reply = (payload) => {
    try {
      daemonProcess?.send({ type: "pick_path_result", id: msg.id, ...payload });
    } catch (err) {
      logStream?.write(`[app] failed to reply to pick_path: ${err.message}\n`);
    }
  };
  try {
    // Single-mode dialogs only. Windows cannot show a combined
    // file+directory picker — ['openFile','openDirectory'] silently
    // degrades to a directory-only selector there, so files never appear
    // (the symptom: an empty chooser when adding a file context entry).
    // Callers pick exactly one mode: include_files → file, else folder.
    const properties = msg.include_files ? ["openFile"] : ["openDirectory", "createDirectory"];
    // We're tray-only (no app window), so the OS open panel surfaces
    // BEHIND whatever is frontmost (Word/Excel). macOS: force-focus the
    // app and the panel comes forward (dock is hidden, so nothing else
    // shows). Windows: app.focus() just flashes the taskbar and the
    // dialog still opens behind Excel — instead, parent the dialog to a
    // transient off-screen always-on-top window so it's brought to the
    // foreground. The helper window is positioned off-screen and
    // destroyed immediately after, so it never visibly appears.
    let dialogParent = null;
    if (process.platform === "darwin") {
      app.focus({ steal: true });
    } else if (process.platform === "win32") {
      dialogParent = new BrowserWindow({
        width: 1,
        height: 1,
        x: -4000,
        y: -4000,
        show: false,
        frame: false,
        skipTaskbar: true,
        alwaysOnTop: true,
      });
      dialogParent.showInactive();
      dialogParent.setAlwaysOnTop(true, "screen-saver");
      dialogParent.focus();
    }
    const dialogOpts = {
      title: msg.title || (msg.include_files ? "Choose a folder or file" : "Choose a folder"),
      buttonLabel: msg.button_label || "Use this",
      defaultPath: msg.default_path || undefined,
      properties,
    };
    let result;
    try {
      result = dialogParent
        ? await dialog.showOpenDialog(dialogParent, dialogOpts)
        : await dialog.showOpenDialog(dialogOpts);
    } finally {
      dialogParent?.destroy();
    }
    if (result.canceled || result.filePaths.length === 0) {
      reply({ ok: true, canceled: true });
      return;
    }
    const picked = result.filePaths[0];
    const s = await stat(picked);
    reply({ ok: true, path: picked, kind: s.isDirectory() ? "directory" : "file" });
  } catch (err) {
    reply({ ok: false, error: err.message });
  }
}

// --------------------------------------------------------------------------
// Tray UI
// --------------------------------------------------------------------------
function statusLabel() {
  switch (daemonStatus) {
    case "starting":
      return "● Starting…";
    case "running":
      return "● Ready";
    case "crashed":
      return restartAttempts >= MAX_RESTART
        ? "● Crashed (won't restart)"
        : `● Crashed (restart ${restartAttempts}/${MAX_RESTART})`;
    case "stopped":
      return "● Stopped";
    default:
      return "● ?";
  }
}

function workspaceLabel() {
  if (!currentWorkspace) return "Workspace: (none)";
  const name = currentWorkspace.split(/[\\/]/).filter(Boolean).pop();
  return `Workspace: ${name}`;
}

function buildMenu() {
  return Menu.buildFromTemplate([
    { label: statusLabel(), enabled: false },
    { label: workspaceLabel(), enabled: false },
    { type: "separator" },
    { label: "Open logs", click: () => shell.openPath(LOG_FILE) },
    {
      label: daemonProcess ? "Restart daemon" : "Start daemon",
      click: restartDaemon,
    },
    ...(daemonProcess ? [{ label: "Stop daemon", click: stopDaemon }] : []),
    {
      label: "Show in Finder",
      submenu: [
        { label: "Project folder", click: () => shell.openPath(PROJECT_ROOT) },
        ...(currentWorkspace
          ? [{ label: "Current workspace", click: () => shell.openPath(currentWorkspace) }]
          : []),
        { label: "Log file", click: () => shell.showItemInFolder(LOG_FILE) },
      ],
    },
    { type: "separator" },
    { label: "Open Microsoft Excel", click: () => shell.openExternal("ms-excel:") },
    { type: "separator" },
    {
      label: addinInstalled ? "Reinstall add-in in Excel" : "Install add-in in Excel…",
      click: () => runInstall({ interactive: true }),
    },
    ...(addinInstalled
      ? [{ label: "Uninstall add-in", click: () => runUninstall({ interactive: true }) }]
      : []),
    { type: "separator" },
    { label: "Quit Excel Assistant", click: () => app.quit() },
  ]);
}

function updateTray() {
  if (!tray) return;
  tray.setToolTip(`Excel Assistant — ${statusLabel().replace(/^●\s*/, "")}`);
  tray.setContextMenu(buildMenu());
}

// --------------------------------------------------------------------------
// Add-in sideload — install / uninstall the Excel manifest so the
// user never has to manually drop files into wef/ or touch Trust Center.
// --------------------------------------------------------------------------
let addinInstalled = false;

async function refreshAddinInstalled() {
  try {
    addinInstalled = await isAddinInstalled();
  } catch {
    addinInstalled = false;
  }
  updateTray();
}

async function runInstall({ interactive }) {
  try {
    const result = await installAddin();
    addinInstalled = true;
    updateTray();
    if (interactive) {
      dialog
        .showMessageBox({
          type: "info",
          title: "Excel Assistant installed",
          message: "The add-in is now registered with Excel.",
          detail:
            "Quit and reopen Excel (if it's already running), then look for " +
            "Excel Assistant under Insert → Office Add-ins.\n\n" +
            (process.platform === "win32"
              ? `Trusted catalog registered at:\n${result.catalog}`
              : `Manifests copied to each app's wef/ folder.`),
          buttons: ["OK"],
        })
        .catch(() => {});
    }
    return result;
  } catch (err) {
    logStream?.write(`[app] install failed: ${err.message}\n`);
    if (interactive) {
      dialog
        .showMessageBox({
          type: "error",
          title: "Couldn't install the add-in",
          message: err.message,
          detail:
            "You can sideload manually instead — see the README's Sideload section. " +
            "The shortest path: in Excel, Insert → My Add-ins → Upload My Add-in → " +
            `pick excel.xml from ${join(PROJECT_ROOT, "manifests")}.`,
          buttons: ["OK"],
        })
        .catch(() => {});
    }
  }
}

async function runUninstall({ interactive }) {
  try {
    await uninstallAddin();
    addinInstalled = false;
    updateTray();
    if (interactive) {
      dialog
        .showMessageBox({
          type: "info",
          title: "Add-in uninstalled",
          message: "Excel Assistant is no longer registered with Excel.",
          detail:
            "The daemon is still running. Quit Excel Assistant from the tray menu to stop it entirely.",
          buttons: ["OK"],
        })
        .catch(() => {});
    }
  } catch (err) {
    logStream?.write(`[app] uninstall failed: ${err.message}\n`);
    if (interactive) {
      dialog
        .showMessageBox({
          type: "error",
          title: "Couldn't uninstall the add-in",
          message: err.message,
          buttons: ["OK"],
        })
        .catch(() => {});
    }
  }
}

// First-run prompt — invoked once at app start if the add-in isn't already
// registered. Skipped silently on unsupported platforms (Linux, etc.).
async function offerFirstRunInstall() {
  if (process.platform !== "darwin" && process.platform !== "win32") return;
  if (addinInstalled) return;
  const { response } = await dialog
    .showMessageBox({
      type: "question",
      title: "Install Excel Assistant in Excel?",
      message:
        "Excel Assistant can install itself in Excel automatically — no manifest copying or registry editing needed.",
      detail:
        "Click Install to register the add-in now. You can install later from the tray menu " +
        "if you'd prefer. After installing, open Excel and find Excel Assistant under " +
        "Insert → Office Add-ins.",
      buttons: ["Install", "Not now"],
      defaultId: 0,
      cancelId: 1,
    })
    .catch(() => ({ response: 1 }));
  if (response === 0) await runInstall({ interactive: true });
}

// --------------------------------------------------------------------------
// App lifecycle
// --------------------------------------------------------------------------
app.whenReady().then(async () => {
  // Background-only app — no dock icon on macOS.
  if (process.platform === "darwin") app.dock?.hide();

  openLogStream();

  // Tray icon. macOS expects a template (monochrome with transparency, auto-
  // recolored per theme); Windows expects a full-color .ico (or PNG fallback).
  // Convention: tray-icon.png is the macOS template; tray-icon-win.ico is the
  // Windows full-color version. Falls back to the PNG if .ico is missing.
  const macIconPath = join(__dirname, "tray-icon.png");
  const winIconPath = join(__dirname, "tray-icon-win.ico");
  const iconPath =
    process.platform === "win32" && fs.existsSync(winIconPath) ? winIconPath : macIconPath;
  const icon = nativeImage.createFromPath(iconPath);
  if (process.platform === "darwin") icon.setTemplateImage(true);
  tray = new Tray(icon);
  tray.setToolTip("Excel Assistant");

  // Check whether the add-in is already registered and refresh the tray
  // menu, then offer to install on first run.
  await refreshAddinInstalled();
  updateTray();
  startDaemon();
  offerFirstRunInstall().catch((err) =>
    logStream?.write(`[app] first-run prompt failed: ${err.message}\n`),
  );
});

// Don't quit on "all windows closed" — we have no windows; tray is the UI.
app.on("window-all-closed", (e) => {
  if (e && e.preventDefault) e.preventDefault();
});

app.on("before-quit", () => {
  stopDaemon();
  logStream?.write(`=== ${new Date().toISOString()} app quit ===\n`);
  logStream?.end();
});
