import { test } from "node:test";
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { EventEmitter } from "node:events";
import vm from "node:vm";

// app/main.mjs imports Electron, so its lifecycle functions are evaluated from
// source with a stub process spawner and a clock the test drives.
const source = readFileSync(new URL("../app/main.mjs", import.meta.url), "utf8")
  .split(String.fromCharCode(13))
  .join("");
const slice = (from, to) => {
  const start = source.indexOf(from);
  const end = source.indexOf(to, start);
  assert.ok(start >= 0 && end > start, `missing anchors ${from} / ${to}`);
  return source.slice(start, end);
};

function lifecycle({ workspace = async () => null } = {}) {
  const timers = [];
  let spawned = 0;
  const sandbox = {
    findInitialWorkspace: workspace,
    currentWorkspace: null,
    daemonStatus: "stopped",
    daemonProcess: null,
    stableTimer: null,
    restartAttempts: 0,
    MAX_RESTART: 5,
    STABLE_AFTER_MS: 30_000,
    updateTray() {},
    DAEMON_ENTRY: "daemon",
    PROJECT_ROOT: "root",
    process: { execPath: "node", env: {} },
    logStream: null,
    handleDaemonMessage() {},
    setTimeout: (fn, ms) => {
      const timer = { fn, ms, cancelled: false };
      timers.push(timer);
      return timer;
    },
    clearTimeout: (timer) => {
      if (timer) timer.cancelled = true;
    },
    spawn() {
      spawned += 1;
      const child = new EventEmitter();
      child.stdout = new EventEmitter();
      child.stderr = new EventEmitter();
      child.kill = () => child.emit("exit", null, "SIGTERM");
      return child;
    },
  };
  vm.createContext(sandbox);
  vm.runInContext(slice("let restartTimer = null;", "async function findInitialWorkspace"), sandbox);
  vm.runInContext(
    slice("function startDaemon() {", "// --------------------------------------------------------------------------\n// IPC:") +
      "\nglobalThis.api = { startDaemon, stopDaemon, restartDaemon };",
    sandbox,
  );
  const fire = async (ms) => {
    for (const timer of timers.filter((t) => t.ms === ms && !t.cancelled)) await timer.fn();
  };
  return { sandbox, api: sandbox.api, fire, spawned: () => spawned };
}

test("Stop cancels a restart a crash had scheduled", async () => {
  const app = lifecycle();
  await app.api.startDaemon();
  app.sandbox.daemonProcess.emit("exit", 1, null); // crash: restart in 1 s
  app.api.restartDaemon(); // the user starts it at once
  await new Promise(setImmediate);
  app.api.stopDaemon(); // and stops it again
  await app.fire(1000); // the old one-second timer comes due
  await new Promise(setImmediate);
  assert.equal(app.spawned(), 2, "no third daemon from the cancelled restart");
  assert.equal(app.sandbox.daemonStatus, "stopped");
});

test("a start still looking up its workspace does not spawn after Stop", async () => {
  let release;
  const gate = new Promise((resolve) => (release = resolve));
  const app = lifecycle({ workspace: () => gate });
  const starting = app.api.startDaemon();
  app.api.stopDaemon();
  release(null);
  await starting;
  assert.equal(app.spawned(), 0);
});

test("two starts at once spawn one daemon", async () => {
  const app = lifecycle();
  await Promise.all([app.api.startDaemon(), app.api.startDaemon()]);
  assert.equal(app.spawned(), 1);
});
