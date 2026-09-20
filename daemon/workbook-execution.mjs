import { randomUUID } from "node:crypto";
import { win32 } from "node:path";
import { fileURLToPath } from "node:url";
import { createWorkbookCoordinator } from "./vendor/pi-coordinator.mjs";

export function canonicalWorkbookId(value) {
  let id = String(value || "");
  if (id.startsWith("file:")) id = fileURLToPath(id);
  if (/^[a-z]:[\\/]|^\\\\/i.test(id)) return win32.normalize(id).toLowerCase();
  return id;
}

function failure(message, code, commitStatus = "not_committed", extra = {}) {
  return Object.assign(new Error(message), { code, commitStatus, ...extra });
}

// Pi supplies FIFO scheduling. Its runRead doesn't wait for writes and its
// revision counts only successful writes, so all gateway operations use its
// queue while this adapter owns the externally visible mutation revision.
export function createWorkbookExecution() {
  const queue = createWorkbookCoordinator();
  const states = new Map();
  const stateFor = (id) => {
    const key = canonicalWorkbookId(id);
    if (!states.has(key)) states.set(key, { revision: 0, blocked: null });
    return states.get(key);
  };
  return {
    snapshot(id) {
      const key = canonicalWorkbookId(id);
      return { ...stateFor(key), active: queue.getSnapshot(key).activeWrite };
    },
    settled(id, opId) {
      const state = stateFor(id);
      if (state.blocked?.opId === opId) state.blocked = null;
    },
    acknowledge(id) {
      if (queue.getSnapshot(canonicalWorkbookId(id)).activeWrite)
        throw new Error("The previous operation is still running. Wait for it to finish.");
      const state = stateFor(id);
      state.blocked = null;
      state.revision++;
      return state.revision;
    },
    async run(
      id,
      {
        write = false,
        expectedRevision,
        signal,
        timeoutMs = 60_000,
        toolName,
        opId = randomUUID(),
      } = {},
      execute,
    ) {
      const workbookId = canonicalWorkbookId(id);
      const state = stateFor(workbookId);
      const stop = new AbortController();
      let started = false,
        rejectCaller;
      const cancelled = new Promise((_, reject) => {
        rejectCaller = reject;
      });
      let interrupted = false;
      const interrupt = (reason, code) => {
        if (interrupted) return;
        interrupted = true;
        const error = failure(reason, code, started && write ? "unknown" : "not_committed", {
          workbookRevision: state.revision,
        });
        if (started && write) state.blocked = { opId, toolName, reason };
        stop.abort(error);
        rejectCaller(error);
      };
      const onAbort = () =>
        interrupt(
          "Workbook operation cancelled; an already dispatched operation may still be finishing.",
          "TOOL_CANCELLED",
        );
      const timer = setTimeout(
        () =>
          interrupt(
            `Workbook operation timed out after ${timeoutMs}ms; check its status before retrying.`,
            "TOOL_TIMEOUT",
          ),
        timeoutMs,
      );
      signal?.addEventListener("abort", onAbort, { once: true });
      if (signal?.aborted) onAbort();
      const work = queue
        .runWrite({ workbookId, sessionId: "daemon", opId, toolName }, async () => {
          if (stop.signal.aborted) throw stop.signal.reason;
          if (write && state.blocked)
            throw failure(
              "A previous workbook operation has an unresolved outcome. Inspect Excel, then use the panel's recovery action before writing again.",
              "WORKBOOK_UNCERTAIN",
            );
          if (write && expectedRevision !== undefined && expectedRevision !== state.revision) {
            throw failure(
              `Workbook changed since revision ${expectedRevision}; current revision is ${state.revision}. Re-read the data before writing.`,
              "STALE_WORKBOOK_REVISION",
              "not_committed",
              { workbookRevision: state.revision },
            );
          }
          started = true;
          if (write) state.revision++;
          try {
            const result = await execute({
              signal: stop.signal,
              revision: state.revision,
              opId,
              write,
            });
            this.settled(workbookId, opId);
            return { result, revision: state.revision, uncertain: Boolean(state.blocked) };
          } catch (error) {
            if (write && error.commitStatus !== "not_committed" && !error.executionSettled) {
              state.blocked = { opId, toolName, reason: error.message };
            } else this.settled(workbookId, opId);
            throw Object.assign(error, { workbookRevision: state.revision });
          }
        })
        .then(({ result }) => result);
      try {
        return await Promise.race([work, cancelled]);
      } finally {
        clearTimeout(timer);
        signal?.removeEventListener("abort", onAbort);
      }
    },
  };
}
