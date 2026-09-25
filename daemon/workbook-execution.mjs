import { randomUUID } from "node:crypto";
import { win32 } from "node:path";
import { fileURLToPath } from "node:url";
import { createWorkbookCoordinator } from "./vendor/pi-coordinator.mjs";

// A stale expectation usually means something else changed the workbook. The
// most common something is the panel's own restore button, which the agent
// never sees, so name it rather than leaving the model to guess.
function describeLastWrite(state) {
  if (state.lastWriteUncertain) {
    return `The ${state.lastWrite || "previous"} operation ended without a confirmed commit; its effect on the workbook is unknown.`;
  }
  if (state.lastWrite === "excel_workbook_history") {
    return "The change was a restore from the panel's backups, so the cells you read earlier may be back to their previous values.";
  }
  return state.lastWrite
    ? `The change came from ${state.lastWrite}.`
    : "The change did not come from this session.";
}

// A tool result that reports its own refusal (the overwrite guard) or says it
// never committed left the workbook untouched.
function didCommit(result) {
  if (!result || typeof result !== "object") return true;
  return result.success !== false && result.commitStatus !== "not_committed";
}

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
              `Workbook changed since revision ${expectedRevision}; current revision is ${state.revision}. ` +
                `${describeLastWrite(state)} ` +
                "Read the target range again before writing — retrying this write unchanged fails the same way.",
              "STALE_WORKBOOK_REVISION",
              "not_committed",
              { workbookRevision: state.revision },
            );
          }
          started = true;
          // The revision means "the workbook changed", and writers compare
          // their expectation against it. A write that is refused before it
          // touches a cell — the overwrite guard, a shape check — must not
          // burn a number, or every later write in the session is stale
          // against a change that never happened. So run against the next
          // number and only keep it when the write actually landed.
          const pending = write ? state.revision + 1 : state.revision;
          try {
            const result = await execute({
              signal: stop.signal,
              revision: pending,
              opId,
              write,
            });
            if (write && didCommit(result)) {
              state.revision = pending;
              state.lastWrite = toolName;
              state.lastWriteUncertain = false;
            }
            this.settled(workbookId, opId);
            return { result, revision: state.revision, uncertain: Boolean(state.blocked) };
          } catch (error) {
            // Settled means the operation stopped running, not that nothing
            // changed: a write can apply its first batch and fail on a later
            // sync, and a cancelled write can finish after the caller gave up.
            // Only an explicit not_committed keeps the number; tying it to
            // "still running" let stale writers through after a partial write.
            const mayHaveChanged = write && error.commitStatus !== "not_committed";
            if (mayHaveChanged) {
              state.revision = pending;
              state.lastWrite = toolName;
              state.lastWriteUncertain = true;
            }
            if (mayHaveChanged && !error.executionSettled) {
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
