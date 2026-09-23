// Runs the Office tool calls the daemon forwards and sends back one tool_result
// each. Everything that touches Excel, recovery or the bridge is passed in, so
// this file holds only the ordering rules: cell-edit probe, restore point, write
// under the workbook coordinator, receipt, and the cancel handshake.

export const WRITE_TOOLS = new Set([
  "excel_set_cell_range",
  "excel_clear_cell_range",
  "excel_copy_to",
  "excel_modify_sheet_structure",
  "excel_modify_workbook_structure",
  "excel_resize_range",
  "excel_modify_object",
  "excel_set_format",
  "excel_sort_range",
  "excel_autofilter",
  "excel_create_table",
  "excel_add_table_rows",
]);

export function isMutationCall(name, args) {
  return WRITE_TOOLS.has(name) || (name === "excel_workbook_history" && args?.action === "restore");
}

export function mutationTargets(name, args, result) {
  switch (name) {
    case "excel_set_cell_range":
      return [result?.writtenRange ?? args.copyToRange ?? args.range].filter(Boolean);
    case "excel_clear_cell_range":
      return [result?.clearedRange ?? args.range].filter(Boolean);
    case "excel_copy_to":
      return [result?.destination ?? args.destinationRange].filter(Boolean);
    case "excel_set_format":
    case "excel_sort_range":
    case "excel_create_table":
      return [result?.address ?? result?.range ?? args.address].filter(Boolean);
    case "excel_resize_range":
      return [args.range ?? "entire worksheet"];
    case "excel_modify_sheet_structure":
      return [`${args.dimension ?? "dimension"}:${args.reference ?? "pane"}:${args.count ?? 1}`];
    case "excel_modify_workbook_structure":
      return [
        result?.sheetName ?? args.newName ?? args.sheetName ?? `sheetId:${args.sheetId ?? "new"}`,
      ];
    case "excel_modify_object":
      return [
        result?.id ??
          args.id ??
          args.properties?.range ??
          args.properties?.anchor ??
          args.objectType,
      ].filter(Boolean);
    case "excel_autofilter":
      return [args.clear ? "worksheet autofilter" : (result?.address ?? args.address)].filter(
        Boolean,
      );
    case "excel_add_table_rows":
      return [args.table].filter(Boolean);
    case "excel_workbook_history":
      return result?.addresses ?? [];
    default:
      return [];
  }
}

// Excel.run has synced by the time a tool resolves. Attach one receipt shape to
// every mutation and distinguish mechanical read-back from the semantic check
// the agent still has to perform.
export function withMutationReceipt(name, args, result, receiptId) {
  if (!isMutationCall(name, args)) return result;
  const success = result?.success !== false;
  const readBack = name === "excel_set_cell_range" || name === "excel_copy_to";
  return {
    ...(result && typeof result === "object" ? result : { result }),
    success,
    commitStatus: result?.commitStatus ?? (success ? "committed" : "not_committed"),
    receiptId,
    operation: name,
    affectedTargets: mutationTargets(name, args, result),
    verification: {
      status: readBack ? "read_back" : "commit_acknowledged",
      semanticCheckRequired: true,
    },
  };
}

// Excel's rejection message is localized but says what it rejected; pass it
// through and only translate the error code. An earlier version appended a list
// of syntax rules Excel had not complained about, which buried its actual
// "argument invalid or missing" wording.
function hasFormulaInput(args) {
  return JSON.stringify(args?.cells ?? []).includes('"formula"');
}

export async function describeOfficeToolError(error, args, { getWorkbookMetadata } = {}) {
  // Office.js names the failing API (e.g. "RangeFormat.columnWidth") only in
  // debugInfo; the localized message alone doesn't say what was rejected.
  const location = error?.debugInfo?.errorLocation;
  const message = `${error?.message ?? String(error)}${location ? ` [at ${location}]` : ""}`;
  if (error?.code === "InvalidOperationInCellEditMode" && error?.commitStatus !== "not_committed") {
    return (
      `${message} (Excel entered cell-edit mode while this change was running, so part of it may ` +
      "have been applied. Ask the user to press Enter or Esc, then read the range before retrying.)"
    );
  }
  if (error?.code === "InvalidArgument" && hasFormulaInput(args)) {
    return (
      `${message} (Excel rejected the formula itself, not the range: InvalidArgument means a function ` +
      "argument is missing, extra or of the wrong kind. Read the message above — it is Excel's own " +
      "wording — and change the formula rather than resending it.)"
    );
  }
  if (!/Worksheet with ID .+ not found/i.test(message) || !getWorkbookMetadata) return message;
  try {
    const metadata = await getWorkbookMetadata();
    const valid = (metadata.sheetsMetadata ?? [])
      .map((sheet) => `${sheet.name}=${sheet.id}`)
      .join(", ");
    return `${message}. Valid worksheets in the current workbook: ${valid || "none"}. Refresh metadata and retry.`;
  } catch {
    return message;
  }
}

// Excel fails a whole batch while someone is editing a cell, before running any
// of it (Office.js RunOptions.delayForCellEdit). Probe before a write so it never
// starts and the outcome is known — nothing changed — instead of the write
// failing mid-way with an unknown outcome, which locks the workbook until
// someone confirms it by hand. A user who starts editing between the probe and
// the write still gets an unknown outcome; that race is real and stays honest.
export const CELL_EDIT_MODE_MESSAGE =
  "Excel is in cell-edit mode (someone is typing in a cell), so nothing was changed. " +
  "Ask the user to press Enter or Esc in Excel, then try again — retrying before that fails the same way.";

export async function refuseInCellEditMode(excel = globalThis.Excel) {
  try {
    await excel.run(async (context) => {
      context.workbook.load("name");
      await context.sync();
    });
  } catch (error) {
    if (error?.code === "InvalidOperationInCellEditMode") {
      throw Object.assign(new Error(CELL_EDIT_MODE_MESSAGE), {
        code: "InvalidOperationInCellEditMode",
        commitStatus: "not_committed",
      });
    }
    // Any other probe failure is the write's to report.
  }
}

// Saving a restore point is secondary to reporting what the write did. A storage
// failure (IndexedDB quota, a closed transaction) used to escape into the failure
// path, which committed a second time, threw again, and sent no tool_result at
// all while the workbook had in fact changed.
export async function settleRecovery(commit, result, failure) {
  try {
    return await commit(result, failure);
  } catch (error) {
    return {
      status: "not_available",
      reason: `The backup could not be saved: ${error?.message ?? String(error)}`,
    };
  }
}

const CANCELLED = Symbol("cancelled-tool-result");

/**
 * @param {object} deps
 * @param {(message: object) => void} deps.send               bridge send
 * @param {(name, args, id) => Promise<any>} deps.executeTool  the Office call itself
 * @param {(name, args, id) => Promise<Function|null>} deps.prepareMutationRecovery
 * @param {(id) => any} deps.takeMutationDiff
 * @param {(id, name, execute) => Promise<{result, revision}>} deps.runWorkbookWrite
 * @param {() => Promise<void>} [deps.refuseInCellEditMode]
 * @param {(error, args) => Promise<string>} [deps.describeError]
 * @param {(id, update) => void} [deps.onSettled]  pane-side card update
 */
export function createOfficeRunner(deps) {
  const {
    send,
    executeTool,
    prepareMutationRecovery,
    takeMutationDiff,
    runWorkbookWrite,
    refuseInCellEditMode: probe = refuseInCellEditMode,
    describeError = describeOfficeToolError,
    onSettled = () => {},
    log = console,
  } = deps;
  const cancelled = new Set();

  function settleCancelled(id) {
    cancelled.delete(id);
    send({ type: "tool_settled", id });
  }

  // The daemon stopped waiting (timeout, Stop, session end). A call that has not
  // started never starts; one already in Excel cannot be interrupted, so its late
  // result is dropped and the daemon is told it settled.
  function cancel(id, forgetAfterMs = 65_000) {
    cancelled.add(id);
    setTimeout(() => cancelled.delete(id), forgetAfterMs);
  }

  async function run({ id, name, args }) {
    if (cancelled.has(id)) {
      settleCancelled(id);
      return;
    }
    const mutation = isMutationCall(name, args);
    let commitRecovery = null;
    let committedRecovery = null;
    const cancelledBeforeExecution = new Error("Tool call cancelled before execution");
    try {
      const execute = async () => {
        if (cancelled.delete(id)) throw cancelledBeforeExecution;
        if (mutation) await probe();
        commitRecovery = WRITE_TOOLS.has(name)
          ? await prepareMutationRecovery(name, args, id)
          : null;
        if (cancelled.delete(id)) throw cancelledBeforeExecution;
        let result = await executeTool(name, args, id);
        // Office.js cannot interrupt a context.sync already in progress, but a
        // daemon timeout/session stop must prevent a late result from being
        // mistaken for the current turn's result.
        if (cancelled.delete(id)) {
          // Cancelled from the daemon's side, but the write did complete.
          if (commitRecovery) committedRecovery = await settleRecovery(commitRecovery, result);
          return CANCELLED;
        }
        if (commitRecovery) {
          committedRecovery = await settleRecovery(commitRecovery, result);
          result = { ...result, recovery: committedRecovery };
        }
        if (cancelled.delete(id)) return CANCELLED;
        return withMutationReceipt(name, args, result, id);
      };

      let result;
      if (mutation) {
        const coordinated = await runWorkbookWrite(id, name, execute);
        if (coordinated.result === CANCELLED) return settleCancelled(id);
        result = { ...coordinated.result, workbookRevision: coordinated.revision };
      } else {
        result = await execute();
        if (result === CANCELLED) return settleCancelled(id);
      }
      onSettled(id, {
        ok: true,
        name,
        args,
        result,
        diff: mutation ? takeMutationDiff(id) : undefined,
      });
      send({ type: "tool_result", id, ok: true, result });
    } catch (err) {
      if (err === cancelledBeforeExecution) return settleCancelled(id);
      // The failure path commits once, and tells the checkpoint the write failed:
      // a range snapshot is still true, but a structure inverse would not be.
      const recovery =
        committedRecovery ??
        (commitRecovery && mutation
          ? await settleRecovery(commitRecovery, undefined, {
              commitStatus: err?.commitStatus ?? "unknown",
            })
          : null);
      takeMutationDiff(id); // failed call: drop the pane-side diff
      if (cancelled.has(id)) return settleCancelled(id);
      log.error?.(`[tool ${name}] failed:`, err);
      const error = {
        message: await describeError(err, args),
        ...(err?.code ? { code: err.code } : {}),
        ...(err?.commitStatus
          ? { commitStatus: err.commitStatus }
          : mutation
            ? { commitStatus: "unknown" }
            : {}),
        ...(recovery ? { recovery } : {}),
      };
      onSettled(id, { ok: false, name, args, error });
      send({ type: "tool_result", id, ok: false, error });
    }
  }

  return { run, cancel };
}
