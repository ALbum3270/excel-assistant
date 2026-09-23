import { randomUUID } from "node:crypto";

const OFFICE_WRITE_TOOLS = new Set(
  [
    "set_cell_range",
    "fill_formula",
    "copy_to",
    "clear_cell_range",
    "modify_sheet_structure",
    "modify_workbook_structure",
    "resize_range",
    "modify_object",
    "set_format",
    "sort_range",
    "autofilter",
    "create_table",
    "add_table_rows",
  ].map((name) => `mcp__office__excel_${name}`),
);

// ThepExcel groups many operations behind one tool name. Only actions known
// to leave the workbook unchanged bypass approval; unknown/new actions remain
// approval-gated until they are classified.
const THEPEXCEL_READ_ACTIONS = new Map(
  Object.entries({
    excel_workbook: ["list", "info"],
    excel_sheet: ["list"],
    excel_range: ["read", "read_spill"],
    excel_table: ["list", "read"],
    excel_powerquery: ["list", "get", "analyze", "analyze_raw", "get_parameter", "list_parameters"],
    excel_pivot: ["list", "read"],
    excel_datamodel: [
      "info",
      "list_tables",
      "list_relationships",
      "list_measures",
      "cube_value",
      "cube_member",
    ],
    excel_name: ["list", "get"],
    excel_chart: ["list", "export_image"],
    excel_screenshot: ["range", "sheet", "chart"],
    excel_shape: ["list"],
    excel_slicer: ["list"],
    excel_sparkline: ["list"],
    excel_validation: ["list"],
    excel_page_setup: ["get", "export_pdf"],
    excel_comment: ["list", "get"],
    excel_hyperlink: ["list"],
    excel_protection: ["status"],
    excel_find_replace: ["find", "count"],
    excel_diff: ["ranges", "sheets"],
    excel_snapshot: ["snapshot", "list", "delete"],
    excel_vba: ["list_modules", "get_module"],
  }).map(([tool, actions]) => [tool, new Set(actions)]),
);

export function needsApproval(toolName, input) {
  if (OFFICE_WRITE_TOOLS.has(toolName)) return true;
  if (toolName === "mcp__office__excel_workbook_history") {
    return Boolean(input?.action && input.action !== "list");
  }
  // excel_bash is not gated here: its one write path, csv-to-sheet, asks for
  // approval itself at the moment it writes (see compute-tool.mjs).
  if (toolName === "mcp__office__excel_bash") return false;

  const prefix = "mcp__thepexcel-excel__";
  if (!toolName.startsWith(prefix)) return false;
  const readActions = THEPEXCEL_READ_ACTIONS.get(toolName.slice(prefix.length));
  const action = String(input?.action ?? "")
    .trim()
    .toLowerCase();
  return !readActions?.has(action);
}

export class ApprovalManager {
  constructor({ sendEvent, timeoutMs = 10 * 60_000 }) {
    this.sendEvent = sendEvent;
    this.timeoutMs = timeoutMs;
    this.enabledByKey = new Map();
    this.pending = new Map();
  }

  isEnabled(key) {
    return this.enabledByKey.get(key) === true;
  }

  setEnabled(key, enabled) {
    this.enabledByKey.set(key, Boolean(enabled));
    if (!enabled) this.settleKey(key, "disabled");
  }

  clearKey(key) {
    this.enabledByKey.delete(key);
    this.settleKey(key, "cancelled");
  }

  settleKey(key, decision) {
    for (const pending of this.pending.values()) {
      if (pending.key === key) pending.settle(decision);
    }
  }

  respond(key, requestId, decision) {
    const pending = this.pending.get(requestId);
    if (!pending || pending.key !== key) return false;
    const normalized = decision === "approve" || decision === "approve_turn" ? decision : "reject";
    pending.settle(normalized);
    return true;
  }

  replay(key) {
    for (const pending of this.pending.values()) {
      if (pending.key === key) this.sendEvent(pending.event, key);
    }
  }

  request(key, session, toolName, input) {
    if (session?.approveRestOfTurn) return Promise.resolve("approve");
    const requestId = randomUUID();
    const event = {
      event: "approval_request",
      request_id: requestId,
      tool: toolName.replace(/^mcp__[^_]+__/, ""),
      input,
    };

    return new Promise((resolve) => {
      const settle = (decision) => {
        if (!this.pending.delete(requestId)) return;
        clearTimeout(timer);
        session?.abortController?.signal.removeEventListener("abort", onAbort);
        this.sendEvent({ event: "approval_resolved", request_id: requestId, decision }, key);
        resolve(decision);
      };
      const onAbort = () => settle("cancelled");
      const timer = setTimeout(() => settle("timeout"), this.timeoutMs);
      session?.abortController?.signal.addEventListener("abort", onAbort, { once: true });
      this.pending.set(requestId, { key, event, settle });
      this.sendEvent(event, key);
    });
  }
}
