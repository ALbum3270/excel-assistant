// What the pane calls things. Display only: nothing here is sent to the model.
// Every function reads the current language, so call them at render time.
import { t } from "./i18n.js";

// Tools with their own name in the locale files ("tool.<name>").
const NAMED_TOOLS = new Set([
  "excel_get_workbook_metadata",
  "excel_get_selected_range",
  "excel_get_cell_ranges",
  "excel_get_range_as_csv",
  "excel_search_data",
  "excel_get_all_objects",
  "excel_explain_formula",
  "excel_trace_dependencies",
  "excel_set_cell_range",
  "excel_clear_cell_range",
  "excel_copy_to",
  "excel_modify_sheet_structure",
  "excel_modify_workbook_structure",
  "excel_resize_range",
  "excel_modify_object",
  "excel_select_range",
  "excel_set_format",
  "excel_sort_range",
  "excel_autofilter",
  "excel_create_table",
  "excel_add_table_rows",
  "excel_workbook_history",
  "excel_bash",
  "Read",
  "Write",
  "Edit",
  "MultiEdit",
  "Bash",
  "Glob",
  "Grep",
  "WebFetch",
  "WebSearch",
]);

// The in-process bridge server is named "office"; to the user it is Excel.
function mcpServerDisplayName(raw) {
  if (raw === "office") return "Excel";
  return raw.charAt(0).toUpperCase() + raw.slice(1);
}

export function localToolName(name) {
  return /^mcp__[^_]+__(.+)$/.exec(name || "")?.[1] ?? name;
}

// [key, params] for the working indicator, so a stored status can be shown in
// whatever language is current when it renders.
export function statusKeyForTool(name) {
  const local = localToolName(name);
  if (NAMED_TOOLS.has(name)) return [`tool.${name}`];
  if (NAMED_TOOLS.has(local)) return [`tool.${local}`];
  const server = /^mcp__([^_]+)__/.exec(name || "");
  if (server) return ["tool.calling", { server: mcpServerDisplayName(server[1]) }];
  return ["tool.working"];
}

export function statusForTool(name) {
  const [key, params] = statusKeyForTool(name);
  return t(key, params);
}

export function toolTitle(name) {
  const local = localToolName(name);
  return NAMED_TOOLS.has(local) ? t(`tool.${local}`) : (local ?? t("tool.generic"));
}

const OPERATIONS = new Set([
  "excel_set_cell_range",
  "excel_clear_cell_range",
  "excel_copy_to",
  "excel_set_format",
  "excel_sort_range",
  "excel_resize_range",
  "excel_modify_sheet_structure",
  "excel_modify_workbook_structure",
  "excel_create_table",
  "excel_autofilter",
  "excel_add_table_rows",
  "excel_modify_object",
  "restore",
]);

export function recoveryOperationLabel(operation) {
  if (OPERATIONS.has(operation)) return t(`op.${operation}`);
  if (!operation) return t("op.fallback");
  return String(operation)
    .replace(/^excel_/, "")
    .replaceAll("_", " ");
}

// Tool arguments as labelled lines instead of a JSON dump (pi-for-excel renders
// tool parameters the same way in humanize-params.ts). The material arguments
// of each write come first. Argument names stay as the model wrote them.
const FIELDS_BY_TOOL = {
  excel_set_cell_range: [
    "sheetId",
    "range",
    "cells",
    "copyToRange",
    "resizeWidth",
    "resizeHeight",
    "allow_overwrite",
  ],
  excel_fill_formula: ["sheetId", "range", "formula", "allow_overwrite"],
  excel_copy_to: ["sheetId", "sourceRange", "destinationRange", "allow_overwrite"],
  excel_clear_cell_range: ["sheetId", "range", "clearType"],
  excel_modify_sheet_structure: [
    "sheetId",
    "operation",
    "dimension",
    "reference",
    "count",
    "position",
  ],
  excel_modify_workbook_structure: ["operation", "sheetId", "sheetName", "newName", "tabColor"],
  excel_resize_range: ["sheetId", "range", "width", "height"],
  excel_modify_object: ["sheetId", "operation", "objectType", "id", "properties"],
  excel_set_format: [
    "sheet",
    "address",
    "number_format",
    "bold",
    "italic",
    "font_size",
    "font_name",
    "font_color",
    "fill_color",
    "border",
  ],
  excel_sort_range: ["sheet", "address", "key", "ascending", "has_headers"],
  excel_autofilter: ["sheet", "address", "clear"],
  excel_create_table: ["sheet", "address", "name", "has_headers"],
  excel_add_table_rows: ["table", "values", "index"],
  excel_workbook_history: ["action", "snapshot_id"],
  excel_bash: ["command"],
};

function clip(text, max = 500) {
  return text.length > max ? text.slice(0, max - 3) + "..." : text;
}

function argValue(value) {
  if (Array.isArray(value)) {
    const rows = value.length;
    const columns = Array.isArray(value[0])
      ? Math.max(0, ...value.slice(0, 20).map((row) => row.length))
      : null;
    const shape = columns === null ? t("args.items", { count: rows }) : `${rows}×${columns}`;
    return `${shape}; ${t("args.sample")}: ${clip(JSON.stringify(value.slice(0, 3)))}`;
  }
  if (value && typeof value === "object") return clip(JSON.stringify(value));
  return clip(String(value));
}

export function describeArgs(tool, input = {}) {
  if (typeof input === "string") return input;
  const args = input && typeof input === "object" ? input : {};
  const preferred = FIELDS_BY_TOOL[localToolName(tool)] ?? ["action"];
  const fields = [...preferred, ...Object.keys(args).filter((field) => !preferred.includes(field))];
  const lines = fields
    .filter((field) => args[field] !== undefined && args[field] !== "")
    .map((field) => `${field}: ${argValue(args[field])}`);
  return lines.length ? lines.join("\n") : t("args.none");
}

// Operation names the model passes ("insert", "delete", …) shown in the pane's
// language; anything else is shown as written.
function operationWord(operation) {
  const known = ["insert", "delete", "create", "rename", "duplicate", "hide", "unhide", "update"];
  return known.includes(operation)
    ? t(`summary.op.${operation}`)
    : operation || t("summary.op.change");
}

export function mutationSummary(name, args = {}) {
  switch (name) {
    case "excel_set_cell_range": {
      const rows = Array.isArray(args.cells) ? args.cells.length : 0;
      const columns = rows ? Math.max(...args.cells.map((row) => row.length)) : 0;
      let formulas = 0;
      let values = 0;
      for (const row of args.cells || []) {
        for (const cell of row) {
          if (cell?.formula !== undefined) formulas++;
          else if (cell?.value !== undefined) values++;
        }
      }
      const parts = [
        rows && columns ? t("summary.cells", { rows, columns }) : t("summary.cellsPlain"),
      ];
      if (formulas) parts.push(t("summary.formulas", { count: formulas }));
      if (values) parts.push(t("summary.values", { count: values }));
      return parts.join(" · ");
    }
    case "excel_clear_cell_range":
      return t(
        ["all", "formats"].includes(args.clearType)
          ? `summary.clear.${args.clearType}`
          : "summary.clear.contents",
      );
    case "excel_copy_to":
      return `${args.sourceRange || t("summary.source")} → ${args.destinationRange || t("summary.destination")}`;
    case "excel_set_format":
      return t("summary.format");
    case "excel_sort_range":
      return t("summary.sort");
    case "excel_resize_range":
      return t("summary.resize");
    case "excel_modify_sheet_structure":
      return t("summary.structure", {
        operation: operationWord(args.operation),
        count: args.count || 1,
        dimension: t(args.dimension === "columns" ? "summary.columns" : "summary.rows"),
      });
    case "excel_modify_workbook_structure":
      return t("summary.sheet", { operation: operationWord(args.operation) });
    case "excel_modify_object":
      return t("summary.object", {
        operation: operationWord(args.operation),
        object: args.objectType || t("summary.objectFallback"),
      });
    case "excel_autofilter":
      return t(args.clear ? "summary.filter.clear" : "summary.filter.apply");
    case "excel_create_table":
      return t("summary.table");
    case "excel_add_table_rows":
      return t("summary.tableRows", { count: args.values?.length || 0 });
    case "excel_workbook_history":
      return t("summary.restore");
    default:
      return null;
  }
}

const GROUPED_TOOLS = new Set([
  "excel_set_cell_range",
  "excel_fill_formula",
  "excel_get_cell_ranges",
  "excel_get_range_as_csv",
  "excel_search_data",
  "excel_set_format",
  "excel_clear_cell_range",
  "excel_copy_to",
  "excel_modify_sheet_structure",
  "excel_modify_workbook_structure",
  "excel_bash",
]);

export function describeGroup(toolName, count) {
  return GROUPED_TOOLS.has(toolName)
    ? t(`group.${toolName}`, { count })
    : t("group.fallback", { count, tool: toolTitle(toolName) });
}

const DECISIONS = new Set([
  "approve",
  "approve_turn",
  "reject",
  "timeout",
  "cancelled",
  "disabled",
]);

export function decisionLabel(decision) {
  return t(DECISIONS.has(decision) ? `decision.${decision}` : "decision.closed");
}
