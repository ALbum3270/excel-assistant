// Office tools are always loaded in the SDK request. Searching only for those
// names cannot reveal a new schema, and one evaluated turn spent 17 calls
// repeatedly searching them without ever making its intended write.
export function guardOfficeToolSearch(input) {
  const query = input?.tool_input?.query;
  if (input?.tool_name !== "ToolSearch" || typeof query !== "string") return {};
  const match = /^select:(.+)$/.exec(query.trim());
  if (!match) return {};
  const names = match[1].split(",").map((name) => name.trim());
  if (!names.length || names.some((name) => !/^mcp__office__excel_[a-z_]+$/.test(name))) {
    return {};
  }
  return {
    hookSpecificOutput: {
      hookEventName: "PreToolUse",
      permissionDecision: "deny",
      permissionDecisionReason:
        "These Office Excel tools are already loaded. Call the named mcp__office__excel_* tool directly with its arguments; repeating ToolSearch will not load anything new.",
    },
  };
}
