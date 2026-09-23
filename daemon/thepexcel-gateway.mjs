import { win32 } from "node:path";
import { createSdkMcpServer, tool } from "@anthropic-ai/claude-agent-sdk";
import { Client } from "@modelcontextprotocol/sdk/client/index.js";
import {
  StdioClientTransport,
  getDefaultEnvironment,
} from "@modelcontextprotocol/sdk/client/stdio.js";
import { z } from "zod";
import { needsApproval } from "./approval.mjs";
import { canonicalWorkbookId } from "./workbook-execution.mjs";
import { backupWorkbookFile } from "./com-backup.mjs";

const UPSTREAM_TIMEOUT_MS = 10 * 60_000;
const UNSUPPORTED_WORKBOOK_ACTIONS = new Set(["open", "create", "save_as", "close"]);

function errorResult(error) {
  return {
    content: [
      {
        type: "text",
        text: JSON.stringify({
          success: false,
          error: error?.message ?? String(error),
          ...(error?.code ? { code: error.code } : {}),
          ...(error?.commitStatus ? { commitStatus: error.commitStatus } : {}),
          ...(Number.isInteger(error?.workbookRevision)
            ? { workbookRevision: error.workbookRevision }
            : {}),
        }),
      },
    ],
    isError: true,
  };
}

function schemaShape(jsonSchema) {
  const required = new Set(jsonSchema?.required ?? []);
  const shape = {};
  for (const [name, schema] of Object.entries(jsonSchema?.properties ?? {})) {
    let field = z.fromJSONSchema(schema);
    if (!required.has(name)) field = field.optional();
    shape[name] = field;
  }
  return shape;
}

function parseTextResult(result) {
  for (const item of result?.content ?? []) {
    if (item.type !== "text") continue;
    try {
      return JSON.parse(item.text);
    } catch {}
  }
  return null;
}

function isUpstreamTimeout(result) {
  return Boolean(
    result?.isError &&
    (result.content ?? []).some(
      (item) => item.type === "text" && /timed? out|timeout/i.test(item.text),
    ),
  );
}

function targetArguments(name, args, workbookName, supportsWorkbook) {
  const targeted = { ...args };
  delete targeted.expected_revision;
  if (name === "excel_workbook") {
    const action = String(targeted.action ?? "").toLowerCase();
    if (UNSUPPORTED_WORKBOOK_ACTIONS.has(action)) {
      throw Object.assign(
        new Error(
          `excel_workbook action '${action}' changes the task pane's workbook boundary and is not available through the coordinated gateway.`,
        ),
        { code: "WORKBOOK_BOUNDARY_CHANGE", commitStatus: "not_committed" },
      );
    }
    if (action !== "list") targeted.workbook = workbookName;
    return targeted;
  }
  if (name === "excel_diff") {
    targeted.left_workbook = workbookName;
    targeted.right_workbook = workbookName;
    return targeted;
  }
  if (name === "excel_snapshot" && String(targeted.action).toLowerCase() === "restore") {
    throw Object.assign(
      new Error(
        "Snapshot restore opens another workbook and is not available through the active-workbook gateway.",
      ),
      { code: "WORKBOOK_BOUNDARY_CHANGE", commitStatus: "not_committed" },
    );
  }
  if (supportsWorkbook) targeted.workbook = workbookName;
  return targeted;
}

function resultWithRevision(result, revision) {
  const { comBackup, ...rest } = result ?? {};
  return {
    ...rest,
    content: [
      ...(result?.content ?? []),
      {
        type: "text",
        text: JSON.stringify({
          workbookRevision: revision,
          coordinated: true,
          ...(comBackup
            ? {
                fileBackup:
                  comBackup.status === "skipped"
                    ? { taken: false, reason: comBackup.reason }
                    : {
                        taken: true,
                        file: comBackup.file,
                        savedState: comBackup.savedAt,
                        note: "A copy of the workbook file as it was last saved; edits not yet saved are not in it.",
                      },
              }
            : {}),
        }),
      },
    ],
  };
}

// `client` lets a test drive the gateway with an in-memory MCP client instead
// of spawning the COM server; production always passes a stdio config.
export async function createThepExcelGateway(config, execution, { client: injected } = {}) {
  if (!injected && (!config || config.type !== "stdio" || !config.command)) return null;
  let transport = null;
  let client = injected;
  if (!client) {
    transport = new StdioClientTransport({
      command: config.command,
      args: config.args ?? [],
      cwd: config.cwd,
      env: { ...getDefaultEnvironment(), ...(config.env ?? {}) },
      stderr: "inherit",
    });
    client = new Client({ name: "excel-assistant-workbook-gateway", version: "0.2.0" });
    await client.connect(transport);
  }
  const listed = await client.listTools({}, { timeout: 30_000 });
  const upstreamTools = listed.tools;

  async function callUpstream(name, args, signal) {
    return client.callTool({ name, arguments: args }, undefined, {
      timeout: UPSTREAM_TIMEOUT_MS,
      maxTotalTimeout: UPSTREAM_TIMEOUT_MS,
      ...(signal ? { signal } : {}),
    });
  }

  // A call that was stopped or timed out while it waited on the workbook check
  // or the backup must not go on to write: the caller has already been told it
  // was cancelled. Checked before every step that could start one.
  function stopIfCancelled(runSignal) {
    if (!runSignal?.aborted) return;
    throw Object.assign(new Error("The COM operation was cancelled before it started writing."), {
      code: "TOOL_CANCELLED",
      commitStatus: "not_committed",
    });
  }

  return {
    createSessionServer({ workbookPath, signal, revisionState = { value: undefined } }) {
      const workbookId = canonicalWorkbookId(workbookPath);
      const workbookName = win32.basename(workbookId);
      const hasStableTarget = /^[a-z]:\\|^\\\\/i.test(workbookId);

      const tools = upstreamTools.map((definition) => {
        const shape = schemaShape(definition.inputSchema);
        const supportsWorkbook = Object.hasOwn(
          definition.inputSchema?.properties ?? {},
          "workbook",
        );
        shape.expected_revision = z
          .number()
          .int()
          .nonnegative()
          .optional()
          .describe(
            "Workbook revision observed by a preceding coordinated read. Usually supplied automatically.",
          );
        return tool(
          definition.name,
          definition.description ?? definition.name,
          shape,
          async (rawArgs) => {
            const write = needsApproval(`mcp__thepexcel-excel__${definition.name}`, rawArgs);
            try {
              if (!hasStableTarget) {
                throw Object.assign(
                  new Error("COM tools require the active workbook to have a local file path."),
                  { code: "WORKBOOK_TARGET_UNAVAILABLE", commitStatus: "not_committed" },
                );
              }
              const expectedRevision = Number.isInteger(rawArgs.expected_revision)
                ? rawArgs.expected_revision
                : write
                  ? revisionState.value
                  : undefined;
              const completed = await execution.run(
                workbookId,
                {
                  write,
                  expectedRevision,
                  signal,
                  timeoutMs: 60_000,
                  toolName: definition.name,
                },
                async ({ signal: runSignal } = {}) => {
                  const args = targetArguments(
                    definition.name,
                    rawArgs,
                    workbookName,
                    supportsWorkbook,
                  );
                  const isUntargetedList =
                    definition.name === "excel_workbook" &&
                    String(args.action).toLowerCase() === "list";
                  if (!isUntargetedList) {
                    let info;
                    try {
                      info = await callUpstream(
                        "excel_workbook",
                        { action: "info", workbook: workbookName },
                        runSignal,
                      );
                    } catch (error) {
                      // MCP rejects immediately when its signal is aborted.
                      // This request only read workbook identity; no write was
                      // dispatched, so it must not consume a revision or lock.
                      throw Object.assign(error, {
                        commitStatus: "not_committed",
                        executionSettled: true,
                      });
                    }
                    stopIfCancelled(runSignal);
                    if (info.isError)
                      return { ...info, success: false, commitStatus: "not_committed" };
                    const actualPath = parseTextResult(info)?.path;
                    if (!actualPath || canonicalWorkbookId(actualPath) !== workbookId) {
                      throw Object.assign(
                        new Error(
                          `COM resolved '${workbookName}' to a different workbook. Expected ${workbookPath}; got ${actualPath || "no path"}.`,
                        ),
                        { code: "WORKBOOK_TARGET_MISMATCH", commitStatus: "not_committed" },
                      );
                    }
                  }
                  // COM writes cannot be undone from a range snapshot, so keep
                  // a copy of the file first. It is the last saved state.
                  let backup = null;
                  if (write) {
                    try {
                      backup = await backupWorkbookFile(workbookId);
                    } catch (error) {
                      throw Object.assign(
                        new Error(
                          `The workbook could not be backed up before this COM write, so it was not run: ${error?.message ?? String(error)}`,
                        ),
                        { code: "COM_BACKUP_FAILED", commitStatus: "not_committed" },
                      );
                    }
                  }
                  // Once dispatched, a COM write cannot be interrupted, so the
                  // last cancellation check is here, right before it.
                  stopIfCancelled(runSignal);
                  const result = await callUpstream(
                    definition.name,
                    args,
                    write ? undefined : runSignal,
                  );
                  if (backup) result.comBackup = backup;
                  if (write && isUpstreamTimeout(result)) {
                    throw Object.assign(
                      new Error("The COM operation timed out and may still be running in Excel."),
                      { code: "TOOL_TIMEOUT", commitStatus: "unknown" },
                    );
                  }
                  return result;
                },
              );
              revisionState.value = completed.revision;
              return resultWithRevision(completed.result, completed.revision);
            } catch (error) {
              return errorResult(error);
            }
          },
          { annotations: definition.annotations },
        );
      });

      return createSdkMcpServer({
        name: "thepexcel-excel",
        version: "0.2.0",
        instructions: client.getInstructions(),
        timeout: UPSTREAM_TIMEOUT_MS,
        tools,
      });
    },
    toolCount: upstreamTools.length,
    close: () => transport?.close(),
  };
}
