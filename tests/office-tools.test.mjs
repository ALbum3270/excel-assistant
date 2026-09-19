import { test } from "node:test";
import assert from "node:assert/strict";
import { createOfficeBridgeMcp } from "../daemon/office-tools.mjs";

function schemas() {
  const server = createOfficeBridgeMcp({ callTaskpaneTool() {} }, "excel", "test-pane");
  const tools = server.instance._registeredTools;
  return Object.fromEntries(
    Object.entries(tools).map(([name, registration]) => [name, registration.inputSchema]),
  );
}

test("set-cell schema rejects empty, ragged and malformed formula matrices", () => {
  const schema = schemas().excel_set_cell_range;
  const base = { sheetId: 1, range: "A1" };
  assert.equal(schema.safeParse({ ...base, cells: [] }).success, false);
  assert.equal(
    schema.safeParse({ ...base, cells: [[{ value: 1 }], [{ value: 2 }, { value: 3 }]] }).success,
    false,
  );
  assert.equal(schema.safeParse({ ...base, cells: [[{ formula: "SUM(A1:A2)" }]] }).success, false);
  assert.equal(
    schema.safeParse({ ...base, cells: [[{ cellStyles: { fontWeight: "bold" } }]] }).success,
    true,
  );
});

test("read and structure schemas enforce positive bounds", () => {
  const all = schemas();
  assert.equal(all.excel_get_selected_range.safeParse({ cellLimit: 0 }).success, false);
  assert.equal(all.excel_get_selected_range.safeParse({ cellLimit: 5000 }).success, true);
  assert.equal(all.excel_get_selected_range.safeParse({ cellLimit: 5001 }).success, false);
  assert.equal(
    all.excel_get_cell_ranges.safeParse({ sheetId: 1, ranges: [], cellLimit: 0 }).success,
    false,
  );
  assert.equal(
    all.excel_get_cell_ranges.safeParse({ sheetId: 1, ranges: ["A1"], cellLimit: 1 }).success,
    true,
  );
  assert.equal(
    all.excel_modify_sheet_structure.safeParse({
      sheetId: 1,
      operation: "insert",
      dimension: "rows",
      reference: "1",
      count: 0,
    }).success,
    false,
  );
});

test("copy schema exposes explicit overwrite authorization", () => {
  const schema = schemas().excel_copy_to;
  const parsed = schema.parse({
    sheetId: 1,
    sourceRange: "A1",
    destinationRange: "B1:B3",
    allow_overwrite: true,
  });
  assert.equal(parsed.allow_overwrite, true);
});

test("table-row schema rejects empty and ragged matrices before dispatch", () => {
  const schema = schemas().excel_add_table_rows;
  assert.equal(schema.safeParse({ table: "T", values: [] }).success, false);
  assert.equal(schema.safeParse({ table: "T", values: [[]] }).success, false);
  assert.equal(schema.safeParse({ table: "T", values: [[1, 2], [3]] }).success, false);
  assert.equal(schema.safeParse({ table: "T", values: [[1]], index: -1 }).success, false);
  assert.equal(
    schema.safeParse({
      table: "T",
      values: [
        [1, 2],
        [3, 4],
      ],
    }).success,
    true,
  );
});

test("set-cell normalizes model-shaped payloads and fills a larger target", async () => {
  const calls = [];
  const server = createOfficeBridgeMcp(
    {
      async callTaskpaneTool(name, args) {
        calls.push({ name, args });
        return { success: true, commitStatus: "committed" };
      },
    },
    "excel",
    "test-pane",
  );
  const handler = server.instance._registeredTools.excel_set_cell_range.handler;

  await handler({
    sheetId: 1,
    range: "C2:C3",
    cells: '[{"formula":"=A2+B2"},{"formula":"=A3+B3"}]',
  });
  await handler({ sheetId: 1, range: "F2:F15", cells: [{ formula: "=B2" }] });

  assert.deepEqual(calls[0], {
    name: "excel_set_cell_range",
    args: {
      sheetId: 1,
      range: "C2:C3",
      cells: [[{ formula: "=A2+B2" }], [{ formula: "=A3+B3" }]],
    },
  });
  assert.deepEqual(calls[1], {
    name: "excel_set_cell_range",
    args: {
      sheetId: 1,
      range: "F2",
      copyToRange: "F2:F15",
      cells: [[{ formula: "=B2" }]],
    },
  });
});

test("range reads accept a model-supplied bracketed string", async () => {
  const calls = [];
  const server = createOfficeBridgeMcp(
    {
      async callTaskpaneTool(name, args) {
        calls.push({ name, args });
        return { success: true };
      },
    },
    "excel",
    "test-pane",
  );

  await server.instance._registeredTools.excel_get_cell_ranges.handler({
    sheetId: 1,
    ranges: "\n[A1:E15]\n",
  });

  assert.deepEqual(calls[0], {
    name: "excel_get_cell_ranges",
    args: { sheetId: 1, ranges: ["A1:E15"], includeStyles: false },
  });
});

test("set-cell recovers unescaped formula JSON and expands a row pattern", async () => {
  const calls = [];
  const server = createOfficeBridgeMcp(
    {
      async callTaskpaneTool(name, args) {
        calls.push({ name, args });
        return { success: true, commitStatus: "committed" };
      },
    },
    "excel",
    "test-pane",
  );
  const handler = server.instance._registeredTools.excel_set_cell_range.handler;

  await handler({
    sheetId: 1,
    range: "C2:D11",
    cells: '[["=IF(A2="","",A2)"], ["=IF(B2="","",B2)"]]',
    allow_overwrite: true,
  });

  assert.deepEqual(calls[0], {
    name: "excel_set_cell_range",
    args: {
      sheetId: 1,
      range: "C2:D2",
      cells: [[{ formula: '=IF(A2="","",A2)' }, { formula: '=IF(B2="","",B2)' }]],
      copyToRange: "C2:D11",
      allow_overwrite: true,
    },
  });
});

test("fill-formula sends one formula and a translated fill range", async () => {
  const calls = [];
  const server = createOfficeBridgeMcp(
    {
      async callTaskpaneTool(name, args) {
        calls.push({ name, args });
        return { success: true, commitStatus: "committed" };
      },
    },
    "excel",
    "test-pane",
  );

  await server.instance._registeredTools.excel_fill_formula.handler({
    sheetId: 1,
    range: "F2:F5000",
    formula: "=SUM(B2:E2)",
    allow_overwrite: true,
  });

  assert.deepEqual(calls[0], {
    name: "excel_set_cell_range",
    args: {
      sheetId: 1,
      range: "F2",
      copyToRange: "F2:F5000",
      cells: [[{ formula: "=SUM(B2:E2)" }]],
      allow_overwrite: true,
    },
  });
});
