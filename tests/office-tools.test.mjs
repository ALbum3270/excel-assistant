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

test("write errors expose a recovery checkpoint to the model", async () => {
  const error = Object.assign(new Error("write result unknown"), {
    commitStatus: "unknown",
    recovery: { status: "checkpoint_created", snapshotIds: ["before-write"] },
  });
  const server = createOfficeBridgeMcp(
    { async callTaskpaneTool() { throw error; } },
    "excel",
    "test-pane",
  );
  const result = await server.instance._registeredTools.excel_clear_cell_range.handler({
    sheetId: 1,
    range: "A1",
  });
  const payload = JSON.parse(result.content[0].text);
  assert.equal(result.isError, true);
  assert.equal(payload.commitStatus, "unknown");
  assert.deepEqual(payload.recovery, { status: "checkpoint_created", snapshotIds: ["before-write"] });
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

test("set-cell refuses ambiguous payloads and never repeats plain values", async () => {
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

  const malformed = await handler({ sheetId: 1, range: "C2:D11", cells: '[["=IF(A2="","",A2)"]]' });
  const flat = await handler({ sheetId: 1, range: "A1:B2", cells: [1, 2, 3, 4] });
  // A column sent for a row target (seen in 183-8) must not land in J3:J5.
  const spill = await handler({ sheetId: 1, range: "J3:L3", cells: [["=D3"], ["=E3"], ["=F3"]] });
  await handler({ sheetId: 1, range: "A1:D1", cells: [["Total"]] });

  assert.equal(malformed.isError, true);
  assert.match(malformed.content[0].text, /not valid JSON/);
  assert.equal(flat.isError, true);
  assert.match(flat.content[0].text, /2D array of rows/);
  assert.equal(spill.isError, true);
  assert.match(spill.content[0].text, /3x1 but range J3:L3 is 1x3/);
  assert.deepEqual(calls, [
    { name: "excel_set_cell_range", args: { sheetId: 1, range: "A1:D1", cells: [[{ value: "Total" }]] } },
  ]);
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

test("write receipts keep a bounded sample of formula results and errors", async () => {
  const formulaResults = Object.fromEntries(Array.from({ length: 4999 }, (_, i) => [`F${i + 2}`, i]));
  const formulaErrors = Array.from({ length: 60 }, (_, i) => ({ address: `F${i + 2}`, value: "#N/A" }));
  const server = createOfficeBridgeMcp(
    {
      async callTaskpaneTool() {
        return { success: true, commitStatus: "committed", formulaResults, formulaErrors };
      },
    },
    "excel",
    "test-pane",
  );

  const result = await server.instance._registeredTools.excel_fill_formula.handler({
    sheetId: 1,
    range: "F2:F5000",
    formula: "=B2",
  });
  const receipt = JSON.parse(result.content[0].text);

  assert.equal(receipt.commitStatus, "committed");
  assert.equal(Object.keys(receipt.formulaResults).length, 20);
  assert.equal(receipt.formulaResultCount, 4999);
  assert.equal(receipt.formulaErrors.length, 50);
  assert.equal(receipt.formulaErrorCount, 60);
});
