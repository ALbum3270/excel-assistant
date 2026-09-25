import { test } from "node:test";
import assert from "node:assert/strict";
import {
  buildOverview,
  ChangeTracker,
  createWorkbookCoordinator,
  readSelectionContext,
} from "../taskpane/shared/vendor/pi-context.js";
import { createWorkbookExecution } from "../daemon/workbook-execution.mjs";

test("workbook mutations run in order and receive monotonic revisions", async () => {
  const coordinator = createWorkbookCoordinator();
  const order = [];
  let releaseFirst;
  const firstGate = new Promise((resolve) => {
    releaseFirst = resolve;
  });
  const context = (opId) => ({ workbookId: "book", sessionId: "session", opId });

  const first = coordinator.runWrite(context("first"), async () => {
    order.push("first:start");
    await firstGate;
    order.push("first:end");
    return "one";
  });
  const second = coordinator.runWrite(context("second"), async () => {
    order.push("second");
    return "two";
  });

  await Promise.resolve();
  assert.deepEqual(order, ["first:start"]);
  releaseFirst();
  const [one, two] = await Promise.all([first, second]);
  assert.deepEqual(order, ["first:start", "first:end", "second"]);
  assert.deepEqual([one.revision, two.revision], [1, 2]);
});

test("daemon coordination serializes reads with writes and rejects a stale write", async () => {
  const execution = createWorkbookExecution();
  const order = [];
  let releaseRead;
  const gate = new Promise((resolve) => (releaseRead = resolve));
  const read = execution.run("C:\\Books\\Model.xlsx", { toolName: "read" }, async () => {
    order.push("read:start");
    await gate;
    order.push("read:end");
    return "read";
  });
  const write = execution.run(
    "c:\\books\\model.xlsx",
    { write: true, expectedRevision: 0, toolName: "write" },
    async () => {
      order.push("write");
      return "write";
    },
  );
  await Promise.resolve();
  assert.deepEqual(order, ["read:start"]);
  releaseRead();
  const [, written] = await Promise.all([read, write]);
  assert.deepEqual(order, ["read:start", "read:end", "write"]);
  assert.equal(written.revision, 1);
  await assert.rejects(
    execution.run(
      "C:\\BOOKS\\MODEL.XLSX",
      { write: true, expectedRevision: 0, toolName: "stale" },
      async () => "must not run",
    ),
    (error) => error.code === "STALE_WORKBOOK_REVISION" && error.workbookRevision === 1,
  );
});

test("a write the pane refuses does not consume a revision", async () => {
  const execution = createWorkbookExecution();
  const path = "C:\Books\Guarded.xlsx";
  const first = await execution.run(path, { write: true, toolName: "write" }, async () => ({
    success: true,
  }));
  assert.equal(first.revision, 1);

  // The overwrite guard reports its refusal in the result; nothing was written.
  const refused = await execution.run(
    path,
    { write: true, expectedRevision: 1, toolName: "write" },
    async () => ({
      success: false,
      error: "Would overwrite 5 non-empty cell(s)",
    }),
  );
  assert.equal(refused.revision, 1);

  // So the caller's expectation is still current and the next write runs.
  const second = await execution.run(
    path,
    { write: true, expectedRevision: 1, toolName: "write" },
    async () => ({
      success: true,
    }),
  );
  assert.equal(second.revision, 2);
});

test("a write with an unknown outcome consumes its revision", async () => {
  const execution = createWorkbookExecution();
  const path = "C:\Books\Uncertain.xlsx";
  await assert.rejects(
    execution.run(path, { write: true, toolName: "write" }, async () => {
      throw Object.assign(new Error("sync never returned"), { commitStatus: "unknown" });
    }),
    (error) => error.workbookRevision === 1,
  );
});

test("a write that fails part-way still consumes its revision", async () => {
  const execution = createWorkbookExecution();
  const path = "C:\Books\Partial.xlsx";
  // The pane returns every error settled; this one had already changed cells.
  await assert.rejects(
    execution.run(path, { write: true, toolName: "write" }, async () => {
      throw Object.assign(new Error("format sync failed after values were written"), {
        commitStatus: "unknown",
        executionSettled: true,
      });
    }),
    (error) => error.workbookRevision === 1,
  );
  assert.equal(execution.snapshot(path).blocked, null, "a settled error does not block the queue");
  let staleRan = false;
  await assert.rejects(
    execution.run(path, { write: true, expectedRevision: 0, toolName: "stale" }, async () => {
      staleRan = true;
    }),
    (error) =>
      error.code === "STALE_WORKBOOK_REVISION" &&
      /write operation ended without a confirmed commit/.test(error.message) &&
      !/did not come from this session/.test(error.message),
  );
  assert.equal(staleRan, false);

  // A refusal the pane marks not committed still leaves the number alone.
  await assert.rejects(
    execution.run(path, { write: true, toolName: "write" }, async () => {
      throw Object.assign(new Error("cell-edit mode"), {
        commitStatus: "not_committed",
        executionSettled: true,
      });
    }),
    (error) => error.workbookRevision === 1,
  );
});

test("a timed-out write keeps the shared queue until the executor settles", async () => {
  const execution = createWorkbookExecution();
  let releaseWrite;
  const gate = new Promise((resolve) => (releaseWrite = resolve));
  const timedOut = execution.run(
    "book.xlsx",
    { write: true, timeoutMs: 10, toolName: "slow-write" },
    async () => {
      await gate;
      return "done";
    },
  );
  await assert.rejects(timedOut, (error) => error.code === "TOOL_TIMEOUT");
  let nextStarted = false;
  const next = execution.run("book.xlsx", { toolName: "read-after" }, async () => {
    nextStarted = true;
    return "next";
  });
  await new Promise((resolve) => setTimeout(resolve, 15));
  assert.equal(nextStarted, false);
  releaseWrite();
  await next;
  assert.equal(nextStarted, true);
});

test("selection context reads the address submitted with the turn", async () => {
  const originalExcel = globalThis.Excel;
  const requestedRanges = [];
  const sheetName = "O'Brien!2026";
  const selectedRange = {
    address: "'O''Brien!2026'!B2",
    rowIndex: 1,
    columnIndex: 1,
    rowCount: 1,
    columnCount: 1,
    worksheet: { name: sheetName },
    load() {},
    getCell() {
      return { getResizedRange: () => ({ formulas: [["=1+2"]], load() {} }) };
    },
  };
  const sheet = {
    getUsedRangeOrNullObject() {
      return {
        isNullObject: false,
        rowIndex: 0,
        columnIndex: 0,
        rowCount: 3,
        columnCount: 3,
        load() {},
      };
    },
    getRange(address) {
      requestedRanges.push(address);
      if (address === "B2") return selectedRange;
      if (address === "A1:C3") {
        return {
          address,
          values: [
            ["Header", "Value", "Other"],
            [1, 3, ""],
            [2, "#DIV/0!", ""],
          ],
          load() {},
        };
      }
      throw new Error(`Unexpected range ${address}`);
    },
  };
  globalThis.Excel = {
    async run(callback) {
      return callback({
        workbook: {
          getSelectedRange() {
            throw new Error("Live selection must not be read");
          },
          worksheets: {
            getItem(name) {
              assert.equal(name, sheetName);
              return sheet;
            },
          },
        },
        async sync() {},
      });
    },
  };
  try {
    const result = await readSelectionContext("'O''Brien!2026'!B2");
    assert.deepEqual(requestedRanges, ["B2", "A1:C3"]);
    assert.match(result.text, /Selection:.*B2/);
    assert.match(result.text, /B2: =1\+2/);
    assert.match(result.text, /Errors nearby:.*B3=#DIV\/0!/);
  } finally {
    globalThis.Excel = originalExcel;
  }
});

test("overview bounds the header read and counts shapes without loading them", async () => {
  const originalExcel = globalThis.Excel;
  const ranges = [];
  const sheet = {
    name: "Data",
    position: 0,
    visibility: "Visible",
    getUsedRangeOrNullObject() {
      return { isNullObject: false, rowCount: 2, columnCount: 30, load() {} };
    },
    getRange(address) {
      ranges.push(address);
      return {
        getUsedRangeOrNullObject: () => ({ isNullObject: false, values: [["A", "B"]], load() {} }),
      };
    },
    tables: { items: [], load() {} },
    charts: { count: 0, load() {} },
    pivotTables: { getCount: () => ({ value: 0 }) },
    shapes: { getCount: () => ({ value: 2 }) },
  };
  globalThis.Excel = {
    async run(callback) {
      return callback({
        workbook: {
          name: "Book.xlsx",
          load() {},
          worksheets: { items: [sheet], load() {} },
          names: { items: [], load() {} },
        },
        async sync() {},
      });
    },
  };
  try {
    const overview = await buildOverview();
    assert.deepEqual(ranges, ["A1:T1"]);
    assert.match(overview, /2 shape\(s\)/);
  } finally {
    globalThis.Excel = originalExcel;
  }
});

test("change tracker keeps user edits and skips this add-in's own writes", async () => {
  const originalExcel = globalThis.Excel;
  let onChanged;
  const sheet = {
    id: "s1",
    name: "Data",
    visibility: "Visible",
    onChanged: { add: (handler) => (onChanged = handler) },
  };
  globalThis.Excel = {
    async run(callback) {
      return callback({ workbook: { worksheets: { items: [sheet], load() {} } }, async sync() {} });
    },
  };
  try {
    const tracker = new ChangeTracker();
    await tracker.start();
    await onChanged({ worksheetId: "s1", address: "B2", triggerSource: "Unknown" });
    await onChanged({ worksheetId: "s1", address: "C3", triggerSource: "ThisLocalAddin" });
    await onChanged({ worksheetId: "s1", address: "D4" });
    const flushed = tracker.flush();
    assert.match(flushed, /Data: B2, D4/);
    assert.doesNotMatch(flushed, /C3/);
  } finally {
    globalThis.Excel = originalExcel;
  }
});

test("a stale write says what changed the workbook", async () => {
  const execution = createWorkbookExecution();
  const path = "C:\Books\Stale.xlsx";
  await execution.run(path, { write: true, toolName: "excel_workbook_history" }, async () => ({
    success: true,
  }));
  await assert.rejects(
    execution.run(
      path,
      { write: true, expectedRevision: 0, toolName: "excel_set_cell_range" },
      async () => "never",
    ),
    (error) => /restore from the panel's backups/.test(error.message),
  );
});
