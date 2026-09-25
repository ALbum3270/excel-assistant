// Sandboxed computation for the agent: a just-bash shell (in-memory
// filesystem, no network, bundled CPython) plus two commands that move
// worksheet data in and out of that filesystem without passing it through
// the model context. sheet-to-csv / csv-to-sheet follow hewliyang/office-agents
// (MIT, packages/excel/src/lib/vfs/custom-commands.ts), rewired to the task
// pane tools this project already exposes.

import { Bash, defineCommand } from "just-bash";
import Papa from "papaparse";

// Same cap and notice format as vercel-labs/bash-tool (MIT).
const MAX_OUTPUT_CHARS = 30_000;
// Matches the task pane's bounded read/write chunks.
const WRITE_CHUNK_CELLS = 2000;
const CSV_PAGE_ROWS = 20_000;
const MAX_TYPED_CELLS = 200_000;
const EXCEL_MAX_ROWS = 1_048_576;
const EXCEL_MAX_COLUMNS = 16_384;
// Stay below the bridge's 60s tool deadline so an abandoned computation
// cannot continue into late workbook writes.
const COMPUTE_TIMEOUT_MS = 45_000;

function truncate(text, stream) {
  if (text.length <= MAX_OUTPUT_CHARS) return text;
  const removed = text.length - MAX_OUTPUT_CHARS;
  return `${text.slice(0, MAX_OUTPUT_CHARS)}\n\n[${stream} truncated: ${removed} characters removed]`;
}

function columnLetters(index) {
  let letters = "";
  for (let n = index + 1; n > 0; n = Math.floor((n - 1) / 26)) {
    letters = String.fromCharCode(65 + ((n - 1) % 26)) + letters;
  }
  return letters;
}

function parseCell(a1) {
  const match = /^([A-Z]+)(\d+)$/i.exec(a1);
  if (!match) return null;
  const col =
    [...match[1].toUpperCase()].reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1;
  const row = Number(match[2]) - 1;
  return row >= 0 && row < EXCEL_MAX_ROWS && col < EXCEL_MAX_COLUMNS ? { col, row } : null;
}

function parseRange(a1) {
  if (!/^[A-Z]+\d+(?::[A-Z]+\d+)?$/i.test(String(a1))) return null;
  const [first, last = first] = String(a1).split(":");
  const start = parseCell(first);
  const end = parseCell(last);
  if (!start || !end || start.row > end.row || start.col > end.col) return null;
  return { start, rows: end.row - start.row + 1, columns: end.col - start.col + 1 };
}

function overlaps(a, b) {
  return (
    a.start.row < b.start.row + b.rows &&
    b.start.row < a.start.row + a.rows &&
    a.start.col < b.start.col + b.columns &&
    b.start.col < a.start.col + a.columns
  );
}

// Excel parses a written string the way it parses typing: "00123" becomes 123,
// "9007199254740993" loses its last digits, "3-4" becomes a date, "=1+1" a
// formula (checked in a scratch Excel instance). A leading apostrophe stores the
// exact text and is not part of the value. CSV carries no types, so:
//  - by default, only conversions that destroy data are stopped — integers with
//    leading zeros and integers too long for a double stay text; dates,
//    percentages, booleans and formulas still parse, as a generated CSV expects;
//  - --text writes every cell as its exact text (codes, IDs, a text round trip);
//  - --no-formulas writes a leading "=" as text instead of a formula.
const LOSSY_INTEGER = /^[+-]?(?:0\d+|\d{16,})$/;
const asText = (raw) => ({ value: `'${raw}` });

function cellInput(raw, { text = false, formulas = true } = {}) {
  if (raw === "") return { value: "" };
  if (text) return asText(raw);
  if (raw.startsWith("=")) return formulas ? { formula: raw } : asText(raw);
  if (/^(true|false)$/i.test(raw)) return { value: raw.toLowerCase() === "true" };
  if (LOSSY_INTEGER.test(raw.trim())) return asText(raw);
  const number = Number(raw);
  return raw.trim() !== "" && !Number.isNaN(number) ? { value: number } : { value: raw };
}

const failure = (stderr) => ({ stdout: "", stderr: `${stderr}\n`, exitCode: 1 });
const resolvePath = (ctx, path) => (path.startsWith("/") ? path : `${ctx.cwd}/${path}`);

function workbookCommands(call, approveWrite, turnState, onAssert, onRead) {
  const csvExports = [];
  let exportTurn = turnState?.value;
  function currentExports() {
    if (exportTurn !== turnState?.value) {
      csvExports.length = 0;
      exportTurn = turnState?.value;
    }
    return csvExports;
  }
  async function writeCells({ cells, sheetId, start, startArg, force, source, ctx }) {
    const committed = [];
    try {
      const width = cells[0].length;
      if (
        width < 1 ||
        width > EXCEL_MAX_COLUMNS ||
        start.row + cells.length > EXCEL_MAX_ROWS ||
        start.col + width > EXCEL_MAX_COLUMNS
      ) {
        return failure(
          `Target from ${startArg} with ${cells.length} rows x ${width} columns is outside Excel's limits`,
        );
      }
      const target = `${columnLetters(start.col)}${start.row + 1}:${columnLetters(start.col + width - 1)}${start.row + cells.length}`;
      if (!force) {
        let pending = [target];
        let occupied = [];
        while (pending.length > 0 && occupied.length === 0) {
          const existing = await call(
            "excel_get_cell_ranges",
            { sheetId, ranges: pending, includeStyles: false, cellLimit: 5 },
            { signal: ctx.signal },
          );
          occupied = Object.keys({
            ...existing?.worksheet?.cells,
            ...existing?.worksheet?.formulas,
          });
          pending = existing?.remainingRanges ?? [];
        }
        if (occupied.length > 0) {
          return failure(
            `Target ${target} already has data (e.g. ${occupied.join(", ")}). ` +
              "If the requested edit targets these cells, rerun with --force; otherwise ask before overwriting.",
          );
        }
      }
      // Ask when the sandbox actually writes, including commands stored in scripts.
      const decision = await approveWrite("mcp__office__excel_set_cell_range", {
        sheetId,
        range: target,
        source,
        rows: cells.length,
        columns: width,
      });
      if (decision !== "approve" && decision !== "approve_turn") {
        return failure(
          decision === "reject"
            ? "The user rejected this workbook change. Do not retry it; ask what they would like instead."
            : `Approval for this workbook change did not complete (${decision}); nothing was written.`,
        );
      }
      const rowsPerChunk = Math.max(1, Math.floor(WRITE_CHUNK_CELLS / width));
      const formulaErrors = [];
      for (let offset = 0; offset < cells.length; offset += rowsPerChunk) {
        const chunk = cells.slice(offset, offset + rowsPerChunk);
        const top = start.row + offset + 1;
        const range = `${columnLetters(start.col)}${top}:${columnLetters(start.col + width - 1)}${top + chunk.length - 1}`;
        const result = await call(
          "excel_set_cell_range",
          { sheetId, range, cells: chunk, allow_overwrite: true },
          { signal: ctx.signal },
        );
        if (result?.commitStatus !== "committed") {
          throw new Error(
            `Write to ${range} did not confirm commit (status: ${result?.commitStatus ?? "unknown"}). ` +
              "Re-read this range before retrying.",
          );
        }
        committed.push(result?.writtenRange ?? range);
        formulaErrors.push(...(result?.formulaErrors ?? []));
      }
      const errorNote = formulaErrors.length
        ? `Formula errors in ${formulaErrors.length} cell(s): ` +
          formulaErrors
            .slice(0, 10)
            .map((item) => `${item.address}=${item.value}`)
            .join(", ") +
          (formulaErrors.length > 10 ? ", ..." : "") +
          ". Check whether they are expected (e.g. #N/A for no match) before reporting.\n"
        : "";
      return {
        stdout:
          `Committed ${cells.length} rows x ${width} columns to sheet ${sheetId} at ${target} ` +
          `in ${committed.length} committed chunk(s)\n${errorNote}`,
        stderr: "",
        exitCode: 0,
      };
    } catch (error) {
      const partial = committed.length
        ? `Partial write: ${committed.length} chunk(s) committed through ${committed.at(-1)}. `
        : "";
      return failure(partial + (error?.message ?? String(error)));
    }
  }

  const sheetToCsv = defineCommand("sheet-to-csv", async (args, ctx) => {
    const [sheetArg, second, third] = args;
    const sheetId = Number.parseInt(sheetArg, 10);
    if (!Number.isInteger(sheetId)) {
      return failure(
        "Usage: sheet-to-csv <sheetId> [range] [file]  (sheetId from mcp__office__excel_get_workbook_metadata or the [Auto-context] overview)",
      );
    }
    let range = second && /^[A-Z]+\d+(:[A-Z]+\d+)?$/i.test(second) ? second : undefined;
    const outFile = range ? third : second;
    try {
      if (!range) {
        const probe = await call(
          "excel_get_cell_ranges",
          { sheetId, ranges: ["A1"], includeStyles: false, cellLimit: 1 },
          { signal: ctx.signal },
        );
        range = probe?.worksheet?.dimension;
        if (!range) return failure("Sheet is empty (no used range)");
      }
      const pages = [];
      let next = range;
      let rows = 0;
      let columns = 0;
      let sheetName = "";
      while (next) {
        // includeHeaders keeps each page's first row; nothing here is treated as a header.
        const page = await call(
          "excel_get_range_as_csv",
          {
            sheetId,
            range: next,
            includeHeaders: true,
            maxRows: CSV_PAGE_ROWS,
          },
          { signal: ctx.signal },
        );
        // A page of one blank cell serializes to "" and is still a row, so keep
        // pages by their row count, not by whether their text is empty.
        if (page.rowCount > 0) pages.push(page.csv ?? "");
        rows += page.rowCount;
        columns = page.columnCount;
        sheetName = page.sheetName;
        next = page.hasMore ? page.nextRange : null;
      }
      const csv = pages.join("\n");
      onRead?.({ sheetId, range });
      if (!outFile) return { stdout: `${csv}\n`, stderr: "", exitCode: 0 };
      const path = resolvePath(ctx, outFile);
      const dir = path.slice(0, path.lastIndexOf("/"));
      if (dir) await ctx.fs.mkdir(dir, { recursive: true }).catch(() => {});
      await ctx.fs.writeFile(path, `${csv}\n`);
      currentExports().push({ sheetId, range: parseRange(range) });
      return {
        stdout: `Exported ${rows} rows x ${columns} columns from "${sheetName}"!${range} to ${outFile}\n`,
        stderr: "",
        exitCode: 0,
      };
    } catch (error) {
      return failure(error?.message ?? String(error));
    }
  });

  const csvToSheet = defineCommand("csv-to-sheet", async (args, ctx) => {
    const force = args.includes("--force");
    const exactText = args.includes("--text");
    const formulas = !args.includes("--no-formulas");
    const allowTypeLoss = args.includes("--allow-type-loss");
    const [file, sheetArg, startArg = "A1"] = args.filter((a) => !a.startsWith("--"));
    const sheetId = Number.parseInt(sheetArg, 10);
    const start = parseCell(startArg ?? "");
    if (!file || !Number.isInteger(sheetId) || !start) {
      return failure(
        "Usage: csv-to-sheet <file> <sheetId> [startCell] [--force] [--text] [--no-formulas] [--allow-type-loss]",
      );
    }
    try {
      const text = await ctx.fs.readFile(resolvePath(ctx, file));
      const parsed = Papa.parse(text.replace(/\r?\n$/, ""), { skipEmptyLines: false });
      // PapaParse reports a warning when a valid one-column file has no
      // delimiter to detect. That is a supported CSV shape, not a parse
      // failure. Preserve real quote/field errors.
      const parseErrors = parsed.errors.filter((error) => error.code !== "UndetectableDelimiter");
      if (parseErrors.length > 0) {
        return failure(`CSV parse failed: ${parseErrors[0].message}`);
      }
      const rows = parsed.data;
      if (rows.length === 0) return failure("CSV file is empty");
      const width = rows.reduce((maximum, row) => Math.max(maximum, row.length), 0);
      if (width < 1 || width > EXCEL_MAX_COLUMNS)
        return failure(`CSV width ${width} is outside Excel's limits`);
      if (start.row + rows.length > EXCEL_MAX_ROWS || start.col + width > EXCEL_MAX_COLUMNS) {
        return failure(
          `CSV target from ${startArg} with ${rows.length} rows x ${width} columns is outside Excel's limits`,
        );
      }
      const targetShape = { start, rows: rows.length, columns: width };
      if (
        force &&
        !allowTypeLoss &&
        currentExports().some(
          (item) => item.sheetId === sheetId && item.range && overlaps(item.range, targetShape),
        )
      ) {
        return failure(
          "CSV was exported from an overlapping range in this workbook. Writing it back can change text, number and formula types. Use sheet-to-json/json-to-sheet for a typed round trip; use --allow-type-loss only when those conversions are intentional.",
        );
      }
      const cells = rows.map((row) =>
        Array.from({ length: width }, (_, i) =>
          cellInput(row[i] ?? "", { text: exactText, formulas }),
        ),
      );
      return writeCells({ cells, sheetId, start, startArg, force, source: file, ctx });
    } catch (error) {
      return failure(error?.message ?? String(error));
    }
  });

  // Typed JSON is for editing data already in the workbook. CSV remains useful
  // for analysis or generated output, but its strings cannot encode Excel cell
  // types and therefore cannot safely round-trip a mixed column.
  async function readTypedRows(sheetId, requestedRange, ctx) {
    let range = requestedRange;
    if (!range) {
      const probe = await call(
        "excel_get_cell_ranges",
        { sheetId, ranges: ["A1"], includeStyles: false, cellLimit: 1 },
        { signal: ctx.signal },
      );
      range = probe?.worksheet?.dimension;
    }
    const shape = parseRange(range);
    if (!shape) throw new Error(`Invalid worksheet range: ${range ?? "(none)"}`);
    if (shape.rows * shape.columns > MAX_TYPED_CELLS) {
      throw new Error(`Typed read exceeds ${MAX_TYPED_CELLS} cells; select a smaller range.`);
    }
    const values = {};
    const formulas = {};
    let sheetName;
    let pending = [range];
    while (pending.length > 0) {
      const page = await call(
        "excel_get_cell_ranges",
        { sheetId, ranges: pending, includeStyles: false, cellLimit: 20_000 },
        { signal: ctx.signal },
      );
      if (page?.success === false || !page?.worksheet) {
        throw new Error(page?.error ?? "Workbook read returned no worksheet data");
      }
      Object.assign(values, page.worksheet.cells ?? {});
      Object.assign(formulas, page.worksheet.formulas ?? {});
      sheetName = page.worksheet.name;
      pending = page.remainingRanges ?? [];
    }
    const ambiguous = Object.keys(formulas).filter(
      (address) => values[address] === formulas[address],
    );
    for (let offset = 0; offset < ambiguous.length; offset += 100) {
      const addresses = ambiguous.slice(offset, offset + 100);
      const check = await call(
        "excel_get_formula_flags",
        { sheetName, addresses },
        { signal: ctx.signal },
      );
      for (const address of addresses) {
        if (typeof check?.flags?.[address] !== "boolean") {
          throw new Error(`Could not determine whether ${address} is a formula`);
        }
        if (!check.flags[address]) delete formulas[address];
      }
    }
    const rows = Array.from({ length: shape.rows }, (_, r) =>
      Array.from({ length: shape.columns }, (_, c) => {
        const address = `${columnLetters(shape.start.col + c)}${shape.start.row + r + 1}`;
        return Object.hasOwn(formulas, address)
          ? { formula: formulas[address] }
          : Object.hasOwn(values, address)
            ? { value: values[address] }
            : null;
      }),
    );
    onRead?.({ sheetId, range });
    return { range, shape, rows };
  }

  const sheetToJson = defineCommand("sheet-to-json", async (args, ctx) => {
    const [sheetArg, second, third] = args;
    const sheetId = Number.parseInt(sheetArg, 10);
    const rangeArg = second && parseRange(second) ? second : undefined;
    const outFile = rangeArg ? third : second;
    if (!Number.isInteger(sheetId) || !outFile || (third && !rangeArg)) {
      return failure("Usage: sheet-to-json <sheetId> [range] <file>");
    }
    try {
      const { range, shape, rows } = await readTypedRows(sheetId, rangeArg, ctx);
      const path = resolvePath(ctx, outFile);
      const dir = path.slice(0, path.lastIndexOf("/"));
      if (dir) await ctx.fs.mkdir(dir, { recursive: true });
      await ctx.fs.writeFile(path, JSON.stringify({ format: "excel-typed-cells-v1", range, rows }));
      return {
        stdout: `Exported ${shape.rows} rows x ${shape.columns} typed cells to ${outFile}\n`,
        stderr: "",
        exitCode: 0,
      };
    } catch (error) {
      return failure(error?.message ?? String(error));
    }
  });

  // Compare an independently computed expected matrix with the live workbook.
  // The command only reports a match; the caller remains responsible for how
  // the expected result was derived from the task and source data.
  const assertSheetJson = defineCommand("assert-sheet-json", async (args, ctx) => {
    const [file, sheetArg, range] = args;
    const sheetId = Number.parseInt(sheetArg, 10);
    const shape = parseRange(range);
    if (!file || !Number.isInteger(sheetId) || !shape || args.length !== 3) {
      return failure("Usage: assert-sheet-json <expectedFile> <sheetId> <range>");
    }
    try {
      const expected = JSON.parse(await ctx.fs.readFile(resolvePath(ctx, file)));
      if (
        expected?.format !== "excel-typed-cells-v1" ||
        !Array.isArray(expected.rows) ||
        expected.rows.length !== shape.rows ||
        expected.rows.some((row) => !Array.isArray(row) || row.length !== shape.columns)
      ) {
        return failure(
          `Expected matrix must contain ${shape.rows} rows x ${shape.columns} columns`,
        );
      }
      const actual = (await readTypedRows(sheetId, range, ctx)).rows;
      let mismatches = 0;
      const examples = [];
      for (let r = 0; r < shape.rows; r++) {
        for (let c = 0; c < shape.columns; c++) {
          if (JSON.stringify(expected.rows[r][c]) === JSON.stringify(actual[r][c])) continue;
          mismatches++;
          if (examples.length < 5) {
            const address = `${columnLetters(shape.start.col + c)}${shape.start.row + r + 1}`;
            examples.push(
              `${address}: expected ${JSON.stringify(expected.rows[r][c])}, got ${JSON.stringify(actual[r][c])}`,
            );
          }
        }
      }
      if (mismatches) return failure(`${mismatches} cell(s) differ. ${examples.join("; ")}`);
      onAssert?.({ sheetId, range });
      return {
        stdout: `Matched ${shape.rows * shape.columns} typed cells in ${range}\n`,
        stderr: "",
        exitCode: 0,
      };
    } catch (error) {
      return failure(error?.message ?? String(error));
    }
  });

  const jsonToSheet = defineCommand("json-to-sheet", async (args, ctx) => {
    const force = args.includes("--force");
    const [file, sheetArg, startArg = "A1"] = args.filter((arg) => !arg.startsWith("--"));
    const sheetId = Number.parseInt(sheetArg, 10);
    const start = parseCell(startArg);
    if (!file || !Number.isInteger(sheetId) || !start) {
      return failure("Usage: json-to-sheet <file> <sheetId> [startCell] [--force]");
    }
    try {
      const data = JSON.parse(await ctx.fs.readFile(resolvePath(ctx, file)));
      if (
        data?.format !== "excel-typed-cells-v1" ||
        !Array.isArray(data.rows) ||
        !data.rows.length
      ) {
        return failure("Expected a sheet-to-json file with non-empty rows");
      }
      const width = data.rows[0]?.length;
      if (
        !Number.isInteger(width) ||
        width < 1 ||
        width > EXCEL_MAX_COLUMNS ||
        data.rows.length * width > MAX_TYPED_CELLS ||
        data.rows.some((row) => !Array.isArray(row) || row.length !== width)
      ) {
        return failure("Typed cell rows must be rectangular and within the export size limit");
      }
      const cells = data.rows.map((row) =>
        row.map((cell) => {
          if (cell === null) return { value: "" };
          if (!cell || typeof cell !== "object" || Array.isArray(cell)) {
            throw new Error("Each typed cell must be null, {value}, or {formula}");
          }
          const keys = Object.keys(cell);
          if (keys.length !== 1)
            throw new Error("Each typed cell must have exactly one value or formula field");
          if (
            keys[0] === "formula" &&
            typeof cell.formula === "string" &&
            cell.formula.startsWith("=")
          ) {
            return { formula: cell.formula };
          }
          if (
            keys[0] !== "value" ||
            (!["string", "number", "boolean"].includes(typeof cell.value) && cell.value !== null)
          ) {
            throw new Error("Invalid typed cell value or formula");
          }
          if (typeof cell.value === "number" && !Number.isFinite(cell.value)) {
            throw new Error("Typed cell numbers must be finite");
          }
          return typeof cell.value === "string"
            ? cell.value === ""
              ? { value: "" }
              : asText(cell.value)
            : { value: cell.value ?? "" };
        }),
      );
      return writeCells({ cells, sheetId, start, startArg, force, source: file, ctx });
    } catch (error) {
      return failure(error?.message ?? String(error));
    }
  });

  return [sheetToCsv, csvToSheet, sheetToJson, jsonToSheet, assertSheetJson];
}

export const COMPUTE_TOOL_DESCRIPTION =
  "Run a command in a sandboxed bash shell with python3, awk, sed, sort, jq, sqlite3 and xan. " +
  "python3 has the standard library only (no numpy, pandas, openpyxl or sqlite3 module; use the sqlite3 command). " +
  "The shell has its own in-memory filesystem (cwd /home/user; files persist between calls but are lost when the " +
  "assistant restarts, so re-export if one is missing), no network, and no access to the workbook file, local " +
  "files or Windows programs. Use it for bulk data work too large to read into chat: " +
  "`sheet-to-csv <sheetId> [range] [file]` exports a range (default: used range) to a file; " +
  "`csv-to-sheet <file> <sheetId> [startCell] [--force] [--text] [--no-formulas] [--allow-type-loss]` writes a CSV back (refuses to overwrite data unless --force, " +
  "which is allowed when the user's requested edit targets those cells). CSV carries no types: numbers, dates, booleans and " +
  "leading = are parsed as Excel parses typing, except that integers with leading zeros or over 15 digits stay text. " +
  "Use --text to write every cell as its exact text (codes, IDs, text written back unchanged) and --no-formulas to keep a leading = as text. " +
  "After sheet-to-csv exports a range, csv-to-sheet refuses to overwrite an overlapping range unless --allow-type-loss explicitly accepts possible type changes; prefer typed JSON instead. " +
  "For a mixed-type round trip of at most 200000 cells, use `sheet-to-json <sheetId> [range] <file>` and `json-to-sheet <file> <sheetId> [startCell] [--force]`; " +
  "the JSON has `rows` of null, `{value: ...}` or `{formula: ...}` cells. String values remain text, numbers and booleans keep their types, and formulas remain formulas. " +
  "This preserves cell data types, not formatting; write only the columns the task needs to change. " +
  "For a deterministic transformation, compute an expected typed JSON matrix from the source and run `assert-sheet-json <expectedFile> <sheetId> <range>` after writing. It compares every target cell with live Excel and reports mismatches; do not derive the expected file by exporting the target. " +
  "Typical flow: " +
  "sheet-to-csv 1 A1:D5000 data.csv && python3 script.py && " +
  "csv-to-sheet out.csv 1 F1. sheet-to-csv and csv-to-sheet are shell commands: run them in the shell, " +
  "not from Python (subprocess cannot reach them). Write scripts with heredocs (cat > script.py <<'EOF' ... EOF). " +
  "Use sandbox paths such as data.csv; Windows paths, cd into the project, PowerShell, network access, openpyxl, " +
  "pandas and numpy are unavailable. Output is truncated to 30000 characters per stream.";

export function createComputeShell(
  call,
  { signal, approveWrite = async () => "approve", turnState, onAssert, onRead } = {},
) {
  const bash = new Bash({
    python: true,
    customCommands: workbookCommands(call, approveWrite, turnState, onAssert, onRead),
    executionLimits: {
      maxExecutionTimeMs: COMPUTE_TIMEOUT_MS,
      maxPythonTimeoutMs: 40_000,
      maxFileSystemBytes: 128 * 1024 * 1024,
      maxOutputSize: 4 * 1024 * 1024,
    },
  });
  return async ({ command }) => {
    const timeoutSignal = AbortSignal.timeout(COMPUTE_TIMEOUT_MS);
    const executionSignal = signal ? AbortSignal.any([signal, timeoutSignal]) : timeoutSignal;
    const result = await bash.exec(command, { signal: executionSignal });
    return {
      exitCode: result.exitCode,
      stdout: truncate(result.stdout, "stdout"),
      stderr: truncate(result.stderr, "stderr"),
    };
  };
}
