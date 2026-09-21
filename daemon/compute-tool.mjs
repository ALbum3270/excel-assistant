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
  const col = [...match[1].toUpperCase()].reduce((acc, c) => acc * 26 + c.charCodeAt(0) - 64, 0) - 1;
  return { col, row: Number(match[2]) - 1 };
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

function workbookCommands(call, approveWrite) {
  const sheetToCsv = defineCommand("sheet-to-csv", async (args, ctx) => {
    const [sheetArg, second, third] = args;
    const sheetId = Number.parseInt(sheetArg, 10);
    if (!Number.isInteger(sheetId)) {
      return failure("Usage: sheet-to-csv <sheetId> [range] [file]  (sheetId from mcp__office__excel_get_workbook_metadata or the [Auto-context] overview)");
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
        const page = await call("excel_get_range_as_csv", {
          sheetId,
          range: next,
          includeHeaders: true,
          maxRows: CSV_PAGE_ROWS,
        }, { signal: ctx.signal });
        if (page.csv) pages.push(page.csv);
        rows += page.rowCount;
        columns = page.columnCount;
        sheetName = page.sheetName;
        next = page.hasMore ? page.nextRange : null;
      }
      const csv = pages.join("\n");
      if (!outFile) return { stdout: `${csv}\n`, stderr: "", exitCode: 0 };
      const path = resolvePath(ctx, outFile);
      const dir = path.slice(0, path.lastIndexOf("/"));
      if (dir) await ctx.fs.mkdir(dir, { recursive: true }).catch(() => {});
      await ctx.fs.writeFile(path, `${csv}\n`);
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
    const [file, sheetArg, startArg = "A1"] = args.filter((a) => !a.startsWith("--"));
    const sheetId = Number.parseInt(sheetArg, 10);
    const start = parseCell(startArg ?? "");
    if (!file || !Number.isInteger(sheetId) || !start) {
      return failure("Usage: csv-to-sheet <file> <sheetId> [startCell] [--force] [--text] [--no-formulas]");
    }
    const committed = [];
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
      const width = Math.max(...rows.map((r) => r.length));
      if (width < 1 || width > 16_384) return failure(`CSV width ${width} is outside Excel's limits`);
      const cells = rows.map((row) =>
        Array.from({ length: width }, (_, i) => cellInput(row[i] ?? "", { text: exactText, formulas })),
      );
      const target = `${columnLetters(start.col)}${start.row + 1}:${columnLetters(start.col + width - 1)}${start.row + rows.length}`;
      if (!force) {
        // Check the whole target before the first chunk so a conflict never leaves a partial write.
        // Each read covers at most 20000 cells; remainingRanges continues past that.
        let pending = [target];
        let occupied = [];
        while (pending.length > 0 && occupied.length === 0) {
          const existing = await call("excel_get_cell_ranges", {
            sheetId,
            ranges: pending,
            includeStyles: false,
            cellLimit: 5,
          }, { signal: ctx.signal });
          occupied = Object.keys({ ...existing?.worksheet?.cells, ...existing?.worksheet?.formulas });
          pending = existing?.remainingRanges ?? [];
        }
        if (occupied.length > 0) {
          return failure(
            `Target ${target} already has data (e.g. ${occupied.join(", ")}). ` +
              "If the requested edit targets these cells, rerun with --force; otherwise ask before overwriting.",
          );
        }
      }
      // Approval is asked here, where the sandbox actually writes, not from the
      // command text: a saved script (`bash later.sh`) or a command built in a
      // variable wrote without asking when the check was a regex on the text.
      const decision = await approveWrite("mcp__office__excel_set_cell_range", {
        sheetId,
        range: target,
        source: file,
        rows: rows.length,
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
      // Formula errors (often an expected #N/A) don't stop the write: stopping
      // would leave a half-written table. They are listed once everything is in.
      const formulaErrors = [];
      for (let offset = 0; offset < cells.length; offset += rowsPerChunk) {
        const chunk = cells.slice(offset, offset + rowsPerChunk);
        const top = start.row + offset + 1;
        const range = `${columnLetters(start.col)}${top}:${columnLetters(start.col + width - 1)}${top + chunk.length - 1}`;
        const result = await call("excel_set_cell_range", {
          sheetId,
          range,
          cells: chunk,
          allow_overwrite: true,
        }, { signal: ctx.signal });
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
          formulaErrors.slice(0, 10).map((item) => `${item.address}=${item.value}`).join(", ") +
          (formulaErrors.length > 10 ? ", ..." : "") +
          ". Check whether they are expected (e.g. #N/A for no match) before reporting.\n"
        : "";
      return {
        stdout:
          `Committed ${rows.length} rows x ${width} columns to sheet ${sheetId} at ${target} ` +
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
  });

  return [sheetToCsv, csvToSheet];
}

export const COMPUTE_TOOL_DESCRIPTION =
  "Run a command in a sandboxed bash shell with python3, awk, sed, sort, jq, sqlite3 and xan. " +
  "python3 has the standard library only (no numpy, pandas, openpyxl or sqlite3 module; use the sqlite3 command). " +
  "The shell has its own in-memory filesystem (cwd /home/user; files persist between calls but are lost when the " +
  "assistant restarts, so re-export if one is missing), no network, and no access to the workbook file, local " +
  "files or Windows programs. Use it for bulk data work too large to read into chat: " +
  "`sheet-to-csv <sheetId> [range] [file]` exports a range (default: used range) to a file; " +
  "`csv-to-sheet <file> <sheetId> [startCell] [--force] [--text] [--no-formulas]` writes a CSV back (refuses to overwrite data unless --force, " +
  "which is allowed when the user's requested edit targets those cells). CSV carries no types: numbers, dates, booleans and " +
  "leading = are parsed as Excel parses typing, except that integers with leading zeros or over 15 digits stay text. " +
  "Use --text to write every cell as its exact text (codes, IDs, text written back unchanged) and --no-formulas to keep a leading = as text. " +
  "Typical flow: " +
  "sheet-to-csv 1 A1:D5000 data.csv && python3 script.py && " +
  "csv-to-sheet out.csv 1 F1. sheet-to-csv and csv-to-sheet are shell commands: run them in the shell, " +
  "not from Python (subprocess cannot reach them). Write scripts with heredocs (cat > script.py <<'EOF' ... EOF). " +
  "Use sandbox paths such as data.csv; Windows paths, cd into the project, PowerShell, network access, openpyxl, " +
  "pandas and numpy are unavailable. Output is truncated to 30000 characters per stream.";

export function createComputeShell(call, { signal, approveWrite = async () => "approve" } = {}) {
  const bash = new Bash({
    python: true,
    customCommands: workbookCommands(call, approveWrite),
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
