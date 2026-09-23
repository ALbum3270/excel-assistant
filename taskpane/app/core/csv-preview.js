// A CSV read shown as a table with Excel's own row numbers and column letters,
// the way pi-for-excel renders read_range results (src/ui/render-csv-table.ts).
// The text sent to the model is untouched; this is only what the pane shows.
import Papa from "papaparse";

export const CSV_PREVIEW_ROWS = 12;
export const CSV_PREVIEW_COLUMNS = 8;

export function columnLetter(number) {
  return number > 0
    ? columnLetter(Math.floor((number - 1) / 26)) + String.fromCharCode(65 + ((number - 1) % 26))
    : "";
}

export function csvPreview(csv, range) {
  const rows = Papa.parse(String(csv), { skipEmptyLines: false }).data;
  const start = /^\$?([A-Z]+)\$?(\d+)/i.exec(
    String(range ?? "")
      .split("!")
      .pop() ?? "",
  );
  const firstRow = start ? Number(start[2]) : 1;
  const firstColumn = start
    ? [...start[1].toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0)
    : 1;
  const shown = rows.slice(0, CSV_PREVIEW_ROWS);
  const widest = rows.reduce((maximum, row) => Math.max(maximum, row.length), 0);
  const width = Math.min(widest, CSV_PREVIEW_COLUMNS);
  return {
    columns: Array.from({ length: width }, (_, index) => columnLetter(firstColumn + index)),
    rows: shown.map((row, index) => ({
      number: firstRow + index,
      cells: Array.from({ length: width }, (_, column) => row[column] ?? ""),
    })),
    hiddenRows: rows.length - shown.length,
    hiddenColumns: widest - width,
  };
}
