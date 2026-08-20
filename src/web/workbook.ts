import type { CellValue } from "./host-api";

export const MIN_ROWS = 40;
export const MIN_COLUMNS = 6;

export interface Sheet {
  name: string;
  /** Name the shim reports for this sheet's single table. */
  tableName: string;
  /** Column headers; a blank header hides the column from the table. */
  header: string[];
  /** Data rows, stored as typed to keep round-tripping lossless. */
  rows: string[][];
}

export interface Workbook {
  sheets: Sheet[];
  activeSheet: string;
}

const columnLabels = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

export function columnLabel(index: number): string {
  let label = "";
  let remaining = index;
  do {
    label = columnLabels[remaining % 26] + label;
    remaining = Math.floor(remaining / 26) - 1;
  } while (remaining >= 0);
  return label;
}

export function createSheet(name: string, tableName = "Table1"): Sheet {
  return {
    name,
    tableName,
    header: Array.from({ length: MIN_COLUMNS }, () => ""),
    rows: Array.from({ length: MIN_ROWS }, () => Array.from({ length: MIN_COLUMNS }, () => "")),
  };
}

/** Grows a sheet in place so `rows` x `columns` cells are addressable. */
export function ensureSize(sheet: Sheet, rows: number, columns: number): void {
  const width = Math.max(columns, sheet.header.length, MIN_COLUMNS);
  while (sheet.header.length < width) {
    sheet.header.push("");
  }
  while (sheet.rows.length < Math.max(rows, MIN_ROWS)) {
    sheet.rows.push([]);
  }
  for (const row of sheet.rows) {
    while (row.length < width) {
      row.push("");
    }
  }
}

export function getSheet(workbook: Workbook, name: string): Sheet {
  const sheet = workbook.sheets.find((candidate) => candidate.name === name);
  if (!sheet) {
    throw new Error(`No worksheet named "${name}"`);
  }
  return sheet;
}

/** Column indices that carry a header, i.e. the columns the table exposes. */
function tableColumnIndices(sheet: Sheet): number[] {
  return sheet.header
    .map((name, index) => ({ name: name.trim(), index }))
    .filter((entry) => entry.name !== "")
    .map((entry) => entry.index);
}

export function tableColumns(sheet: Sheet): string[] {
  return tableColumnIndices(sheet).map((index) => sheet.header[index].trim());
}

/** Last row (exclusive) holding data in any of the table's columns. */
export function tableRowCount(sheet: Sheet, indices = tableColumnIndices(sheet)): number {
  for (let row = sheet.rows.length - 1; row >= 0; row -= 1) {
    if (indices.some((index) => (sheet.rows[row][index] ?? "").trim() !== "")) {
      return row + 1;
    }
  }
  return 0;
}

export function hasTable(sheet: Sheet): boolean {
  const indices = tableColumnIndices(sheet);
  return indices.length > 0 && tableRowCount(sheet, indices) > 0;
}

// Cells come out of the grid as strings; Excel hands the taskpane numbers for
// numeric cells and strings otherwise (dates included), so mirror that coercion.
export function coerceCell(raw: string): CellValue {
  const trimmed = raw.trim();
  if (trimmed === "") {
    return null;
  }
  if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(trimmed)) {
    const value = Number(trimmed);
    if (Number.isFinite(value)) {
      return value;
    }
  }
  return trimmed;
}

export function columnValues(sheet: Sheet, columnName: string): CellValue[][] {
  const index = sheet.header.findIndex((name) => name.trim() === columnName);
  if (index === -1) {
    throw new Error(`No column named "${columnName}" in table "${sheet.tableName}"`);
  }
  const rowCount = tableRowCount(sheet);
  return sheet.rows.slice(0, rowCount).map((row) => [coerceCell(row[index] ?? "")]);
}

/** Splits delimited text, honouring RFC 4180 quoting. */
export function parseDelimited(text: string, delimiter?: string): string[][] {
  const separator = delimiter ?? sniffDelimiter(text);
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;

  for (let i = 0; i < text.length; i += 1) {
    const char = text[i];

    if (quoted) {
      if (char === '"') {
        if (text[i + 1] === '"') {
          field += '"';
          i += 1;
        } else {
          quoted = false;
        }
      } else {
        field += char;
      }
      continue;
    }

    if (char === '"' && field === "") {
      quoted = true;
    } else if (char === separator) {
      row.push(field);
      field = "";
    } else if (char === "\r") {
      // Swallow; the following \n ends the row.
    } else if (char === "\n") {
      row.push(field);
      rows.push(row);
      row = [];
      field = "";
    } else {
      field += char;
    }
  }

  if (field !== "" || row.length > 0) {
    row.push(field);
    rows.push(row);
  }

  return rows.filter((entry) => entry.some((cell) => cell !== ""));
}

function sniffDelimiter(text: string): string {
  const sample = text.split("\n").slice(0, 5).join("\n");
  const counts = [
    { delimiter: "\t", count: (sample.match(/\t/g) ?? []).length },
    { delimiter: ",", count: (sample.match(/,/g) ?? []).length },
    { delimiter: ";", count: (sample.match(/;/g) ?? []).length },
  ];
  return counts.sort((a, b) => b.count - a.count)[0].count > 0
    ? counts.sort((a, b) => b.count - a.count)[0].delimiter
    : ",";
}

/** Replaces a sheet's contents with a parsed grid, treating row 0 as headers. */
export function loadRows(sheet: Sheet, grid: string[][]): void {
  const width = Math.max(MIN_COLUMNS, ...grid.map((row) => row.length));
  sheet.header = Array.from({ length: width }, (_, index) => grid[0]?.[index] ?? "");
  sheet.rows = grid.slice(1).map((row) => Array.from({ length: width }, (_, i) => row[i] ?? ""));
  ensureSize(sheet, sheet.rows.length + 5, width);
}

/** A small, plausible SPC dataset so the page is usable before any import. */
export function sampleSheet(): Sheet {
  const sheet = createSheet("Sample data");
  const grid: string[][] = [["Date", "Infections", "Bed days"]];
  const numerators = [
    12, 9, 14, 11, 8, 13, 10, 16, 22, 19, 12, 9, 11, 7, 13, 10, 12, 8, 14, 9, 11, 15, 10, 12,
  ];
  const denominators = [
    980, 1010, 1035, 995, 940, 1120, 1080, 1005, 960, 1015, 1090, 1130, 1002, 985, 1044, 1008, 1075,
    999, 1030, 1065, 1012, 978, 1040, 1098,
  ];

  for (let month = 0; month < numerators.length; month += 1) {
    const date = new Date(Date.UTC(2023, month, 1));
    grid.push([
      date.toISOString().slice(0, 10),
      String(numerators[month]),
      String(denominators[month]),
    ]);
  }

  loadRows(sheet, grid);
  return sheet;
}

export function createWorkbook(): Workbook {
  const sheet = sampleSheet();
  return { sheets: [sheet], activeSheet: sheet.name };
}

const storageKey = "excel-control-charts:workbook";

export function saveWorkbook(workbook: Workbook): void {
  try {
    localStorage.setItem(storageKey, JSON.stringify(workbook));
  } catch {
    // Private browsing or a full quota; the page still works without persistence.
  }
}

export function loadWorkbook(): Workbook {
  try {
    const stored = localStorage.getItem(storageKey);
    if (!stored) {
      return createWorkbook();
    }
    const parsed = JSON.parse(stored) as Workbook;
    if (!Array.isArray(parsed?.sheets) || parsed.sheets.length === 0) {
      return createWorkbook();
    }
    for (const sheet of parsed.sheets) {
      sheet.tableName ||= "Table1";
      ensureSize(sheet, sheet.rows.length, sheet.header.length);
    }
    if (!parsed.sheets.some((sheet) => sheet.name === parsed.activeSheet)) {
      parsed.activeSheet = parsed.sheets[0].name;
    }
    return parsed;
  } catch {
    return createWorkbook();
  }
}

export function uniqueSheetName(workbook: Workbook, base: string): string {
  if (!workbook.sheets.some((sheet) => sheet.name === base)) {
    return base;
  }
  let suffix = 2;
  while (workbook.sheets.some((sheet) => sheet.name === `${base} (${suffix})`)) {
    suffix += 1;
  }
  return `${base} (${suffix})`;
}
