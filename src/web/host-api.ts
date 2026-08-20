// Contract between the static spreadsheet page and the Office.js shim running in
// the embedded taskpane. Both sides are same-origin, so the shim reaches the host
// directly through `window.parent.__excelHost` rather than posting messages.

export type CellValue = string | number | null;

export interface WorkbookHost {
  listWorksheets(): string[];
  getActiveWorksheet(): string;
  listTables(worksheet: string): string[];
  listColumns(worksheet: string, table: string): string[];
  /** Column values shaped like an Excel data-body range: one row per entry. */
  getColumnValues(worksheet: string, table: string, column: string): CellValue[][];
  /** Stands in for `worksheet.shapes.addImage`. `svg` is the raw markup, already decoded. */
  addImage(worksheet: string, svg: string): void;
  /** Element in the hosting page that the chart is rendered into. */
  getChartHost(): HTMLElement;
}

declare global {
  interface Window {
    __excelHost?: WorkbookHost;
    /** Installed by the embedded taskpane so the page can redraw the chart. */
    __refreshChart?: () => void;
  }
}
