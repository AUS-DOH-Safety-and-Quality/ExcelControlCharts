import { columnLabel, ensureSize, MIN_COLUMNS, parseDelimited, type Sheet } from "./workbook";

interface GridOptions {
  onChange: () => void;
}

/**
 * A lightweight editable grid. Row 0 of the rendered table holds the column
 * headers, which become the table's column names for the taskpane.
 */
export class Grid {
  private readonly table = document.createElement("table");
  private sheet: Sheet | null = null;

  constructor(
    private readonly container: HTMLElement,
    private readonly options: GridOptions
  ) {
    this.table.className = "grid";
    this.container.appendChild(this.table);
    this.table.addEventListener("input", (event) => this.handleInput(event));
    this.table.addEventListener("keydown", (event) => this.handleKeydown(event));
    this.table.addEventListener("paste", (event) => this.handlePaste(event as ClipboardEvent));
  }

  render(sheet: Sheet): void {
    this.sheet = sheet;
    ensureSize(sheet, sheet.rows.length, sheet.header.length);

    const head = document.createElement("thead");
    const labelRow = document.createElement("tr");
    labelRow.appendChild(cornerCell());
    sheet.header.forEach((_, column) => {
      const cell = document.createElement("th");
      cell.className = "grid__label";
      cell.textContent = columnLabel(column);
      labelRow.appendChild(cell);
    });
    head.appendChild(labelRow);

    const headerRow = document.createElement("tr");
    const headerCorner = document.createElement("th");
    headerCorner.className = "grid__gutter";
    headerCorner.textContent = "1";
    headerRow.appendChild(headerCorner);
    sheet.header.forEach((value, column) => {
      const cell = document.createElement("th");
      cell.className = "grid__header";
      cell.appendChild(this.makeInput(value, -1, column, "Column name"));
      headerRow.appendChild(cell);
    });
    head.appendChild(headerRow);

    const body = document.createElement("tbody");
    sheet.rows.forEach((row, rowIndex) => {
      const tr = document.createElement("tr");
      const gutter = document.createElement("th");
      gutter.className = "grid__gutter";
      gutter.textContent = String(rowIndex + 2);
      tr.appendChild(gutter);
      row.forEach((value, column) => {
        const cell = document.createElement("td");
        cell.appendChild(this.makeInput(value, rowIndex, column));
        tr.appendChild(cell);
      });
      body.appendChild(tr);
    });

    this.table.replaceChildren(head, body);
  }

  focusCell(row: number, column: number): void {
    this.inputAt(row, column)?.focus();
  }

  private makeInput(
    value: string,
    row: number,
    column: number,
    placeholder?: string
  ): HTMLInputElement {
    const input = document.createElement("input");
    input.className = row === -1 ? "grid__cell grid__cell--header" : "grid__cell";
    input.value = value;
    input.dataset.row = String(row);
    input.dataset.column = String(column);
    input.autocomplete = "off";
    input.spellcheck = false;
    if (placeholder) {
      input.placeholder = placeholder;
    }
    return input;
  }

  private inputAt(row: number, column: number): HTMLInputElement | null {
    return this.table.querySelector(`input[data-row="${row}"][data-column="${column}"]`);
  }

  private coordinates(target: EventTarget | null): { row: number; column: number } | null {
    if (!(target instanceof HTMLInputElement) || target.dataset.row === undefined) {
      return null;
    }
    return { row: Number(target.dataset.row), column: Number(target.dataset.column) };
  }

  private handleInput(event: Event): void {
    const position = this.coordinates(event.target);
    if (!position || !this.sheet) {
      return;
    }
    const value = (event.target as HTMLInputElement).value;
    if (position.row === -1) {
      this.sheet.header[position.column] = value;
    } else {
      this.sheet.rows[position.row][position.column] = value;
    }
    this.options.onChange();
  }

  private handleKeydown(event: KeyboardEvent): void {
    const position = this.coordinates(event.target);
    if (!position) {
      return;
    }

    const move = (rows: number) => {
      event.preventDefault();
      this.focusCell(position.row + rows, position.column);
    };

    if (event.key === "Enter" || event.key === "ArrowDown") {
      move(1);
    } else if (event.key === "ArrowUp") {
      move(-1);
    }
  }

  private handlePaste(event: ClipboardEvent): void {
    const position = this.coordinates(event.target);
    const text = event.clipboardData?.getData("text/plain");
    if (!position || !this.sheet || !text || !/[\t\n]/.test(text)) {
      return;
    }

    event.preventDefault();
    const grid = parseDelimited(text);
    const width = Math.max(...grid.map((row) => row.length));
    ensureSize(this.sheet, position.row + grid.length, position.column + width);

    grid.forEach((row, rowOffset) => {
      row.forEach((value, columnOffset) => {
        const targetRow = position.row + rowOffset;
        const targetColumn = position.column + columnOffset;
        if (targetRow === -1) {
          this.sheet!.header[targetColumn] = value;
        } else {
          this.sheet!.rows[targetRow][targetColumn] = value;
        }
      });
    });

    this.render(this.sheet);
    this.focusCell(position.row, position.column);
    this.options.onChange();
  }

  /** Appends blank rows or columns so users can keep typing past the current edge. */
  grow(rows: number, columns: number): void {
    if (!this.sheet) {
      return;
    }
    ensureSize(
      this.sheet,
      this.sheet.rows.length + rows,
      Math.max(MIN_COLUMNS, this.sheet.header.length + columns)
    );
    this.render(this.sheet);
    this.options.onChange();
  }
}

function cornerCell(): HTMLTableCellElement {
  const cell = document.createElement("th");
  cell.className = "grid__gutter grid__corner";
  return cell;
}
