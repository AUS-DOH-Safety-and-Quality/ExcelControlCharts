import { Grid } from "./grid";
import type { CellValue, WorkbookHost } from "./host-api";
import {
  columnValues,
  createSheet,
  getSheet,
  hasTable,
  loadRows,
  loadWorkbook,
  parseDelimited,
  sampleSheet,
  saveWorkbook,
  tableColumns,
  uniqueSheetName,
  type Workbook,
} from "./workbook";

const workbook: Workbook = loadWorkbook();

const element = <T extends HTMLElement>(id: string): T => {
  const found = document.getElementById(id);
  if (!found) {
    throw new Error(`Missing element #${id}`);
  }
  return found as T;
};

const gridContainer = element("grid-container");
const sheetTabs = element("sheet-tabs");
const statusLine = element("status-line");
const chartHost = element("chart-host");
const workbookColumn = element("workbook-column");
const taskpanePanel = element("taskpane-panel");
const togglePanel = element<HTMLButtonElement>("toggle-panel");
const csvInput = element<HTMLInputElement>("open-csv");
const taskpaneFrame = element<HTMLIFrameElement>("taskpane-frame");
const resizeWorkbookHandle = element("resize-workbook");
const resizePanelHandle = element("resize-panel");

const grid = new Grid(gridContainer, { onChange: persist });

function activeSheet() {
  return getSheet(workbook, workbook.activeSheet);
}

function debounce(callback: () => void, delay: number): () => void {
  let timer: ReturnType<typeof setTimeout> | undefined;
  return () => {
    clearTimeout(timer);
    timer = setTimeout(callback, delay);
  };
}

const scheduleSave = debounce(() => saveWorkbook(workbook), 300);

function persist(): void {
  scheduleSave();
  renderStatus();
  // Edits to the data should show up in the chart without any further action.
  taskpaneFrame.contentWindow?.__refreshChart?.();
}

function renderStatus(): void {
  const sheet = activeSheet();
  const columns = tableColumns(sheet);
  statusLine.textContent = hasTable(sheet)
    ? `${sheet.tableName}: ${columns.length} column${columns.length === 1 ? "" : "s"} — ${columns.join(", ")}`
    : "Add a header in row 1 and some data below it to make this sheet available to the chart panel.";
}

function renderTabs(): void {
  sheetTabs.replaceChildren();

  for (const sheet of workbook.sheets) {
    const tab = document.createElement("button");
    tab.type = "button";
    tab.className = sheet.name === workbook.activeSheet ? "tab tab--active" : "tab";
    tab.textContent = sheet.name;
    tab.title = "Double-click to rename";
    tab.onclick = () => selectSheet(sheet.name);
    tab.ondblclick = () => renameSheet(sheet.name);
    sheetTabs.appendChild(tab);
  }

  const add = document.createElement("button");
  add.type = "button";
  add.className = "tab tab--add";
  add.textContent = "+";
  add.title = "Add a sheet";
  add.onclick = addSheet;
  sheetTabs.appendChild(add);
}

function selectSheet(name: string): void {
  workbook.activeSheet = name;
  renderTabs();
  grid.render(activeSheet());
  persist();
}

function addSheet(): void {
  const sheet = createSheet(uniqueSheetName(workbook, `Sheet${workbook.sheets.length + 1}`));
  workbook.sheets.push(sheet);
  selectSheet(sheet.name);
}

function renameSheet(name: string): void {
  const next = window.prompt("Sheet name", name)?.trim();
  if (!next || next === name) {
    return;
  }
  if (workbook.sheets.some((sheet) => sheet.name === next)) {
    window.alert(`A sheet named "${next}" already exists.`);
    return;
  }
  getSheet(workbook, name).name = next;
  if (workbook.activeSheet === name) {
    workbook.activeSheet = next;
  }
  renderTabs();
  persist();
}

function deleteSheet(): void {
  if (workbook.sheets.length === 1) {
    window.alert("A workbook needs at least one sheet.");
    return;
  }
  const name = workbook.activeSheet;
  if (!window.confirm(`Delete sheet "${name}"?`)) {
    return;
  }
  workbook.sheets = workbook.sheets.filter((sheet) => sheet.name !== name);
  selectSheet(workbook.sheets[0].name);
}

async function openFiles(files: FileList | null): Promise<void> {
  for (const file of Array.from(files ?? [])) {
    const rows = parseDelimited(await file.text());
    if (rows.length === 0) {
      continue;
    }
    const name = uniqueSheetName(workbook, file.name.replace(/\.[^.]+$/, "").slice(0, 31));
    const sheet = createSheet(name);
    loadRows(sheet, rows);
    workbook.sheets.push(sheet);
    workbook.activeSheet = name;
  }
  renderTabs();
  grid.render(activeSheet());
  persist();
}

/** Backs the Office.js shim running inside the taskpane frame. */
const excelHost: WorkbookHost = {
  listWorksheets: () => workbook.sheets.filter(hasTable).map((sheet) => sheet.name),
  getActiveWorksheet: () => workbook.activeSheet,
  listTables: (worksheet) => {
    const sheet = getSheet(workbook, worksheet);
    return hasTable(sheet) ? [sheet.tableName] : [];
  },
  listColumns: (worksheet) => tableColumns(getSheet(workbook, worksheet)),
  getColumnValues: (worksheet, _table, column): CellValue[][] =>
    columnValues(getSheet(workbook, worksheet), column),
  getChartHost: () => chartHost,
  // The chart is already live on the page, so the Excel "insert a picture" action
  // maps to handing the user the same image as a file.
  addImage: () => downloadSvg(),
};

window.__excelHost = excelHost;

/** The taskpane keeps the inactive chart type's container hidden. */
function currentSvg(): SVGSVGElement | null {
  return chartHost.querySelector("div:not([hidden]) svg");
}

function svgSize(svg: SVGSVGElement): { width: number; height: number } {
  const viewBox = svg
    .getAttribute("viewBox")
    ?.split(/[\s,]+/)
    .map(Number);
  const width = Number(svg.getAttribute("width")) || viewBox?.[2] || 640;
  const height = Number(svg.getAttribute("height")) || viewBox?.[3] || 480;
  return { width, height };
}

function download(blob: Blob, filename: string): void {
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  link.click();
  URL.revokeObjectURL(url);
}

function serialiseSvg(svg: SVGSVGElement): string {
  const clone = svg.cloneNode(true) as SVGSVGElement;
  clone.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  const { width, height } = svgSize(svg);
  clone.setAttribute("width", String(width));
  clone.setAttribute("height", String(height));
  return new XMLSerializer().serializeToString(clone);
}

function downloadSvg(): void {
  const svg = currentSvg();
  if (svg) {
    download(new Blob([serialiseSvg(svg)], { type: "image/svg+xml" }), "control-chart.svg");
  }
}

async function downloadPng(): Promise<void> {
  const svg = currentSvg();
  if (!svg) {
    return;
  }
  const { width, height } = svgSize(svg);
  const scale = 2;
  const source = new Blob([serialiseSvg(svg)], { type: "image/svg+xml" });
  const url = URL.createObjectURL(source);

  try {
    const image = new Image();
    image.width = width;
    image.height = height;
    await new Promise<void>((resolve, reject) => {
      image.onload = () => resolve();
      image.onerror = () => reject(new Error("Could not rasterise the chart"));
      image.src = url;
    });

    const canvas = document.createElement("canvas");
    canvas.width = width * scale;
    canvas.height = height * scale;
    const context = canvas.getContext("2d");
    if (!context) {
      return;
    }
    context.fillStyle = "#ffffff";
    context.fillRect(0, 0, canvas.width, canvas.height);
    context.setTransform(scale, 0, 0, scale, 0, 0);
    context.drawImage(image, 0, 0, width, height);

    const png = await new Promise<Blob | null>((resolve) => canvas.toBlob(resolve, "image/png"));
    if (png) {
      download(png, "control-chart.png");
    }
  } finally {
    URL.revokeObjectURL(url);
  }
}

interface ResizeOptions {
  min: number;
  max: number;
  invert?: boolean;
  storageKey: string;
}

// Drags `target`'s flex-basis between min/max, persisting the chosen width.
function makeResizable(handle: HTMLElement, target: HTMLElement, options: ResizeOptions): void {
  const clamp = (value: number) => Math.min(options.max, Math.max(options.min, value));

  const stored = Number(localStorage.getItem(options.storageKey));
  // Tracks the requested width, not the rendered one: flex-shrink can render the
  // column smaller than what was asked for, and re-basing each drag on the
  // rendered value would make repeated short drags converge short of the real limit.
  let requestedWidth = clamp(stored || target.getBoundingClientRect().width);
  if (stored) {
    target.style.flexBasis = `${requestedWidth}px`;
  }

  let startX = 0;
  let startWidth = 0;

  const onPointerMove = (event: PointerEvent) => {
    const delta = (event.clientX - startX) * (options.invert ? -1 : 1);
    requestedWidth = clamp(startWidth + delta);
    target.style.flexBasis = `${requestedWidth}px`;
  };

  const onPointerUp = (event: PointerEvent) => {
    handle.releasePointerCapture(event.pointerId);
    document.removeEventListener("pointermove", onPointerMove);
    document.removeEventListener("pointerup", onPointerUp);
    document.body.classList.remove("resizing");
    target.style.transition = "";
    localStorage.setItem(options.storageKey, String(requestedWidth));
  };

  handle.addEventListener("pointerdown", (event) => {
    startX = event.clientX;
    startWidth = requestedWidth;
    handle.setPointerCapture(event.pointerId);
    document.body.classList.add("resizing");
    // Suppress the collapse-toggle's flex-basis transition so the drag tracks the
    // cursor immediately instead of easing toward each move's target.
    target.style.transition = "none";
    document.addEventListener("pointermove", onPointerMove);
    document.addEventListener("pointerup", onPointerUp);
  });
}

makeResizable(resizeWorkbookHandle, workbookColumn, {
  min: 260,
  max: 640,
  storageKey: "excel-control-charts:workbook-width",
});
makeResizable(resizePanelHandle, taskpanePanel, {
  min: 300,
  max: 640,
  invert: true,
  storageKey: "excel-control-charts:panel-width",
});

const panelStorageKey = "excel-control-charts:panel-collapsed";

function setPanelCollapsed(collapsed: boolean): void {
  taskpanePanel.classList.toggle("taskpane-panel--collapsed", collapsed);
  resizePanelHandle.hidden = collapsed;
  togglePanel.setAttribute("aria-expanded", String(!collapsed));
  togglePanel.title = collapsed ? "Show panel" : "Hide panel";
  togglePanel.setAttribute("aria-label", togglePanel.title);
  try {
    localStorage.setItem(panelStorageKey, collapsed ? "1" : "0");
  } catch {
    // Persistence is optional.
  }
}

togglePanel.onclick = () => {
  setPanelCollapsed(!taskpanePanel.classList.contains("taskpane-panel--collapsed"));
};

element("add-rows").onclick = () => grid.grow(20, 0);
element("add-column").onclick = () => grid.grow(0, 1);
element("delete-sheet").onclick = deleteSheet;
element("load-sample").onclick = () => {
  const sheet = sampleSheet();
  sheet.name = uniqueSheetName(workbook, sheet.name);
  workbook.sheets.push(sheet);
  selectSheet(sheet.name);
};
element("open-csv-button").onclick = () => csvInput.click();
csvInput.onchange = () => {
  void openFiles(csvInput.files).then(() => {
    csvInput.value = "";
  });
};
element("download-svg").onclick = downloadSvg;
element("download-png").onclick = () => void downloadPng();

for (const eventName of ["dragover", "drop"] as const) {
  gridContainer.addEventListener(eventName, (event) => {
    event.preventDefault();
    if (eventName === "drop") {
      void openFiles((event as DragEvent).dataTransfer?.files ?? null);
    }
  });
}

renderTabs();
grid.render(activeSheet());
renderStatus();
setPanelCollapsed(localStorage.getItem(panelStorageKey) === "1");

/** Reads one of the inert payloads the build inlines into the page. */
function payload(id: string): string {
  const carrier = document.getElementById(id);
  if (!carrier?.textContent) {
    throw new Error(`Missing inlined payload #${id}`);
  }
  return carrier.textContent;
}

// Writes the taskpane into the frame instead of pointing it at a URL — a second
// document would be cross-origin under file://; about:blank inherits this page's origin.
function mountTaskpane(): void {
  const doc = taskpaneFrame.contentDocument;
  if (!doc) {
    throw new Error("The taskpane frame has no document to write into");
  }

  doc.open();
  doc.write(payload("taskpane-markup"));
  doc.close();

  // Set as text rather than written as markup, so neither needs escaping.
  const style = doc.createElement("style");
  style.textContent = payload("taskpane-style");
  doc.head.appendChild(style);

  const script = doc.createElement("script");
  script.type = "module";
  script.textContent = payload("taskpane-script");
  doc.body.appendChild(script);
}

// Mounted last so the shim inside the frame always finds `window.__excelHost`.
mountTaskpane();
