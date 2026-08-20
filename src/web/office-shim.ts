/**
 * Minimal stand-in for the parts of Office.js that the taskpane uses, backed by
 * the spreadsheet grid on the hosting page. It is loaded instead of the Office.js
 * CDN bundle in the web build of the taskpane, so `taskpane.ts` runs unmodified.
 *
 * Excel's batching model is reproduced faithfully enough for this consumer:
 * `load()` queues work, and the queued work runs on `context.sync()`.
 */
import type { CellValue, WorkbookHost } from "./host-api";

const noHostMessage =
  "This page needs the spreadsheet host. Open index.html rather than loading the taskpane directly.";

export function resolveHost(): WorkbookHost {
  // Same-origin frame, so the hosting page is reachable directly.
  const found = (window.parent as Window | undefined)?.__excelHost ?? window.__excelHost;
  if (!found) {
    throw new Error(noHostMessage);
  }
  return found;
}

const host = resolveHost;

const readyHooks: (() => void)[] = [];
let readyFired = false;

/**
 * Runs `hook` once the taskpane's own `Office.onReady` callback has built the DOM
 * that the page integration rearranges.
 *
 * The bundler is free to order the two entry scripts either way round, so a hook
 * registered after the callback already ran fires immediately rather than never.
 */
export function afterReady(hook: () => void): void {
  if (readyFired) {
    hook();
    return;
  }
  readyHooks.push(hook);
}

class RequestContext {
  private queue: (() => void)[] = [];

  readonly workbook = new WorkbookProxy(this);

  defer(work: () => void): void {
    this.queue.push(work);
  }

  async sync(): Promise<void> {
    const pending = this.queue;
    this.queue = [];
    for (const work of pending) {
      work();
    }
  }
}

class WorkbookProxy {
  readonly worksheets: WorksheetCollection;

  constructor(context: RequestContext) {
    this.worksheets = new WorksheetCollection(context);
  }
}

// `load()` ignores the Excel property list it is given: the shim's proxies are
// small enough to populate every known property on the next `sync()`.
class WorksheetCollection {
  items: WorksheetProxy[] = [];

  constructor(private readonly context: RequestContext) {}

  load(): this {
    this.context.defer(() => {
      this.items = host()
        .listWorksheets()
        .map((name) => new WorksheetProxy(this.context, name));
    });
    return this;
  }

  getItem(name: string): WorksheetProxy {
    return new WorksheetProxy(this.context, name);
  }

  getActiveWorksheet(): WorksheetProxy {
    const worksheet = new WorksheetProxy(this.context, "");
    this.context.defer(() => {
      worksheet.name = host().getActiveWorksheet();
    });
    return worksheet;
  }
}

class WorksheetProxy {
  readonly tables: TableCollection;
  readonly shapes: ShapeCollection;

  constructor(
    context: RequestContext,
    public name: string
  ) {
    this.tables = new TableCollection(context, this);
    this.shapes = new ShapeCollection(context, this);
  }

  load(): this {
    return this;
  }
}

class TableCollection {
  items: { name: string }[] = [];

  constructor(
    private readonly context: RequestContext,
    private readonly worksheet: WorksheetProxy
  ) {}

  load(): this {
    this.context.defer(() => {
      this.items = host()
        .listTables(this.worksheet.name)
        .map((name) => ({ name }));
    });
    return this;
  }

  getItem(name: string): TableProxy {
    return new TableProxy(this.context, this.worksheet, name);
  }
}

class TableProxy {
  readonly columns: ColumnCollection;

  constructor(
    context: RequestContext,
    worksheet: WorksheetProxy,
    readonly name: string
  ) {
    this.columns = new ColumnCollection(context, worksheet, this);
  }
}

class ColumnCollection {
  items: { name: string }[] = [];

  constructor(
    private readonly context: RequestContext,
    private readonly worksheet: WorksheetProxy,
    private readonly table: TableProxy
  ) {}

  load(): this {
    this.context.defer(() => {
      this.items = host()
        .listColumns(this.worksheet.name, this.table.name)
        .map((name) => ({ name }));
    });
    return this;
  }

  getItem(name: string): ColumnProxy {
    return new ColumnProxy(this.context, this.worksheet, this.table, name);
  }
}

class ColumnProxy {
  constructor(
    private readonly context: RequestContext,
    private readonly worksheet: WorksheetProxy,
    private readonly table: TableProxy,
    private readonly name: string
  ) {}

  getDataBodyRange(): RangeProxy {
    const range = new RangeProxy();
    range.fetch = () => host().getColumnValues(this.worksheet.name, this.table.name, this.name);
    range.context = this.context;
    return range;
  }
}

class RangeProxy {
  values: CellValue[][] = [];
  fetch: () => CellValue[][] = () => [];
  context?: RequestContext;

  load(): this {
    this.context?.defer(() => {
      this.values = this.fetch();
    });
    return this;
  }
}

class ShapeProxy {
  name = "";
  top = 0;
  left = 0;
}

class ShapeCollection {
  constructor(
    private readonly context: RequestContext,
    private readonly worksheet: WorksheetProxy
  ) {}

  addImage(base64: string): ShapeProxy {
    const shape = new ShapeProxy();
    this.context.defer(() => {
      host().addImage(this.worksheet.name, decodeBase64(base64));
    });
    return shape;
  }
}

// The taskpane base64-encodes the chart markup as UTF-8; decode symmetrically.
function decodeBase64(value: string): string {
  const bytes = Uint8Array.from(atob(value), (character) => character.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

async function run<T>(callback: (context: RequestContext) => Promise<T>): Promise<T> {
  const context = new RequestContext();
  const result = await callback(context);
  await context.sync();
  return result;
}

const readyInfo = { host: "Excel", platform: "OfficeOnline" };

function onReady(callback?: (info: typeof readyInfo) => void): Promise<typeof readyInfo> {
  const notify = () => {
    callback?.(readyInfo);
    readyFired = true;
    for (const hook of readyHooks.splice(0)) {
      hook();
    }
  };
  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", notify, { once: true });
  } else {
    notify();
  }
  return Promise.resolve(readyInfo);
}

const officeShim = {
  onReady,
  HostType: { Excel: "Excel" },
  PlatformType: { OfficeOnline: "OfficeOnline" },
  actions: { associate: () => undefined },
};

Object.assign(window, {
  Office: officeShim,
  Excel: { run },
});

export {};
