import { watch } from "node:fs";
import fs from "node:fs/promises";
import path from "node:path";

export type BuildMode = "development" | "production";

const repoRoot = process.cwd();
const distRoot = path.join(repoRoot, "dist");

const isIgnoredAsset = (name: string) => name === "dummy_data.xlsx" || name.startsWith("~$");

const installerScripts = ["install.ps1", "uninstall.ps1", "install.sh", "uninstall.sh"];

const urlDev = "https://localhost:3100/";
const urlDeployed = "https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/";

// Generated next to taskpane.html so its relative asset paths keep resolving.
const webTaskpanePath = path.join(repoRoot, "src/taskpane/taskpane-web.html");
const officeJsTag = /<script[^>]*appsforoffice\.microsoft\.com[^>]*>\s*<\/script>/;

/**
 * The static web build reuses taskpane.html verbatim apart from swapping Office.js
 * for the page integration, which installs the shim that talks to the spreadsheet
 * grid on the hosting page and moves the chart into the page's own column.
 */
async function writeWebTaskpane(): Promise<void> {
  const source = await fs.readFile(path.join(repoRoot, "src/taskpane/taskpane.html"), "utf-8");
  if (!officeJsTag.test(source)) {
    throw new Error("Could not find the Office.js script tag in taskpane.html");
  }
  const generated = source.replace(
    officeJsTag,
    '<script type="module" src="../web/embed.ts"></script>'
  );

  // Watch mode watches src/, so only write when the contents actually change.
  const existing = await fs.readFile(webTaskpanePath, "utf-8").catch(() => null);
  if (existing !== generated) {
    await fs.writeFile(webTaskpanePath, generated);
  }
}

const mimeTypes: Record<string, string> = {
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".jpg": "image/jpeg",
  ".jpeg": "image/jpeg",
  ".gif": "image/gif",
};

/**
 * A `</script` sequence would close the carrier element early. Inside a JS or CSS
 * string literal the escaped form is equivalent, and it cannot legally occur
 * anywhere else, so this is lossless.
 */
function escapePayload(text: string): string {
  return text.replaceAll("</script", "<\\/script");
}

async function toDataUri(fileName: string): Promise<string> {
  const file = path.join(distRoot, fileName);
  const mime = mimeTypes[path.extname(fileName).toLowerCase()] ?? "application/octet-stream";
  return `data:${mime};base64,${(await fs.readFile(file)).toString("base64")}`;
}

/** Swaps every local image reference for a data URI. */
async function inlineImages(html: string): Promise<string> {
  const references = [...html.matchAll(/(?:src|href)="\.\/([^"]+\.(?:png|svg|jpe?g|gif))"/g)];
  let result = html;
  for (const [, fileName] of references) {
    const dataUri = await toDataUri(fileName);
    // Replacement functions, here and below: a literal replacement string would
    // treat `$$` and `$&` in bundled code as substitution patterns and corrupt it.
    result = result.replaceAll(`./${fileName}`, () => dataUri);
  }
  return result;
}

function takeAsset(html: string, pattern: RegExp): { fileName: string; html: string } {
  const match = html.match(pattern);
  if (!match) {
    throw new Error(`Could not find a bundled asset matching ${pattern} in the built HTML`);
  }
  return { fileName: match[1], html: html.replace(match[0], "") };
}

/**
 * Rewrites the built web app into a single self-contained index.html.
 *
 * Bun emits the stylesheet and bundle as separate files tagged `crossorigin`, and
 * both those fetches and the taskpane iframe's own document are blocked under a
 * file:// URL, where every file gets its own opaque origin. Inlining everything
 * leaves the page with no subresource requests at all, so it runs equally well
 * from GitHub Pages, a shared drive, or a USB stick.
 *
 * The taskpane keeps its own document: the frame is left on about:blank, which
 * inherits this page's origin even when that origin is opaque, so the shim can
 * still reach the host through `window.parent` while the panel's stylesheet stays
 * isolated from the page's.
 */
async function inlineWebApp(): Promise<void> {
  const indexPath = path.join(distRoot, "index.html");
  const taskpaneHtml = await fs.readFile(path.join(distRoot, "taskpane-web.html"), "utf-8");

  const taskpaneStyle = takeAsset(taskpaneHtml, /<link[^>]*?href="\.\/([^"]+\.css)"[^>]*>/);
  const taskpaneScript = takeAsset(
    taskpaneStyle.html,
    /<script[^>]*?src="\.\/([^"]+\.js)"[^>]*><\/script>/
  );

  const body = taskpaneScript.html.match(/<body([^>]*)>([\s\S]*)<\/body>/);
  if (!body) {
    throw new Error("Could not find the taskpane body in the built HTML");
  }
  const taskpaneMarkup = await inlineImages(
    `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body${body[1]}>${body[2]}</body></html>`
  );

  let index = await fs.readFile(indexPath, "utf-8");
  const indexStyle = takeAsset(index, /<link[^>]*?href="\.\/([^"]+\.css)"[^>]*>/);
  const indexScript = takeAsset(
    indexStyle.html,
    /<script[^>]*?src="\.\/([^"]+\.js)"[^>]*><\/script>/
  );
  index = indexScript.html;

  const read = (fileName: string) => fs.readFile(path.join(distRoot, fileName), "utf-8");
  const payloads = [
    ["taskpane-markup", taskpaneMarkup],
    ["taskpane-style", await read(taskpaneStyle.fileName)],
    ["taskpane-script", await read(taskpaneScript.fileName)],
  ] as const;

  const inlined = [
    `<style>${await read(indexStyle.fileName)}</style>`,
    ...payloads.map(
      ([id, content]) => `<script type="text/plain" id="${id}">${escapePayload(content)}</script>`
    ),
    `<script type="module">${escapePayload(await read(indexScript.fileName))}</script>`,
  ].join("\n");

  index = await inlineImages(index.replace("</head>", () => `${inlined}\n</head>`));
  await fs.writeFile(indexPath, index);
}

// Read rather than import the manifest so watch mode picks up edits.
async function renderManifest(mode: BuildMode): Promise<string> {
  const content = await fs.readFile(path.join(repoRoot, "manifest.xml"), "utf-8");
  return mode === "development" ? content : content.replaceAll(urlDev, urlDeployed);
}

export async function runBuild(mode: BuildMode): Promise<void> {
  await fs.rm(distRoot, { recursive: true, force: true });
  await writeWebTaskpane();

  const result = await Bun.build({
    entrypoints: [
      "src/taskpane/taskpane.html",
      "src/commands/commands.html",
      // Static, Excel-free version of the add-in, served at the site root.
      "src/web/index.html",
      "src/taskpane/taskpane-web.html",
    ].map((entry) => path.join(repoRoot, entry)),
    outdir: distRoot,
    sourcemap: "linked",
    minify: mode === "production",
    naming: { entry: "[name].[ext]" },
    target: "browser",
  });

  if (!result.success) {
    throw new AggregateError(result.logs, "Build failed");
  }

  await fs.cp(path.join(repoRoot, "assets"), path.join(distRoot, "assets"), {
    recursive: true,
    filter: (src) => !isIgnoredAsset(path.basename(src)),
  });
  await fs.writeFile(path.join(distRoot, "manifest.xml"), await renderManifest(mode));

  // Served alongside the manifest so users can install with a single download.
  for (const script of installerScripts) {
    await fs.cp(path.join(repoRoot, "install", script), path.join(distRoot, script));
  }

  await inlineWebApp();
}

export function watchAndRebuild(mode: BuildMode): void {
  let queued = false;
  let pending: Promise<unknown> = Promise.resolve();

  const rebuild = () => {
    if (queued) {
      return;
    }
    queued = true;
    pending = pending
      .then(async () => {
        queued = false;
        await runBuild(mode);
        console.log(`[${new Date().toLocaleTimeString()}] Rebuilt.`);
      })
      .catch((error) => console.error("Rebuild failed:", error));
  };

  for (const target of ["src", "assets", "manifest.xml"]) {
    watch(path.join(repoRoot, target), { recursive: true }, rebuild);
  }
}

if (import.meta.main) {
  const args = process.argv.slice(2);
  const modeIndex = args.indexOf("--mode");
  const mode = modeIndex === -1 ? "development" : args[modeIndex + 1];

  if (mode !== "development" && mode !== "production") {
    throw new Error(`Unknown --mode "${mode}". Expected "development" or "production".`);
  }

  await runBuild(mode);
  console.log(`Build complete (${mode}) -> ${path.relative(repoRoot, distRoot)}`);

  if (args.includes("--watch")) {
    console.log("Watching for changes...");
    watchAndRebuild(mode);
  }
}
