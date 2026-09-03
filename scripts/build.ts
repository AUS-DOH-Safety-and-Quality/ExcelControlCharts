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

// Swaps Office.js for the page-integration shim; taskpane.html is otherwise untouched.
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

// Escapes "</script" so it can't close the carrier element early; lossless since
// that sequence can't otherwise occur inside a JS/CSS string literal.
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
  const fileNames = new Set(
    [...html.matchAll(/(?:src|href)="\.\/([^"]+\.(?:png|svg|jpe?g|gif))"/g)].map(([, name]) => name)
  );
  const entries = await Promise.all(
    [...fileNames].map(async (fileName) => [fileName, await toDataUri(fileName)] as const)
  );
  let result = html;
  for (const [fileName, dataUri] of entries) {
    // Function form avoids "$$"/"$&" in bundled code being read as replacement patterns.
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

function extractAssets(html: string): { html: string; styleFile: string; scriptFile: string } {
  const style = takeAsset(html, /<link[^>]*?href="\.\/([^"]+\.css)"[^>]*>/);
  const script = takeAsset(style.html, /<script[^>]*?src="\.\/([^"]+\.js)"[^>]*><\/script>/);
  return { html: script.html, styleFile: style.fileName, scriptFile: script.fileName };
}

// Inlines the built JS/CSS/images into a single index.html: under file://, every
// subresource fetch is cross-origin and gets blocked, so this removes them all.
async function inlineWebApp(): Promise<void> {
  const indexPath = path.join(distRoot, "index.html");
  const taskpaneHtml = await fs.readFile(path.join(distRoot, "taskpane-web.html"), "utf-8");
  const taskpane = extractAssets(taskpaneHtml);

  const body = taskpane.html.match(/<body([^>]*)>([\s\S]*)<\/body>/);
  if (!body) {
    throw new Error("Could not find the taskpane body in the built HTML");
  }
  const taskpaneMarkup = await inlineImages(
    `<!DOCTYPE html><html><head><meta charset="utf-8"></head><body${body[1]}>${body[2]}</body></html>`
  );

  const index = extractAssets(await fs.readFile(indexPath, "utf-8"));

  const read = (fileName: string) => fs.readFile(path.join(distRoot, fileName), "utf-8");
  const [indexStyle, taskpaneStyle, taskpaneScript, indexScript] = await Promise.all([
    read(index.styleFile),
    read(taskpane.styleFile),
    read(taskpane.scriptFile),
    read(index.scriptFile),
  ]);

  const payloads = [
    ["taskpane-markup", taskpaneMarkup],
    ["taskpane-style", taskpaneStyle],
    ["taskpane-script", taskpaneScript],
  ] as const;

  const inlined = [
    `<style>${indexStyle}</style>`,
    ...payloads.map(
      ([id, content]) => `<script type="text/plain" id="${id}">${escapePayload(content)}</script>`
    ),
    `<script type="module">${escapePayload(indexScript)}</script>`,
  ].join("\n");

  const result = await inlineImages(index.html.replace("</head>", () => `${inlined}\n</head>`));
  await fs.writeFile(indexPath, result);
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
  await fs.cp(path.join(repoRoot, "src", "web", "manifest.webmanifest"), path.join(distRoot, "manifest.webmanifest"));
  await fs.cp(path.join(repoRoot, "src", "web", "service-worker.js"), path.join(distRoot, "service-worker.js"));

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
