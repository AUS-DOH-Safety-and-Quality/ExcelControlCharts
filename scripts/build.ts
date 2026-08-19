import { watch } from "node:fs";
import fs from "node:fs/promises";
import path from "node:path";

export type BuildMode = "development" | "production";

const repoRoot = process.cwd();
const distRoot = path.join(repoRoot, "dist");
const urlDev = "https://localhost:3100/";
export const deployedBaseUrl = "https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/";

const isIgnoredAsset = (name: string) => name === "dummy_data.xlsx" || name.startsWith("~$");

async function renderManifest(mode: BuildMode, baseUrl?: string): Promise<string> {
  const content = await fs.readFile(path.join(repoRoot, "manifest.xml"), "utf-8");
  return mode === "development"
    ? content
    : content.replaceAll(urlDev, baseUrl ?? process.env.ADDIN_BASE_URL ?? deployedBaseUrl);
}

export async function runBuild(mode: BuildMode, baseUrl?: string): Promise<void> {
  await fs.rm(distRoot, { recursive: true, force: true });

  const result = await Bun.build({
    entrypoints: ["src/taskpane/taskpane.html", "src/commands/commands.html"].map((entry) =>
      path.join(repoRoot, entry)
    ),
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
  await fs.writeFile(path.join(distRoot, "manifest.xml"), await renderManifest(mode, baseUrl));
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
