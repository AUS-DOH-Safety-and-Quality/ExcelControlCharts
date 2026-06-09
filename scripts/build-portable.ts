import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";

// Using the current working directory as the repo root when the build script runs
const repoRoot = process.cwd();
console.log("Repo root")
console.log(repoRoot)
const releaseRoot = path.join(repoRoot, "release");
const buildRoot = path.join(releaseRoot, "_build");
const portableZip = path.join(releaseRoot, "ExcelControlCharts-portable.zip");
const launcherOutput = path.join(releaseRoot, "ExcelControlCharts.exe");
const manifestUrl = process.env.PORTABLE_BASE_URL || "https://localhost:3100/";

type EmbeddedAssetMap = Record<string, string>;

async function run(command: string[], extraEnv?: NodeJS.ProcessEnv): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const child = spawn(command[0], command.slice(1), {
      cwd: repoRoot,
      stdio: "inherit",
      shell: process.platform === "win32",
      env: {
        ...process.env,
        ...extraEnv,
      },
    });

    child.on("exit", (code) => {
      if (code === 0) {
        resolve();
        return;
      }

      reject(new Error(`Command failed with exit code ${code}: ${command.join(" ")}`));
    });

    child.on("error", reject);
  });
}

async function listFiles(rootPath: string): Promise<string[]> {
  const results: string[] = [];
  const entries = await fs.readdir(rootPath, { withFileTypes: true });

  for (const entry of entries) {
    const entryPath = path.join(rootPath, entry.name);
    if (entry.isDirectory()) {
      results.push(...(await listFiles(entryPath)));
    } else {
      results.push(entryPath);
    }
  }

  return results;
}

function toModuleLiteral(value: string): string {
  return JSON.stringify(value);
}

async function generateEmbeddedAssetsModule(distRoot: string): Promise<string> {
  const allFiles = await listFiles(distRoot);
  const embeddedAssets: EmbeddedAssetMap = {};
  let manifestXml = "";
  const excelTaskPaneTemplatePath = path.join(
    repoRoot,
    "node_modules",
    "office-addin-dev-settings",
    "templates",
    "ExcelWorkbookWithTaskPane.xlsx"
  );
  const excelTaskPaneTemplate = (await fs.readFile(excelTaskPaneTemplatePath)).toString("base64");

  for (const filePath of allFiles) {
    const relativePath = `/${path.relative(distRoot, filePath).split(path.sep).join("/")}`;
    const fileBuffer = await fs.readFile(filePath);

    // The manifest is written to disk at runtime for Office sideloading; everything else is embedded into the EXE.
    if (relativePath === "/manifest.xml") {
      manifestXml = fileBuffer.toString("utf-8");
      continue;
    }

    embeddedAssets[relativePath] = fileBuffer.toString("base64");
  }

  if (!manifestXml) {
    throw new Error("dist/manifest.xml was not generated.");
  }

  const embeddedModulePath = path.join(buildRoot, "embedded-assets.ts");
  const moduleSource = [
    `export const manifestXml = ${toModuleLiteral(manifestXml)};`,
    `export const excelTaskPaneTemplate = ${toModuleLiteral(excelTaskPaneTemplate)};`,
    `export const embeddedAssets = ${JSON.stringify(embeddedAssets, null, 2)} as const;`,
  ].join("\n\n");

  await fs.mkdir(buildRoot, { recursive: true });
  await fs.writeFile(embeddedModulePath, moduleSource, "utf-8");
  return embeddedModulePath;
}

async function generateLauncherEntry(): Promise<string> {
  const entryPath = path.join(buildRoot, "portable-entry.ts");
  const entrySource = [
    'import { embeddedAssets, excelTaskPaneTemplate, manifestXml } from "./embedded-assets";',
    'import { launchEmbeddedPortable } from "../../scripts/portable-runtime";',
    "",
    "launchEmbeddedPortable({",
    '  appName: "ExcelControlCharts",',
    '  manifestXml,',
    '  excelTaskPaneTemplate,',
    '  embeddedAssets,',
    '  defaultPort: 3100,',
    "});",
  ].join("\n");

  await fs.writeFile(entryPath, entrySource, "utf-8");
  return entryPath;
}

async function createPortableArchive(): Promise<void> {
  if (process.platform === "win32") {
    await run([
      "powershell",
      "-NoProfile",
      "-ExecutionPolicy",
      "Bypass",
      "-Command",
      `Compress-Archive -Path "${launcherOutput}" -DestinationPath "${portableZip}" -Force`,
    ]);
    return;
  }

  await run(["tar", "-a", "-c", "-f", portableZip, "-C", releaseRoot, path.basename(launcherOutput)]);
}

async function removeIfPossible(targetPath: string, recursive = false): Promise<void> {
  try {
    await fs.rm(targetPath, { recursive, force: true });
  } catch (error) {
    if ((error as NodeJS.ErrnoException).code !== "EACCES") {
      throw error;
    }

    console.warn(`Skipping cleanup for locked path: ${targetPath}`);
  }
}

async function main(): Promise<void> {
  console.log(`Building portable bundle with base URL ${manifestUrl}`);

  await fs.mkdir(releaseRoot, { recursive: true });
  await removeIfPossible(buildRoot, true);
  await removeIfPossible(path.join(releaseRoot, "portable"), true);
  await removeIfPossible(path.join(releaseRoot, "ExcelControlCharts-portable"), true);
  await removeIfPossible(portableZip);
  await removeIfPossible(launcherOutput);
  await run(["bun", "x", "webpack", "--mode", "production"], {
    ADDIN_BASE_URL: manifestUrl,
  });

  const distRoot = path.join(repoRoot, "dist");
  // Generate a temporary TypeScript module so Bun can compile the built web assets directly into the launcher.
  await generateEmbeddedAssetsModule(distRoot);
  const launcherEntry = await generateLauncherEntry();
  await run(["bun", "build", launcherEntry, "--compile", "--outfile", launcherOutput]);
  await createPortableArchive();
  await fs.rm(buildRoot, { recursive: true, force: true });

  console.log(`Single-file launcher ready at ${launcherOutput}`);
  console.log(`Portable zip ready at ${portableZip}`);
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : error);
  process.exit(1);
});