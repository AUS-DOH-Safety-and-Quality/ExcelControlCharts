import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { registerAddIn } from "office-addin-dev-settings";

const deployedBaseUrl = "https://aus-doh-safety-and-quality.github.io/ExcelControlCharts/";
const manifestUrl = new URL("manifest.xml", process.env.ADDIN_BASE_URL ?? deployedBaseUrl).href;

const response = await fetch(manifestUrl);
const manifestXml = await response.text();

if (!response.ok || !manifestXml.includes("<OfficeApp")) {
  throw new Error(`No add-in manifest at ${manifestUrl} (HTTP ${response.status}).`);
}

// Registration refers to the manifest by path, so it has to outlive this process.
const workDir = path.join(os.tmpdir(), "ExcelControlCharts");
await fs.mkdir(workDir, { recursive: true });
const manifestPath = path.join(workDir, "manifest.xml");
await fs.writeFile(manifestPath, manifestXml);

await registerAddIn(manifestPath);

console.log(`Registered the add-in from ${manifestUrl}`);
console.log("Open Excel and choose it from the Home tab.");
