import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { getHttpsServerOptions } from "office-addin-dev-certs";
import { runBuild, watchAndRebuild } from "./build";

const distRoot = path.join(process.cwd(), "dist");
// Set from the "config" block in package.json, which the runner exports as npm_package_config_*.
const port = Number(process.env.npm_package_config_dev_server_port) || 3100;

// Only key/cert: passing the `ca` from getHttpsServerOptions makes Bun.serve request a client certificate.
async function getCertificate(): Promise<{ key: Buffer; cert: Buffer }> {
  try {
    const { key, cert } = await getHttpsServerOptions();
    return { key, cert };
  } catch (error) {
    console.warn(`Falling back to existing dev certificate files: ${error}`);
    const certDirectory = path.join(os.homedir(), ".office-addin-dev-certs");
    return {
      key: await fs.readFile(path.join(certDirectory, "localhost.key")),
      cert: await fs.readFile(path.join(certDirectory, "localhost.crt")),
    };
  }
}

await runBuild("development");
watchAndRebuild("development");

Bun.serve({
  hostname: "localhost",
  port,
  tls: await getCertificate(),
  async fetch(request) {
    const pathname = decodeURIComponent(new URL(request.url).pathname);
    // The static spreadsheet page is the site root; Excel loads /taskpane.html directly.
    const filePath = path.join(distRoot, pathname === "/" ? "index.html" : pathname);
    const file = Bun.file(filePath);

    if (!filePath.startsWith(distRoot) || !(await file.exists())) {
      return new Response("Not found", { status: 404 });
    }

    return new Response(file, { headers: { "Access-Control-Allow-Origin": "*" } });
  },
});

console.log(`Dev server running at https://localhost:${port}/`);
