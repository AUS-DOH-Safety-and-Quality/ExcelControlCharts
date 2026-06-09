import { spawn } from "node:child_process";
import fsSync from "node:fs";
import fs from "node:fs/promises";
import path from "node:path";
import os from "node:os";
import AdmZip, { IZipEntry } from "adm-zip";
import { generateCertificates } from "office-addin-dev-certs";
import { registerAddIn } from "office-addin-dev-settings";
import { OfficeAddinManifest } from "office-addin-manifest";

declare const Bun: {
  serve(options: {
    hostname: string;
    port: number;
    tls: {
      cert: Buffer;
      key: Buffer;
    };
    fetch(request: Request): Response | Promise<Response>;
  }): {
    port: number;
    stop(closeActiveConnections?: boolean): void;
  };
};

type EmbeddedAssetMap = Record<string, string>;

type LaunchEmbeddedPortableOptions = {
  appName: string;
  manifestXml: string;
  excelTaskPaneTemplate: string;
  embeddedAssets: EmbeddedAssetMap;
  defaultPort: number;
};

type ManifestIdentity = {
  id: string;
  version: string;
};

function getCertificatePaths() {
  const certificateDirectory = path.join(os.homedir(), ".office-addin-dev-certs");
  return {
    caPath: path.join(certificateDirectory, "ca.crt"),
    certPath: path.join(certificateDirectory, "localhost.crt"),
    keyPath: path.join(certificateDirectory, "localhost.key"),
  };
}

function getPort(defaultPort: number): number {
  const portValue = process.env.PORTABLE_PORT;
  if (!portValue) {
    return defaultPort;
  }

  const parsedPort = Number.parseInt(portValue, 10);
  if (!Number.isInteger(parsedPort) || parsedPort <= 0 || parsedPort > 65535) {
    throw new Error(`Invalid PORTABLE_PORT value: ${portValue}`);
  }

  return parsedPort;
}

function getContentType(filePath: string): string | undefined {
  switch (path.extname(filePath).toLowerCase()) {
    case ".html":
      return "text/html; charset=utf-8";
    case ".js":
      return "application/javascript; charset=utf-8";
    case ".css":
      return "text/css; charset=utf-8";
    case ".xml":
      return "application/xml; charset=utf-8";
    case ".json":
      return "application/json; charset=utf-8";
    case ".png":
      return "image/png";
    case ".jpg":
    case ".jpeg":
      return "image/jpeg";
    case ".gif":
      return "image/gif";
    case ".ico":
      return "image/x-icon";
    case ".map":
      return "application/json; charset=utf-8";
    case ".xlsx":
      return "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    default:
      return undefined;
  }
}

function normalizeRequestPath(requestPath: string): string {
  const decodedPath = decodeURIComponent(requestPath.split("?")[0]);
  const relativePath = decodedPath === "/" ? "/taskpane.html" : decodedPath;
  const normalizedPath = path.posix.normalize(relativePath);

  if (!normalizedPath.startsWith("/")) {
    return `/${normalizedPath}`;
  }

  return normalizedPath;
}

async function ensureManifestOnDisk(appName: string, manifestXml: string): Promise<string> {
  const manifestDirectory = path.join(os.tmpdir(), `${appName}-portable`);
  const manifestPath = path.join(manifestDirectory, "manifest.xml");
  // Office desktop sideloading still expects a real manifest file path even though the web assets are embedded.
  await fs.mkdir(manifestDirectory, { recursive: true });
  await fs.writeFile(manifestPath, manifestXml, "utf-8");
  return manifestPath;
}

async function tryInstallCurrentUserCaCertificate(caCertificatePath: string): Promise<boolean> {
  // Attempt to install CA certificate to Windows trust store.
  // This is best-effort only - the app works fine even if it fails.
  const command = [
    "$ErrorActionPreference = 'Stop'",
    "$caCertificatePath = $args[0]",
    "$certificate = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2",
    "$certificate.Import($caCertificatePath)",
    "$store = New-Object System.Security.Cryptography.X509Certificates.X509Store('Root', 'CurrentUser')",
    "$store.Open('ReadWrite')",
    "$store.Add($certificate)",
    "$store.Close()",
  ].join("; ");

  return new Promise<boolean>((resolve) => {
    const child = spawn(
      "powershell",
      ["-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", command, caCertificatePath],
      { stdio: "pipe" }
    );

    child.on("exit", (code) => {
      resolve(code === 0);
    });
    child.on("error", () => {
      resolve(false);
    });
  });
}

async function ensureHttpsServerOptions() {
  const { caPath, certPath, keyPath } = getCertificatePaths();
  const certificateFilesExist = await Promise.all([caPath, certPath, keyPath].map(async (filePath) => {
    try {
      await fs.access(filePath);
      return true;
    } catch {
      return false;
    }
  }));

  if (certificateFilesExist.some((exists) => !exists)) {
    console.log("Generating self-signed HTTPS certificates...");
    await generateCertificates(caPath, certPath, keyPath, 365, ["127.0.0.1", "localhost"]);
    
    // Try to install to Windows trust store (best-effort, non-blocking)
    console.log("Attempting to install CA certificate to Windows trusted store...");
    const installed = await tryInstallCurrentUserCaCertificate(caPath);
    
    if (installed) {
      console.log("✓ CA certificate installed successfully");
    } else {
      console.log("⚠ CA certificate installation skipped (non-admin or permission issue)");
      console.log("  Your browser may show security warnings - this is normal.");
      console.log("  To avoid warnings, run the exe as Administrator once.");
    }
  }

  const [ca, cert, key] = await Promise.all([
    fs.readFile(caPath),
    fs.readFile(certPath),
    fs.readFile(keyPath),
  ]);

  return { ca, cert, key };
}

async function readManifestIdentity(manifestPath: string): Promise<ManifestIdentity> {
  const manifest = await OfficeAddinManifest.readManifestFile(manifestPath);
  if (!manifest.id || !manifest.version) {
    throw new Error("The generated manifest is missing an id or version.");
  }

  return {
    id: manifest.id,
    version: manifest.version,
  };
}

function makeUniqueTempPath(fileName: string): string {
  const parsedPath = path.parse(path.join(os.tmpdir(), fileName));
  let candidatePath = path.join(parsedPath.dir, `${parsedPath.name}${parsedPath.ext}`);
  let suffix = 2;

  while (fsSync.existsSync(candidatePath)) {
    candidatePath = path.join(parsedPath.dir, `${parsedPath.name}.${suffix}${parsedPath.ext}`);
    suffix += 1;
  }

  return candidatePath;
}

async function createExcelSideloadWorkbook(manifestPath: string, excelTaskPaneTemplate: string): Promise<string> {
  const { id, version } = await readManifestIdentity(manifestPath);
  const templateZip = new AdmZip(Buffer.from(excelTaskPaneTemplate, "base64"));
  const outputZip = new AdmZip();
  const webExtensionPath = "xl/webextensions/webextension.xml";
  const webExtensionEntry = templateZip.getEntry(webExtensionPath);

  if (!webExtensionEntry) {
    throw new Error("The embedded Excel sideload template is missing xl/webextensions/webextension.xml.");
  }

  const webExtensionXml = templateZip
    .readAsText(webExtensionEntry)
    .replace(/00000000-0000-0000-0000-000000000000/g, id)
    .replace(/1.0.0.0/g, version);

  templateZip.getEntries().forEach((entry: IZipEntry) => {
    let entryData = entry.getData();
    if (entry.entryName === webExtensionPath) {
      entryData = Buffer.from(webExtensionXml, "utf-8");
    }
    outputZip.addFile(entry.entryName, entryData, entry.comment, entry.attr);
  });

  const workbookPath = makeUniqueTempPath(`Excel add-in ${id}.xlsx`);
  await outputZip.writeZipPromise(workbookPath);
  return workbookPath;
}

async function launchWorkbook(workbookPath: string): Promise<void> {
  await new Promise<void>((resolve, reject) => {
    const child = spawn("cmd.exe", ["/c", "start", "", workbookPath], {
      detached: true,
      stdio: "ignore",
    });

    child.on("error", reject);
    child.on("spawn", () => {
      child.unref();
      resolve();
    });
  });
}

export async function launchEmbeddedPortable(options: LaunchEmbeddedPortableOptions): Promise<void> {
  const port = getPort(options.defaultPort);
  const httpsOptions = await ensureHttpsServerOptions();
  const manifestPath = await ensureManifestOnDisk(options.appName, options.manifestXml);
  const assetCache = new Map<string, Buffer>();

  // Decode once on startup so requests can be served from memory without touching disk.
  for (const [assetPath, base64Content] of Object.entries(options.embeddedAssets)) {
    assetCache.set(assetPath, Buffer.from(base64Content, "base64"));
  }

  const server = Bun.serve({
    hostname: "localhost",
    port,
    tls: {
      cert: httpsOptions.cert,
      key: httpsOptions.key,
    },
    async fetch(request: Request) {
      const requestUrl = new URL(request.url);

      if (requestUrl.pathname === "/health") {
        return new Response(JSON.stringify({ ok: true, port }), {
          headers: { "content-type": "application/json; charset=utf-8" },
        });
      }

      const assetPath = normalizeRequestPath(requestUrl.pathname);
      const asset = assetCache.get(assetPath);
      if (!asset) {
        return new Response("Not found", { status: 404 });
      }

      const headers = new Headers();
      const contentType = getContentType(assetPath);
      if (contentType) {
        headers.set("content-type", contentType);
      }
      headers.set("cache-control", "no-cache");

      const responseBody = new Blob([Uint8Array.from(asset)]);
      return new Response(responseBody, { headers });
    },
  });

  console.log(`${options.appName} portable host running at https://localhost:${server.port}`);
  console.log(`Serving ${assetCache.size} embedded files from memory`);

  await registerAddIn(manifestPath);
  const workbookPath = await createExcelSideloadWorkbook(manifestPath, options.excelTaskPaneTemplate);
  await launchWorkbook(workbookPath);

  const shutdown = () => {
    server.stop(true);
    process.exit(0);
  };

  process.on("SIGINT", shutdown);
  process.on("SIGTERM", shutdown);

  await new Promise(() => {
    // Keep the host process alive while Excel is using the add-in.
  });
}