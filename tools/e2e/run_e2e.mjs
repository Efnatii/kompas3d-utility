import fs from "node:fs";
import path from "node:path";
import http from "node:http";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";
import { buildRuntimeConfig } from "../runtime/build_runtime_config.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..", "..");
const docsRoot = path.join(repoRoot, "docs");
const fixturePath = path.join(repoRoot, "xlsx-to-kompas-tbl", "fixtures", "table_M2.xlsx");
const outputRoot = path.join(repoRoot, "out", "e2e");

const utilityCandidates = [
  path.join("C:\\_GIT_\\web-bridge-utility", "artifacts", "publish", "utility", "win-x64", "WebBridge.Utility.exe"),
  path.join("C:\\_GIT_\\web-bridge-utility", "src", "WebBridge.Utility", "bin", "Release", "net8.0", "win-x64", "WebBridge.Utility.exe"),
  path.join("C:\\_GIT_\\web-bridge-utility", "artifacts", "release", "web-bridge-utility-1.0.0-win-x64", "WebBridge.Utility.exe"),
];

const kompasSampleCandidates = [
  path.join("C:\\Program Files\\ASCON\\KOMPAS-3D v24\\Libs\\Cable3D", "Plug.frw"),
  path.join("C:\\Program Files\\ASCON\\KOMPAS-3D v24\\Libs\\Cable3D", "Point.frw"),
  path.join("C:\\Program Files\\ASCON\\KOMPAS-3D v24\\Libs\\Cable3D", "Socket.frw"),
];

function nowStamp() {
  const date = new Date();
  const yyyy = String(date.getFullYear());
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const dd = String(date.getDate()).padStart(2, "0");
  const hh = String(date.getHours()).padStart(2, "0");
  const mi = String(date.getMinutes()).padStart(2, "0");
  const ss = String(date.getSeconds()).padStart(2, "0");
  return `${yyyy}${mm}${dd}-${hh}${mi}${ss}`;
}

function parseArgs(argv) {
  const parsed = { browser: "msedge" };
  for (let index = 2; index < argv.length; index += 1) {
    const token = argv[index];
    const value = argv[index + 1];
    switch (token) {
      case "--browser":
        parsed.browser = value;
        index += 1;
        break;
      default:
        throw new Error(`Unknown argument: ${token}`);
    }
  }
  return parsed;
}

function ensureDir(targetPath) {
  fs.mkdirSync(targetPath, { recursive: true });
}

function resolveExistingPath(candidates, errorMessage) {
  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }
  throw new Error(errorMessage);
}

function startStaticServer(rootPath, host, port) {
  const server = http.createServer((request, response) => {
    const requestUrl = new URL(request.url || "/", `http://${host}:${port}`);
    const pathname = requestUrl.pathname === "/" ? "/index.html" : requestUrl.pathname;
    const filePath = path.normalize(path.join(rootPath, pathname));
    if (!filePath.startsWith(rootPath)) {
      response.statusCode = 403;
      response.end("forbidden");
      return;
    }

    if (!fs.existsSync(filePath) || fs.statSync(filePath).isDirectory()) {
      response.statusCode = 404;
      response.end("not found");
      return;
    }

    const contentTypes = {
      ".html": "text/html; charset=utf-8",
      ".js": "text/javascript; charset=utf-8",
      ".css": "text/css; charset=utf-8",
      ".json": "application/json; charset=utf-8",
      ".svg": "image/svg+xml",
      ".png": "image/png",
      ".jpg": "image/jpeg",
      ".jpeg": "image/jpeg",
      ".woff": "font/woff",
      ".woff2": "font/woff2",
    };
    response.setHeader("Content-Type", contentTypes[path.extname(filePath).toLowerCase()] || "application/octet-stream");
    fs.createReadStream(filePath).pipe(response);
  });

  return new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, host, () => resolve(server));
  });
}

function spawnLogged(fileName, args, options) {
  const child = spawn(fileName, args, {
    cwd: options.cwd,
    windowsHide: true,
    stdio: ["ignore", "pipe", "pipe"],
  });
  child.stdout.pipe(fs.createWriteStream(options.stdoutPath));
  child.stderr.pipe(fs.createWriteStream(options.stderrPath));
  return child;
}

async function waitForHealth(baseUrl, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  let lastError = null;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(`${baseUrl}/health`);
      if (response.ok) {
        return;
      }
      lastError = new Error(`health status ${response.status}`);
    } catch (error) {
      lastError = error;
    }
    await new Promise((resolve) => setTimeout(resolve, 1000));
  }
  throw new Error(`Utility healthcheck failed: ${lastError}`);
}

async function executeCommand(baseUrl, pairingToken, origin, profileId, commandId, args = {}, timeoutMilliseconds = 60000) {
  const response = await fetch(`${baseUrl}/commands/execute`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Origin: origin,
      "X-KWB-Pairing-Token": pairingToken,
    },
    body: JSON.stringify({
      profileId,
      commandId,
      arguments: args,
      timeoutMilliseconds,
      reportVerbosity: "Compact",
    }),
  });

  const payload = await response.json();
  const execution = payload?.payload ?? payload;
  if (!response.ok || !execution?.success) {
    throw new Error(`${commandId} failed: ${JSON.stringify(payload)}`);
  }
  return execution;
}

function parseStdoutJson(execution) {
  const result = execution?.result;
  if (!result?.hasExited) {
    throw new Error("Process result is missing.");
  }

  const stdout = String(result.stdout || "").trim();
  if (!stdout) {
    return {};
  }
  return JSON.parse(stdout);
}

function stopChild(child) {
  return new Promise((resolve) => {
    if (!child || child.exitCode !== null) {
      resolve();
      return;
    }

    child.once("exit", () => resolve());
    child.kill();
    setTimeout(() => {
      if (child.exitCode === null) {
        child.kill("SIGKILL");
      }
    }, 5000);
  });
}

async function main() {
  const args = parseArgs(process.argv);
  const stamp = nowStamp();
  const artifactRoot = path.join(outputRoot, stamp);
  const logsRoot = path.join(artifactRoot, "logs");
  const runtimeRoot = path.join(artifactRoot, "runtime");
  const workspaceRoot = path.join(artifactRoot, "workspace");
  const screenshotsRoot = path.join(artifactRoot, "screenshots");
  ensureDir(logsRoot);
  ensureDir(runtimeRoot);
  ensureDir(workspaceRoot);
  ensureDir(screenshotsRoot);

  const pagesHost = "127.0.0.1";
  const pagesPort = 5511;
  const utilityPort = 38741;
  const staticUrl = `http://${pagesHost}:${pagesPort}/index.html`;
  const pagesOrigin = `http://${pagesHost}:${pagesPort}`;
  const utilityUrl = `http://127.0.0.1:${utilityPort}`;
  const pairingToken = "kompas-pages-e2e";
  const utilityExePath = resolveExistingPath(utilityCandidates, "WebBridge.Utility.exe was not found.");
  const kompasSamplePath = resolveExistingPath(kompasSampleCandidates, "KOMPAS sample drawing was not found.");
  const outputTablePath = path.join(workspaceRoot, "table_M2.e2e.tbl");

  const runtime = buildRuntimeConfig({
    outputPath: path.join(runtimeRoot, "config.bootstrap.json"),
    listenUrl: utilityUrl,
    uiUrl: staticUrl,
    pairingToken,
    allowedOrigins: [
      pagesOrigin,
      `http://localhost:${pagesPort}`,
    ],
    logFilePath: path.join(logsRoot, "utility.log"),
    diagnosticsDirectory: path.join(runtimeRoot, "diagnostics"),
    profileDirectory: path.join(runtimeRoot, "profiles"),
    cacheDirectory: path.join(runtimeRoot, "cache"),
    environmentName: "PagesE2E",
    configVersion: `kompas-pages-e2e-${stamp}`,
  });
  fs.writeFileSync(runtime.outputPath, `${JSON.stringify(runtime.config, null, 2)}\n`, "utf8");

  const server = await startStaticServer(docsRoot, pagesHost, pagesPort);
  const utilityProcess = spawnLogged(
    utilityExePath,
    ["--config", runtime.outputPath],
    {
      cwd: path.dirname(utilityExePath),
      stdoutPath: path.join(logsRoot, "utility.stdout.log"),
      stderrPath: path.join(logsRoot, "utility.stderr.log"),
    },
  );

  let browser;
  const report = {
    success: false,
    artifactRoot,
    utilityUrl,
    staticUrl,
    outputTablePath,
  };

  try {
    await waitForHealth(utilityUrl, 90000);

    const launchOptions = args.browser === "chromium"
      ? { headless: true }
      : { channel: args.browser, headless: true };
    browser = await chromium.launch(launchOptions);
    const page = await browser.newPage({
      viewport: { width: 1440, height: 980 },
    });
    page.setDefaultTimeout(60000);

    await page.goto(
      `${staticUrl}?utilityUrl=${encodeURIComponent(utilityUrl)}&pairingToken=${encodeURIComponent(pairingToken)}&autoConnect=1&workspaceRoot=${encodeURIComponent(repoRoot)}`,
      { waitUntil: "networkidle" },
    );
    await page.screenshot({ path: path.join(screenshotsRoot, "00-loaded.png"), fullPage: true });

    await page.locator("#bridge-badge").filter({ hasText: "bridge online" }).waitFor();
    await page.locator("#runtime-badge").filter({ hasText: "runtime ready" }).waitFor();
    await page.locator('.module-host[data-module-id="xlsx-to-kompas-tbl"].is-active').waitFor();

    const tabCount = await page.locator(".rail__tab").count();
    if (tabCount < 3) {
      throw new Error(`Expected at least 3 tabs, got ${tabCount}`);
    }

    const openExecution = await executeCommand(
      utilityUrl,
      pairingToken,
      pagesOrigin,
      "kompas-pages-executor",
      "xlsx-to-kompas-tbl.open-document",
      {
        stdin: JSON.stringify({ documentPath: kompasSamplePath }),
      },
      60000,
    );
    const openPayload = parseStdoutJson(openExecution);
    if (!openPayload.success) {
      throw new Error(`Failed to open sample drawing: ${JSON.stringify(openPayload)}`);
    }

    await page.locator("#xlsx-refresh-status").click();
    await page.locator("#module-badge").filter({ hasText: "doc ok" }).waitFor();
    await page.screenshot({ path: path.join(screenshotsRoot, "10-status.png"), fullPage: true });

    await page.locator("#xlsx-file-input").setInputFiles(fixturePath);
    await page.locator("#xlsx-matrix-size").filter({ hasText: "8 x 13" }).waitFor();
    await page.locator("#xlsx-preview-table").filter({ hasText: "M2.2" }).waitFor();
    await page.screenshot({ path: path.join(screenshotsRoot, "20-uploaded.png"), fullPage: true });

    await page.locator("#xlsx-output-path").fill(outputTablePath);
    await page.locator("#xlsx-export-button").click();
    await page.locator("#xlsx-result-box").filter({ hasText: "OK" }).waitFor({ timeout: 180000 });
    await page.screenshot({ path: path.join(screenshotsRoot, "30-exported.png"), fullPage: true });

    if (!fs.existsSync(outputTablePath)) {
      throw new Error(`Exported table was not created: ${outputTablePath}`);
    }

    const outputStat = fs.statSync(outputTablePath);
    if (outputStat.size <= 0) {
      throw new Error("Exported table is empty.");
    }

    await page.locator("#xlsx-insert-button").click();
    await page.locator("#xlsx-result-box").filter({ hasText: "Inserted" }).waitFor({ timeout: 60000 });
    await page.screenshot({ path: path.join(screenshotsRoot, "40-inserted.png"), fullPage: true });

    const [download] = await Promise.all([
      page.waitForEvent("download"),
      page.locator("#xlsx-download-button").click(),
    ]);
    const downloadedPath = path.join(workspaceRoot, await download.suggestedFilename());
    await download.saveAs(downloadedPath);
    if (!fs.existsSync(downloadedPath) || fs.statSync(downloadedPath).size <= 0) {
      throw new Error("Downloaded table artifact is empty.");
    }

    await page.locator('.rail__tab[data-module-id="kompas-text-sync"]').click();
    await page.locator('.module-host[data-module-id="kompas-text-sync"].is-active').waitFor();
    await page.locator('.rail__tab[data-module-id="xlsx-to-kompas-tbl"]').click();
    await page.locator('.module-host[data-module-id="xlsx-to-kompas-tbl"].is-active').waitFor();
    await page.locator("#xlsx-matrix-size").filter({ hasText: "8 x 13" }).waitFor();
    await page.screenshot({ path: path.join(screenshotsRoot, "50-tab-switch.png"), fullPage: true });

    await page.locator("#xlsx-reset-button").click();
    await page.locator("#xlsx-matrix-size").filter({ hasText: "0 x 0" }).waitFor();
    await page.locator("#xlsx-download-button").isDisabled();
    await page.screenshot({ path: path.join(screenshotsRoot, "60-reset.png"), fullPage: true });

    Object.assign(report, {
      success: true,
      outputTableSize: outputStat.size,
      downloadedPath,
      downloadedSize: fs.statSync(downloadedPath).size,
      openDocumentPayload: openPayload,
    });
    fs.writeFileSync(path.join(artifactRoot, "report.json"), `${JSON.stringify(report, null, 2)}\n`, "utf8");
    process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
  } catch (error) {
    report.error = String(error.stack || error);
    fs.writeFileSync(path.join(artifactRoot, "report.json"), `${JSON.stringify(report, null, 2)}\n`, "utf8");
    throw error;
  } finally {
    if (browser) {
      await browser.close();
    }
    await stopChild(utilityProcess);
    await new Promise((resolve, reject) => server.close((closingError) => (closingError ? reject(closingError) : resolve())));
  }
}

main().catch((error) => {
  process.stderr.write(`${error.stack || error}\n`);
  process.exitCode = 1;
});
