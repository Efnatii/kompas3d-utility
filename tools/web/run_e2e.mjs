import fs from "node:fs";
import path from "node:path";
import http from "node:http";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";
import { buildRuntimeConfig } from "./build_runtime_config.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..", "..");
const webBridgeRepoRootCandidates = [
  "C:\\__SHARED_FOLDER__\\__GIT__\\web-bridge-utility",
  "C:\\__MY_PROJECTS__\\git\\web-bridge-utility",
];
const webBridgeRepoRoot = resolveExistingPath(
  webBridgeRepoRootCandidates,
  "web-bridge-utility repository was not found.",
);
const siteRoot = repoRoot;
const fixturePath = path.join(__dirname, "fixtures", "table_M2.xlsx");
const outputRoot = path.join(repoRoot, "out", "e2e");
const developmentConfigCandidates = [
  path.join(webBridgeRepoRoot, "configs", "config.development.sample.json"),
];

const utilityCandidates = [
  path.join("C:\\__LATENCY__\\__KOMPAS_3D__\\__UTILITY__", "web-bridge-utility-v1.0.0-preview.1-win-x64", "WebBridge.Utility.exe"),
  path.join(webBridgeRepoRoot, "artifacts", "publish", "utility", "win-x64", "WebBridge.Utility.exe"),
  path.join(webBridgeRepoRoot, "artifacts", "publish", "agent", "win-x64", "WebBridge.Utility.exe"),
  path.join(webBridgeRepoRoot, "src", "WebBridge.Utility", "bin", "Release", "net8.0", "win-x64", "WebBridge.Utility.exe"),
  path.join(webBridgeRepoRoot, "src", "WebBridge.Utility", "bin", "Debug", "net8.0", "win-x64", "WebBridge.Utility.exe"),
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
  const parsed = {
    browser: "msedge",
    utilityConfigMode: "development-flat",
    headed: false,
    pauseAfterInlineMs: 0,
    stopAfterInline: false,
  };
  for (let index = 2; index < argv.length; index += 1) {
    const token = argv[index];
    const value = argv[index + 1];
    switch (token) {
      case "--browser":
        parsed.browser = value;
        index += 1;
        break;
      case "--utility-exe":
        parsed.utilityExePath = value;
        index += 1;
        break;
      case "--utility-config-mode":
        if (!["bootstrap", "development-flat", "development-legacy"].includes(value)) {
          throw new Error(`Unknown utility config mode: ${value}`);
        }
        parsed.utilityConfigMode = value;
        index += 1;
        break;
      case "--headed":
        parsed.headed = true;
        break;
      case "--pause-after-inline-ms":
        parsed.pauseAfterInlineMs = Math.max(0, Number.parseInt(value, 10) || 0);
        index += 1;
        break;
      case "--stop-after-inline":
        parsed.stopAfterInline = true;
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

function uniqueStrings(values) {
  return [...new Set((values || []).filter(Boolean).map((value) => String(value).trim()).filter(Boolean))];
}

function cloneJson(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function loadDevelopmentComAdapter(adapterName) {
  const templatePath = resolveExistingPath(
    developmentConfigCandidates,
    "config.development.sample.json was not found.",
  );
  const template = JSON.parse(fs.readFileSync(templatePath, "utf8"));
  const adapter = (template.ComAdapters || []).find((candidate) =>
    String(candidate?.AdapterName || "").toLowerCase() === String(adapterName || "").toLowerCase());
  if (!adapter) {
    throw new Error(`Adapter '${adapterName}' was not found in ${templatePath}.`);
  }
  return cloneJson(adapter);
}

function buildDevelopmentRuntimeConfig(options = {}) {
  const templatePath = resolveExistingPath(
    developmentConfigCandidates,
    "config.development.sample.json was not found.",
  );
  const template = JSON.parse(fs.readFileSync(templatePath, "utf8"));
  const outputPath = options.outputPath || path.join(repoRoot, "out", "e2e", "config.development.json");

  template.Versions = {
    ...(template.Versions || {}),
    ConfigVersion: options.configVersion || `kompas-pages-e2e-dev-${Date.now()}`,
  };
  template.Runtime = {
    ...(template.Runtime || {}),
    EnvironmentName: options.environmentName || "PagesE2EDevelopment",
    DevMode: true,
    NoBrowser: true,
  };
  template.Server = {
    ...(template.Server || {}),
    ListenUrl: options.listenUrl,
  };
  template.Ui = {
    ...(template.Ui || {}),
    Url: options.uiUrl,
    OpenMode: "Never",
  };
  template.Logging = {
    ...(template.Logging || {}),
    FilePath: options.logFilePath,
  };
  template.Storage = {
    ...(template.Storage || {}),
    DiagnosticsDirectory: options.diagnosticsDirectory,
    CacheDirectory: options.cacheDirectory,
    ProfileDirectory: options.profileDirectory,
  };
  template.Security = {
    ...(template.Security || {}),
    PairingToken: options.pairingToken,
    AllowedOrigins: uniqueStrings(options.allowedOrigins),
  };

  return { outputPath, config: template };
}

function buildLegacyDevelopmentRuntimeConfig(options = {}) {
  const outputPath = options.outputPath || path.join(repoRoot, "out", "e2e", "config.development.legacy.json");
  return {
    outputPath,
    config: {
      Versions: {
        UtilityVersion: "1.0.0",
        ConfigVersion: options.configVersion || `kompas-pages-e2e-dev-legacy-${Date.now()}`,
        ConfigSchemaVersion: 2,
      },
      Metadata: {
        ProductName: "WebBridge.Utility",
        Author: "Гороховицкий Егор Русланович",
        Description: "Legacy nested development config for published WebBridge.Utility preview builds.",
      },
      Runtime: {
        EnvironmentName: options.environmentName || "PagesE2ELegacyDevelopment",
        DevMode: true,
        NoBrowser: true,
      },
      Server: {
        ListenUrl: options.listenUrl,
      },
      Ui: {
        Url: options.uiUrl,
        OpenMode: "Never",
        SessionWaitSeconds: 5,
      },
      Lifecycle: {
        ShutdownPolicy: "WhenIdle",
        IdleSeconds: 120,
      },
      Logging: {
        Level: "Debug",
        DebugMode: true,
        FilePath: options.logFilePath,
      },
      Storage: {
        ProfileDirectory: options.profileDirectory,
        CacheDirectory: options.cacheDirectory,
        DiagnosticsDirectory: options.diagnosticsDirectory,
      },
      Catalog: {
        Profiles: [],
      },
      Adapters: {
        Com: cloneJson(options.comAdapters || []),
        System: {},
      },
      Security: {
        LoopbackOnly: true,
        PairingToken: options.pairingToken,
        AllowedOrigins: uniqueStrings(options.allowedOrigins),
      },
      Session: {
        HeartbeatIntervalSeconds: 10,
        HeartbeatTimeoutSeconds: 30,
        PresenceTimeoutSeconds: 60,
        SuppressAutoOpenOnPresenceSessions: true,
        SweepIntervalSeconds: 2,
      },
    },
  };
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
  const pairingToken = args.utilityConfigMode === "bootstrap"
    ? "kompas-pages-local"
    : "replace-this-token";
  const utilityExePath = args.utilityExePath
    ? path.resolve(args.utilityExePath)
    : resolveExistingPath(utilityCandidates, "WebBridge.Utility.exe was not found.");
  const kompasSamplePath = resolveExistingPath(kompasSampleCandidates, "KOMPAS sample drawing was not found.");
  const outputTablePath = path.join(workspaceRoot, "table_M2.e2e.tbl");

  const runtimeOptions = {
    outputPath: path.join(
      runtimeRoot,
      args.utilityConfigMode === "bootstrap"
        ? "config.bootstrap.json"
        : args.utilityConfigMode === "development-flat"
          ? "config.development.json"
          : "config.development.legacy.json",
    ),
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
    environmentName: args.utilityConfigMode === "bootstrap"
      ? "PagesE2E"
      : args.utilityConfigMode === "development-flat"
        ? "PagesE2EDevelopment"
        : "PagesE2ELegacyDevelopment",
    configVersion: `kompas-pages-e2e-${args.utilityConfigMode}-${stamp}`,
    comAdapters: args.utilityConfigMode === "development-legacy"
      ? [loadDevelopmentComAdapter("kompas")]
      : [],
  };
  const runtime = args.utilityConfigMode === "bootstrap"
    ? buildRuntimeConfig(runtimeOptions)
    : args.utilityConfigMode === "development-flat"
      ? buildDevelopmentRuntimeConfig(runtimeOptions)
      : buildLegacyDevelopmentRuntimeConfig(runtimeOptions);
  fs.writeFileSync(runtime.outputPath, `${JSON.stringify(runtime.config, null, 2)}\n`, "utf8");

  const server = await startStaticServer(siteRoot, pagesHost, pagesPort);
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
    utilityConfigMode: args.utilityConfigMode,
    webBridgeRepoRoot,
    utilityExePath,
    utilityUrl,
    staticUrl,
    outputTablePath,
  };

  try {
    await waitForHealth(utilityUrl, 90000);

    const launchOptions = args.browser === "chromium"
      ? { headless: !args.headed }
      : { channel: args.browser, headless: !args.headed };
    browser = await chromium.launch(launchOptions);
    const page = await browser.newPage({
      viewport: { width: 1440, height: 980 },
    });
    page.setDefaultTimeout(60000);

    await page.goto(
      staticUrl,
      { waitUntil: "networkidle" },
    );
    await page.screenshot({ path: path.join(screenshotsRoot, "00-loaded.png"), fullPage: true });

    await page.locator("#bridge-badge").filter({ hasText: "bridge online" }).waitFor();
    await page.locator("#runtime-badge").filter({ hasText: "runtime ready" }).waitFor();
    await page.locator('.module-host[data-module-id="xlsx-to-kompas-tbl"].is-active').waitFor();
    if (await page.locator("#connect-button").count() !== 0) {
      throw new Error("Bridge connect button should be absent.");
    }

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
      { path: kompasSamplePath },
      60000,
    );
    if (!openExecution.result?.handleId) {
      throw new Error(`Failed to open sample drawing: ${JSON.stringify(openExecution)}`);
    }

    await page.locator("#module-badge").filter({ hasText: "doc ok" }).waitFor();
    await page.screenshot({ path: path.join(screenshotsRoot, "10-status.png"), fullPage: true });

    await page.locator("#xlsx-file-input").setInputFiles(fixturePath);
    await page.locator("#xlsx-matrix-size").filter({ hasText: "8 x 13" }).waitFor();
    await page.locator("#xlsx-preview-table").filter({ hasText: "M2.2" }).waitFor();
    const autoFollowPath = await page.locator("#xlsx-output-path").inputValue();
    if (!autoFollowPath || !autoFollowPath.toLowerCase().endsWith(".tbl")) {
      throw new Error(`Expected auto-follow output path, got: ${autoFollowPath}`);
    }
    await page.screenshot({ path: path.join(screenshotsRoot, "20-uploaded.png"), fullPage: true });

    await page.locator("#xlsx-output-path").fill(outputTablePath);
    const inlineCountBeforeExecution = await executeCommand(
      utilityUrl,
      pairingToken,
      pagesOrigin,
      "kompas-pages-executor",
      "xlsx-to-kompas-tbl.view-table-count",
      {},
      15000,
    );
    const inlineCountBefore = Number(inlineCountBeforeExecution.result) || 0;
    await page.locator("#xlsx-inline-button").click();
    await page.locator("#xlsx-result-box").filter({ hasText: "Inline |" }).waitFor({ timeout: 180000 });
    await page.screenshot({ path: path.join(screenshotsRoot, "25-inline.png"), fullPage: true });

    const inlineCountAfterExecution = await executeCommand(
      utilityUrl,
      pairingToken,
      pagesOrigin,
      "kompas-pages-executor",
      "xlsx-to-kompas-tbl.view-table-count",
      {},
      15000,
    );
    const inlineCountAfter = Number(inlineCountAfterExecution.result) || 0;
    if (inlineCountAfter !== inlineCountBefore + 1) {
      throw new Error(`Inline should add exactly one table: ${inlineCountBefore} -> ${inlineCountAfter}`);
    }
    if (fs.existsSync(outputTablePath)) {
      throw new Error(`Inline must not create a .tbl file: ${outputTablePath}`);
    }
    const inlineResultText = await page.locator("#xlsx-result-box").textContent();

    if (args.pauseAfterInlineMs > 0) {
      await page.waitForTimeout(args.pauseAfterInlineMs);
      await page.screenshot({ path: path.join(screenshotsRoot, "26-inline-paused.png"), fullPage: true });
    }

    if (args.stopAfterInline) {
      Object.assign(report, {
        success: true,
        autoFollowPath,
        inlineCountBefore,
        inlineCountAfter,
        inlineResultText,
        stopAfterInline: true,
        pauseAfterInlineMs: args.pauseAfterInlineMs,
      });
      fs.writeFileSync(path.join(artifactRoot, "report.json"), `${JSON.stringify(report, null, 2)}\n`, "utf8");
      process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
      return;
    }

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
    const insertResultText = await page.locator("#xlsx-result-box").textContent();

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
      autoFollowPath,
      inlineCountBefore,
      inlineCountAfter,
      inlineResultText,
      insertResultText,
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
