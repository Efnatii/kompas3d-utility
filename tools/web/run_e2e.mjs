import fs from "node:fs";
import path from "node:path";
import http from "node:http";
import { spawn, spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";
import { chromium } from "playwright";
import { buildRuntimeConfig } from "./build_runtime_config.mjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..", "..");
const webBridgeRepoRootCandidates = [
  "C:\\_GIT_\\web-bridge-utility",
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
const HIDDEN_SENTINEL_RE = /[\u200B\u2060\uFEFF]/g;
const API5_TEXT_LINE_ARRAY_TYPE = 3;
const API5_TEXT_ITEM_ARRAY_TYPE = 4;
const API5_TEXT_ITEM_STRING = 0;
const API5_TEXT_FLAG_ITALIC_ON = 0x40;
const API5_TEXT_FLAG_ITALIC_OFF = 0x80;
const API5_TEXT_FLAG_BOLD_ON = 0x100;
const API5_TEXT_FLAG_BOLD_OFF = 0x200;
const API5_TEXT_FLAG_UNDERLINE_ON = 0x400;
const API5_TEXT_FLAG_UNDERLINE_OFF = 0x800;
const COMMAND_PROOF_SAFE_THRESHOLD = 10;
const COMMAND_PROOF_SAFE_COMMANDS = Object.freeze([
  "xlsx-to-kompas-tbl.application.info",
  "xlsx-to-kompas-tbl.active-document",
  "xlsx-to-kompas-tbl.active-view",
  "xlsx-to-kompas-tbl.active-view-x",
  "xlsx-to-kompas-tbl.active-view-y",
  "xlsx-to-kompas-tbl.active-view-update",
  "xlsx-to-kompas-tbl.active-frame-center-x",
  "xlsx-to-kompas-tbl.active-frame-center-y",
  "xlsx-to-kompas-tbl.active-frame-refresh",
  "xlsx-to-kompas-tbl.view-table-count",
  "xlsx-to-kompas-tbl.create-table",
  "xlsx-to-kompas-tbl.table-cell-set-text",
  "xlsx-to-kompas-tbl.table-cell-clear-text",
  "xlsx-to-kompas-tbl.table-cell-set-one-line",
  "xlsx-to-kompas-tbl.table-cell-set-line-align",
  "xlsx-to-kompas-tbl.table-cell-add-line",
  "xlsx-to-kompas-tbl.table-cell-add-item",
  "xlsx-to-kompas-tbl.table-cell-add-item-before",
  "xlsx-to-kompas-tbl.table-cell-set-item",
  "xlsx-to-kompas-tbl.table-cell-get-text",
  "xlsx-to-kompas-tbl.table-cell-get-one-line",
  "xlsx-to-kompas-tbl.table-cell-get-line-count",
  "xlsx-to-kompas-tbl.table-cell-get-line-align",
  "xlsx-to-kompas-tbl.table-cell-get-line-item-count",
  "xlsx-to-kompas-tbl.table-cell-get-item-text",
  "xlsx-to-kompas-tbl.table-cell-get-item-font-name",
  "xlsx-to-kompas-tbl.table-cell-get-item-height",
  "xlsx-to-kompas-tbl.table-cell-get-item-bold",
  "xlsx-to-kompas-tbl.table-cell-get-item-italic",
  "xlsx-to-kompas-tbl.table-cell-get-item-underline",
  "xlsx-to-kompas-tbl.table-cell-get-item-color",
  "xlsx-to-kompas-tbl.table-cell-get-item-width-factor",
  "xlsx-to-kompas-tbl.table-save",
  "xlsx-to-kompas-tbl.table-update",
  "xlsx-to-kompas-tbl.table-delete",
  "xlsx-to-kompas-tbl.table-get-temp",
  "xlsx-to-kompas-tbl.table-get-valid",
  "xlsx-to-kompas-tbl.table-get-x",
  "xlsx-to-kompas-tbl.table-get-y",
  "xlsx-to-kompas-tbl.table-get-reference",
  "xlsx-to-kompas-tbl.table-open-by-reference",
  "xlsx-to-kompas-tbl.table-set-position",
  "xlsx-to-kompas-tbl.insert-table",
  "xlsx-to-kompas-tbl.api5-active-document2d",
  "xlsx-to-kompas-tbl.api5-create-text-param",
  "xlsx-to-kompas-tbl.api5-create-text-line-param",
  "xlsx-to-kompas-tbl.api5-create-text-item-param",
  "xlsx-to-kompas-tbl.api5-create-dynamic-array",
  "xlsx-to-kompas-tbl.api5-object-init",
  "xlsx-to-kompas-tbl.api5-text-param-get-line-array",
  "xlsx-to-kompas-tbl.api5-text-param-set-line-array",
  "xlsx-to-kompas-tbl.api5-text-line-param-get-item-array",
  "xlsx-to-kompas-tbl.api5-text-line-param-set-item-array",
  "xlsx-to-kompas-tbl.api5-text-item-param-get-font",
  "xlsx-to-kompas-tbl.api5-text-item-param-set-basic",
  "xlsx-to-kompas-tbl.api5-text-item-param-set-font",
  "xlsx-to-kompas-tbl.api5-text-item-font-set",
  "xlsx-to-kompas-tbl.api5-dynamic-array-count",
  "xlsx-to-kompas-tbl.api5-dynamic-array-add-item",
  "xlsx-to-kompas-tbl.api5-document-open-table",
  "xlsx-to-kompas-tbl.api5-document-end-obj",
  "xlsx-to-kompas-tbl.api5-document-clear-table-cell-text",
  "xlsx-to-kompas-tbl.api5-document-set-table-cell-text",
  "xlsx-to-kompas-tbl.document-set-active",
  "xlsx-to-kompas-tbl.document-get-path",
  "xlsx-to-kompas-tbl.document-get-active",
  "xlsx-to-kompas-tbl.document-get-visible",
]);
const COMMAND_PROOF_SINGLETON_THRESHOLDS = Object.freeze({
  "xlsx-to-kompas-tbl.open-document": 10,
});
const COMMAND_PROOF_THRESHOLDS = Object.freeze({
  ...Object.fromEntries(COMMAND_PROOF_SAFE_COMMANDS.map((commandId) => [commandId, COMMAND_PROOF_SAFE_THRESHOLD])),
  ...COMMAND_PROOF_SINGLETON_THRESHOLDS,
});

let activeCommandStats = null;

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
    scenario: "all",
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
      case "--scenario":
        if (!["all", "workflow", "rich-proof", "autofit-proof", "command-proof"].includes(value)) {
          throw new Error(`Unknown scenario: ${value}`);
        }
        parsed.scenario = value;
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
  recordCommandInvocation(commandId, "single");
  return execution;
}

async function executeBatch(baseUrl, pairingToken, origin, commands, options = {}) {
  const response = await fetch(`${baseUrl}/commands/execute-batch`, {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      Origin: origin,
      "X-KWB-Pairing-Token": pairingToken,
    },
    body: JSON.stringify({
      reportVerbosity: options.reportVerbosity || "Compact",
      stopOnError: options.stopOnError !== false,
      sharedContextId: options.sharedContextId || null,
      commands,
    }),
  });

  const payload = await response.json();
  const execution = payload?.payload ?? payload;
  if (!response.ok || !execution?.success) {
    throw new Error(`execute-batch failed: ${JSON.stringify(payload)}`);
  }
  for (const command of Array.isArray(commands) ? commands : []) {
    recordCommandInvocation(command?.commandId, "batch");
  }
  return execution;
}

function assertCondition(condition, message, details = undefined) {
  if (condition) {
    return;
  }
  const suffix = details === undefined ? "" : ` | ${JSON.stringify(details)}`;
  throw new Error(`${message}${suffix}`);
}

function normalizeWhitespace(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function coerceBooleanLike(value) {
  if (typeof value === "boolean") {
    return value;
  }
  if (typeof value === "number") {
    return value !== 0;
  }
  const normalized = String(value ?? "").trim().toLowerCase();
  if (!normalized || normalized === "0" || normalized === "false" || normalized === "no") {
    return false;
  }
  if (normalized === "1" || normalized === "true" || normalized === "yes" || normalized === "-1") {
    return true;
  }
  return Boolean(value);
}

function normalizeWindowsPathForComparison(value) {
  return path.win32.normalize(String(value || "").replace(/\//g, "\\")).toLowerCase();
}

function parseInlineHandleId(resultText) {
  const match = String(resultText || "").match(/\|\s*handle=([^\s|]+)\b/i);
  return match ? match[1].trim() : "";
}

function parseInlineTableReference(resultText) {
  const match = String(resultText || "").match(/\bref=(\d+)\b/i);
  return match ? Number.parseInt(match[1], 10) : Number.NaN;
}

function getInlineTempDirectory() {
  const tempRoot = process.env.TEMP
    || process.env.TMP
    || (process.env.LOCALAPPDATA ? path.join(process.env.LOCALAPPDATA, "Temp") : "");
  return tempRoot ? path.join(tempRoot, "kompas-pages", "inline") : "";
}

function listInlineTempArtifacts() {
  const directory = getInlineTempDirectory();
  if (!directory || !fs.existsSync(directory)) {
    return [];
  }
  return fs.readdirSync(directory, { withFileTypes: true })
    .filter((entry) => entry.isFile() && entry.name.toLowerCase().endsWith(".tbl"))
    .map((entry) => {
      const filePath = path.join(directory, entry.name);
      const stat = fs.statSync(filePath);
      return {
        name: entry.name,
        size: stat.size,
      };
    })
    .sort((left, right) => left.name.localeCompare(right.name));
}

function compareFiles(leftPath, rightPath) {
  return fs.readFileSync(leftPath).equals(fs.readFileSync(rightPath));
}

function requireHandleId(execution, label) {
  const handleId = String(execution?.result?.handleId || "");
  assertCondition(Boolean(handleId), `${label} did not return a handle`, execution);
  return handleId;
}

function expectApi5Success(result, label, details = undefined) {
  if (result !== false && result !== null && result !== undefined) {
    return;
  }
  assertCondition(false, `${label} failed`, details ?? result);
}

function coerceNumberLike(value) {
  const direct = Number(value);
  if (Number.isFinite(direct)) {
    return direct;
  }
  if (value && typeof value === "object") {
    for (const key of ["value", "Value", "reference", "Reference", "result", "Result", "rawValue", "RawValue", "currentValue", "CurrentValue"]) {
      const nested = Number(value[key]);
      if (Number.isFinite(nested)) {
        return nested;
      }
    }
  }
  return Number.NaN;
}

function toApi5TextFlags(item) {
  let flags = 0;
  flags |= item?.italic ? API5_TEXT_FLAG_ITALIC_ON : API5_TEXT_FLAG_ITALIC_OFF;
  flags |= item?.bold ? API5_TEXT_FLAG_BOLD_ON : API5_TEXT_FLAG_BOLD_OFF;
  flags |= item?.underline ? API5_TEXT_FLAG_UNDERLINE_ON : API5_TEXT_FLAG_UNDERLINE_OFF;
  return flags;
}

function toApi5CellNumber(rowIndex, columnIndex, columnCount) {
  return (Number(rowIndex) * Number(columnCount)) + Number(columnIndex) + 1;
}

function normalizeCellSnapshot(snapshot) {
  const normalizeTextContent = (value) => String(value ?? "").replace(HIDDEN_SENTINEL_RE, "");
  const lines = Array.isArray(snapshot?.lines)
    ? snapshot.lines.map((line) => ({
      align: Number(line?.align) || 0,
      items: Array.isArray(line?.items)
        ? line.items
          .map((item) => ({
            text: normalizeTextContent(item?.text),
            fontName: String(item?.fontName || ""),
            heightMm: Number((Number(item?.heightMm) || 0).toFixed(4)),
            bold: Boolean(item?.bold),
            italic: Boolean(item?.italic),
            underline: Boolean(item?.underline),
            color: Number(item?.color) || 0,
            widthFactor: Number((Number(item?.widthFactor) || 0).toFixed(4)),
          }))
          .filter((item) => item.text.length > 0)
        : [],
    }))
    : [];

  return {
    text: normalizeTextContent(snapshot?.text),
    oneLine: Boolean(snapshot?.oneLine),
    lineCount: lines.length,
    lines,
  };
}

function createExpectedRichCellMatrix() {
  return buildRichProofLayout().cellMatrix;
}

function convertStyleProofRunToSnapshotItem(run, chunkText) {
  return {
    text: String(chunkText ?? run?.text ?? ""),
    fontName: String(run?.fontName || DEFAULT_STYLE_PROOF_FONT.fontName),
    heightMm: Number((Number(run?.heightPt || DEFAULT_STYLE_PROOF_FONT.heightPt) * POINTS_TO_MM).toFixed(4)),
    bold: Boolean(run?.bold),
    italic: Boolean(run?.italic),
    underline: Boolean(run?.underline),
    color: Number.parseInt(resolveStyleProofColorHex(run?.colorSpec), 16),
    widthFactor: 1,
  };
}

function stripHeightMmFromNormalizedSnapshot(snapshot) {
  return {
    text: String(snapshot?.text || ""),
    oneLine: Boolean(snapshot?.oneLine),
    lineCount: Number(snapshot?.lineCount) || 0,
    lines: Array.isArray(snapshot?.lines)
      ? snapshot.lines.map((line) => ({
        align: Number(line?.align) || 0,
        items: Array.isArray(line?.items)
          ? line.items.map((item) => ({
            text: String(item?.text || ""),
            fontName: String(item?.fontName || ""),
            bold: Boolean(item?.bold),
            italic: Boolean(item?.italic),
            underline: Boolean(item?.underline),
            color: Number(item?.color) || 0,
            widthFactor: Number((Number(item?.widthFactor) || 0).toFixed(4)),
          }))
          : [],
      }))
      : [],
  };
}

function summarizeSnapshotHeights(snapshot) {
  const lineHeights = [];
  let itemCount = 0;
  let maxItemHeightMm = 0;
  for (const line of Array.isArray(snapshot?.lines) ? snapshot.lines : []) {
    let lineMaxHeight = 0;
    for (const item of Array.isArray(line?.items) ? line.items : []) {
      const heightMm = Number(item?.heightMm) || 0;
      if (heightMm > 0) {
        itemCount += 1;
      }
      if (heightMm > maxItemHeightMm) {
        maxItemHeightMm = heightMm;
      }
      if (heightMm > lineMaxHeight) {
        lineMaxHeight = heightMm;
      }
    }
    lineHeights.push(Number(lineMaxHeight.toFixed(4)));
  }
  const totalLineHeightMm = Number(lineHeights.reduce((sum, value) => sum + value, 0).toFixed(4));
  return {
    itemCount,
    lineCount: lineHeights.length,
    lineHeights,
    maxItemHeightMm: Number(maxItemHeightMm.toFixed(4)),
    totalLineHeightMm,
  };
}

function buildSnapshotHeightMetrics(snapshotMap) {
  const metrics = {};
  for (const [address, snapshot] of Object.entries(snapshotMap || {})) {
    metrics[address] = summarizeSnapshotHeights(snapshot);
  }
  return metrics;
}

function parseAdjustedCellCount(summaryText) {
  const match = String(summaryText || "").match(/adjusted=(\d+)/u);
  return match ? Number(match[1]) : 0;
}

function splitStyleProofRunsIntoLines(runs) {
  const lines = [{ items: [] }];
  for (const run of Array.isArray(runs) ? runs : []) {
    const chunks = String(run?.text ?? "").split(/\r\n|\r|\n/u);
    for (let index = 0; index < chunks.length; index += 1) {
      if (index > 0) {
        lines.push({ items: [] });
      }
      const chunk = chunks[index];
      if (chunk !== "" || chunks.length === 1) {
        lines[lines.length - 1].items.push(convertStyleProofRunToSnapshotItem(run, chunk));
      }
    }
  }
  return lines;
}

function encodeWorksheetColumnName(columnIndex) {
  let current = Number(columnIndex);
  let result = "";
  while (current >= 0) {
    result = String.fromCharCode(65 + (current % 26)) + result;
    current = Math.floor(current / 26) - 1;
  }
  return result;
}

function encodeWorksheetAddress(rowIndex, columnIndex) {
  return `${encodeWorksheetColumnName(columnIndex)}${rowIndex + 1}`;
}

function createStyleProofCell(caseDefinition, rowIndex, columnIndex, address) {
  const lines = splitStyleProofRunsIntoLines(caseDefinition.runs);
  return {
    label: caseDefinition.label,
    kind: caseDefinition.kind,
    address,
    rowIndex,
    columnIndex,
    text: caseDefinition.runs.map((run) => String(run?.text ?? "")).join(""),
    horizontal: String(caseDefinition.horizontal || ""),
    alignCode: STYLE_PROOF_ALIGN_CODE_BY_HORIZONTAL[String(caseDefinition.horizontal || "")] ?? 0,
    wrapText: Boolean(caseDefinition.wrapText),
    oneLine: !caseDefinition.wrapText && lines.length <= 1,
    hasContent: lines.some((line) => line.items.some((item) => item.text !== "")),
    lines,
    sourceRuns: caseDefinition.runs,
  };
}

function buildRichProofLayout() {
  const styleCases = buildStyleProofCaseCatalog();
  const cellMatrix = [];
  const expectedSnapshots = {};
  const cases = [];

  styleCases.forEach((styleCase, index) => {
    const rowIndex = Math.floor(index / STYLE_PROOF_COLUMN_COUNT);
    const columnIndex = index % STYLE_PROOF_COLUMN_COUNT;
    const address = encodeWorksheetAddress(rowIndex, columnIndex);
    const cell = createStyleProofCell(styleCase, rowIndex, columnIndex, address);
    if (!Array.isArray(cellMatrix[rowIndex])) {
      cellMatrix[rowIndex] = [];
    }
    cellMatrix[rowIndex][columnIndex] = cell;
    expectedSnapshots[address] = normalizeCellSnapshot({
      text: cell.text,
      oneLine: cell.oneLine,
      lineCount: cell.lines.length,
      lines: cell.lines.map((line) => ({
        align: cell.alignCode,
        items: line.items,
      })),
    });
    cases.push({
      label: styleCase.label,
      kind: styleCase.kind,
      address,
      rowIndex,
      columnIndex,
      horizontal: cell.horizontal,
      wrapText: cell.wrapText,
      oneLine: cell.oneLine,
      text: cell.text,
      lineCount: cell.lines.length,
      itemCount: cell.lines.reduce((total, line) => total + (Array.isArray(line.items) ? line.items.length : 0), 0),
      sourceRuns: styleCase.runs,
      styleIndex: 0,
    });
  });

  return {
    rows: Math.ceil(styleCases.length / STYLE_PROOF_COLUMN_COUNT),
    cols: STYLE_PROOF_COLUMN_COUNT,
    cases,
    cellMatrix,
    expectedSnapshots,
  };
}

function createFirstLineOnlyCellMatrix(cellMatrix) {
  return (Array.isArray(cellMatrix) ? cellMatrix : []).map((row) => (
    Array.isArray(row) ? row.map((cell) => {
      if (!cell?.hasContent || !Array.isArray(cell.lines) || !cell.lines.length) {
        return cell;
      }
      const firstLine = cell.lines[0];
      return {
        ...cell,
        text: firstLine.items.map((item) => String(item?.text || "")).join(""),
        lines: [firstLine],
        hasContent: firstLine.items.some((item) => String(item?.text || "") !== ""),
      };
    }) : []
  ));
}

function createAdditionalLinesOnlyCellMatrix(cellMatrix) {
  return (Array.isArray(cellMatrix) ? cellMatrix : []).map((row) => (
    Array.isArray(row) ? row.map((cell) => {
      if (!cell?.hasContent || !Array.isArray(cell.lines) || cell.lines.length <= 1) {
        return {
          ...cell,
          hasContent: false,
          lines: [],
        };
      }
      const extraLines = cell.lines.slice(1);
      return {
        ...cell,
        text: extraLines.map((line) => line.items.map((item) => String(item?.text || "")).join("")).join("\n"),
        lines: extraLines,
        hasContent: true,
        oneLine: false,
      };
    }) : []
  ));
}

function createExpectedRichSnapshots() {
  return buildRichProofLayout().expectedSnapshots;
}

function runPython(script) {
  const attempts = [
    ["py", ["-3", "-c", script]],
    ["python", ["-c", script]],
  ];
  let lastFailure = null;
  for (const [command, args] of attempts) {
    const result = spawnSync(command, args, {
      encoding: "utf8",
      windowsHide: true,
    });
    if (result.status === 0) {
      return result;
    }
    lastFailure = result;
  }
  throw new Error(`Python fixture generation failed: ${lastFailure?.stderr || lastFailure?.stdout || "unknown error"}`);
}

function runPowerShell(script) {
  const result = spawnSync("powershell", ["-NoProfile", "-Command", script], {
    encoding: "utf8",
    windowsHide: true,
  });
  if (result.status === 0) {
    return result;
  }
  throw new Error(`PowerShell fixture generation failed: ${result.stderr || result.stdout || "unknown error"}`);
}

function runPowerShellJson(script) {
  const result = runPowerShell(script);
  const text = String(result.stdout || "").trim();
  if (!text) {
    return null;
  }
  return JSON.parse(text);
}

function queryKompasProcesses() {
  const payload = runPowerShellJson(`
$items = Get-Process |
  Where-Object { $_.ProcessName -eq 'KOMPAS' } |
  Select-Object Id, ProcessName, MainWindowTitle, MainWindowHandle, StartTime
@($items) | ConvertTo-Json -Compress -Depth 4
`);
  return Array.isArray(payload) ? payload : (payload ? [payload] : []);
}

function assertSingleKompasProcessForHeadedRun() {
  const processes = queryKompasProcesses();
  if (processes.length !== 1) {
    const summary = processes.length
      ? processes.map((item) => `PID=${item.Id} title='${String(item.MainWindowTitle || "").trim() || "-"}'`).join("; ")
      : "none";
    throw new Error(`Headed visual run requires exactly one KOMPAS process. Found ${processes.length}: ${summary}. Close extra KOMPAS instances and rerun.`);
  }
  const processInfo = processes[0];
  if (!(Number(processInfo.MainWindowHandle) > 0)) {
    throw new Error(`The only KOMPAS process does not expose a visible main window: PID=${processInfo.Id}.`);
  }
  return processInfo;
}

function focusKompasWindow(processId) {
  return runPowerShellJson(`
Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class WinApi {
  [DllImport("user32.dll")]
  public static extern bool SetForegroundWindow(IntPtr hWnd);

  [DllImport("user32.dll")]
  public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);
}
"@
$proc = Get-Process -Id ${Number(processId)} -ErrorAction Stop
if ($proc.MainWindowHandle -eq 0) {
  throw "KOMPAS main window handle is 0."
}
[WinApi]::ShowWindowAsync([IntPtr]$proc.MainWindowHandle, 9) | Out-Null
Start-Sleep -Milliseconds 250
[WinApi]::SetForegroundWindow([IntPtr]$proc.MainWindowHandle) | Out-Null
[pscustomobject]@{
  Id = $proc.Id
  MainWindowHandle = $proc.MainWindowHandle
  MainWindowTitle = $proc.MainWindowTitle
} | ConvertTo-Json -Compress -Depth 3
`);
}

function escapeXmlText(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeXmlAttribute(value) {
  return escapeXmlText(value)
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}

function escapeWorksheetText(value) {
  return escapeXmlText(value).replace(/\r\n|\r|\n/g, "&#10;");
}

function needsPreservedWhitespace(text) {
  return /^[\s]|[\s]$| {2,}/.test(String(text ?? ""));
}

function encodeWorksheetTextNode(text) {
  const attrs = needsPreservedWhitespace(text) ? ' xml:space="preserve"' : "";
  return `<t${attrs}>${escapeWorksheetText(text)}</t>`;
}

function buildStyleProofColorAttributes(colorSpec) {
  const normalized = normalizeStyleProofColorSpec(colorSpec);
  if (normalized.type === "indexed") {
    return `indexed="${normalized.index}"`;
  }
  if (normalized.type === "auto") {
    return 'auto="1"';
  }
  return `rgb="${normalized.value}"`;
}

function buildStyleProofFontKey(run) {
  const normalizedColor = normalizeStyleProofColorSpec(run?.colorSpec);
  return [
    String(run?.fontName || DEFAULT_STYLE_PROOF_FONT.fontName),
    Number(run?.heightPt || DEFAULT_STYLE_PROOF_FONT.heightPt).toFixed(4),
    Boolean(run?.bold) ? "1" : "0",
    Boolean(run?.italic) ? "1" : "0",
    Boolean(run?.underline) ? "1" : "0",
    normalizedColor.type,
    normalizedColor.type === "indexed"
      ? String(normalizedColor.index)
      : (normalizedColor.type === "auto" ? "auto" : normalizedColor.value),
  ].join("|");
}

function buildStyleProofFontXml(run, tagName = "name") {
  const parts = [];
  if (run?.bold) {
    parts.push("<b/>");
  }
  if (run?.italic) {
    parts.push("<i/>");
  }
  if (run?.underline) {
    parts.push('<u val="single"/>');
  }
  parts.push(`<sz val="${Number(run?.heightPt || DEFAULT_STYLE_PROOF_FONT.heightPt)}"/>`);
  parts.push(`<color ${buildStyleProofColorAttributes(run?.colorSpec)}/>`);
  parts.push(`<${tagName} val="${escapeXmlAttribute(run?.fontName || DEFAULT_STYLE_PROOF_FONT.fontName)}"/>`);
  parts.push('<family val="2"/>');
  return parts.join("");
}

function buildStyleProofStylesXml(styleCases) {
  const fonts = [{ ...DEFAULT_STYLE_PROOF_FONT }];
  const fontIdByKey = new Map([[buildStyleProofFontKey(DEFAULT_STYLE_PROOF_FONT), 0]]);

  const ensureFontId = (run) => {
    const key = buildStyleProofFontKey(run);
    if (!fontIdByKey.has(key)) {
      fontIdByKey.set(key, fonts.length);
      fonts.push({
        fontName: String(run?.fontName || DEFAULT_STYLE_PROOF_FONT.fontName),
        heightPt: Number(run?.heightPt || DEFAULT_STYLE_PROOF_FONT.heightPt),
        bold: Boolean(run?.bold),
        italic: Boolean(run?.italic),
        underline: Boolean(run?.underline),
        colorSpec: normalizeStyleProofColorSpec(run?.colorSpec),
      });
    }
    return fontIdByKey.get(key);
  };

  const xfs = [{
    fontId: 0,
    horizontal: "",
    wrapText: false,
  }];
  const xfIdByKey = new Map([["0||0", 0]]);

  for (const styleCase of styleCases) {
    const plainFont = styleCase.kind === "plain" ? styleCase.sourceRuns[0] : DEFAULT_STYLE_PROOF_FONT;
    const fontId = ensureFontId(plainFont);
    const horizontal = String(styleCase.horizontal || "");
    const wrapText = Boolean(styleCase.wrapText);
    const xfKey = `${fontId}|${horizontal}|${wrapText ? 1 : 0}`;
    if (!xfIdByKey.has(xfKey)) {
      xfIdByKey.set(xfKey, xfs.length);
      xfs.push({ fontId, horizontal, wrapText });
    }
    styleCase.styleIndex = xfIdByKey.get(xfKey);
  }

  const fontXml = fonts.map((font) => `<font>${buildStyleProofFontXml(font)}</font>`).join("\n    ");
  const xfXml = xfs.map((xf) => {
    const applyFont = xf.fontId !== 0 ? ' applyFont="1"' : "";
    const hasAlignment = Boolean(xf.horizontal) || xf.wrapText;
    if (!hasAlignment) {
      return `<xf numFmtId="0" fontId="${xf.fontId}" fillId="0" borderId="0" xfId="0"${applyFont}/>`;
    }
    const horizontalAttr = xf.horizontal ? ` horizontal="${escapeXmlAttribute(xf.horizontal)}"` : "";
    return `<xf numFmtId="0" fontId="${xf.fontId}" fillId="0" borderId="0" xfId="0"${applyFont} applyAlignment="1"><alignment${horizontalAttr} wrapText="${xf.wrapText ? 1 : 0}"/></xf>`;
  }).join("\n    ");
  const indexedColorXml = STYLE_PROOF_INDEXED_COLORS
    .map((rgb) => `<rgbColor rgb="${rgb}"/>`)
    .join("");

  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <colors>
    <indexedColors>${indexedColorXml}</indexedColors>
  </colors>
  <fonts count="${fonts.length}">
    ${fontXml}
  </fonts>
  <fills count="2">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
  </cellStyleXfs>
  <cellXfs count="${xfs.length}">
    ${xfXml}
  </cellXfs>
  <cellStyles count="1">
    <cellStyle name="Normal" xfId="0" builtinId="0"/>
  </cellStyles>
</styleSheet>
`;
}

function buildStyleProofCellXml(styleCase) {
  if (styleCase.kind === "rich") {
    const richRunsXml = styleCase.sourceRuns.map((run) => (
      `<r><rPr>${buildStyleProofFontXml(run, "rFont")}</rPr>${encodeWorksheetTextNode(run.text)}</r>`
    )).join("");
    return `<c r="${styleCase.address}" s="${styleCase.styleIndex}" t="inlineStr"><is>${richRunsXml}</is></c>`;
  }
  return `<c r="${styleCase.address}" s="${styleCase.styleIndex}" t="inlineStr"><is>${encodeWorksheetTextNode(styleCase.text)}</is></c>`;
}

function buildStyleProofWorksheetXml(styleCases, rows, cols) {
  const rowXml = [];
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const cells = styleCases
      .filter((styleCase) => styleCase.rowIndex === rowIndex)
      .sort((left, right) => left.columnIndex - right.columnIndex);
    rowXml.push(`<row r="${rowIndex + 1}">${cells.map((cell) => buildStyleProofCellXml(cell)).join("")}</row>`);
  }
  const lastAddress = encodeWorksheetAddress(rows - 1, cols - 1);
  const colsXml = Array.from({ length: cols }, (_, index) => (
    `<col min="${index + 1}" max="${index + 1}" width="24" customWidth="1"/>`
  )).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:${lastAddress}"/>
  <cols>${colsXml}</cols>
  <sheetData>${rowXml.join("")}</sheetData>
</worksheet>
`;
}

function createRichFixtureArchive(workspaceRoot) {
  const layout = buildRichProofLayout();
  const fixturePath = path.join(workspaceRoot, "fixture-rich.xlsx");
  const tempRoot = path.join(workspaceRoot, `fixture-rich-${Date.now()}`);
  const stylesXml = buildStyleProofStylesXml(layout.cases);
  const worksheetXml = buildStyleProofWorksheetXml(layout.cases, layout.rows, layout.cols);
  const files = new Map([
    ["[Content_Types].xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
`],
    ["_rels/.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
`],
    ["xl/workbook.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
    <sheet name="RichProof" sheetId="1" r:id="rId1"/>
  </sheets>
</workbook>
`],
    ["xl/_rels/workbook.xml.rels", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>
`],
    ["xl/styles.xml", stylesXml],
    ["xl/worksheets/sheet1.xml", worksheetXml],
  ]);

  fs.rmSync(tempRoot, { recursive: true, force: true });
  for (const [relativePath, content] of files.entries()) {
    const targetPath = path.join(tempRoot, ...relativePath.split("/"));
    fs.mkdirSync(path.dirname(targetPath), { recursive: true });
    fs.writeFileSync(targetPath, content, "utf8");
  }
  fs.rmSync(fixturePath, { force: true });
  const escapedSource = tempRoot.replace(/'/g, "''");
  const escapedTarget = fixturePath.replace(/'/g, "''");
  try {
    runPowerShell(`
Add-Type -AssemblyName System.IO.Compression.FileSystem
[System.IO.Compression.ZipFile]::CreateFromDirectory('${escapedSource}', '${escapedTarget}')
`);
  } finally {
    fs.rmSync(tempRoot, { recursive: true, force: true });
  }
  assertCondition(fs.existsSync(fixturePath), "Rich fixture archive was not created", { fixturePath });
  return fixturePath;
}

function createRichFixtureWithFallback(workspaceRoot) {
  return createRichFixtureArchive(workspaceRoot);
}

function createRichFixture(workspaceRoot) {
  return createRichFixtureArchive(workspaceRoot);
}

async function readCellSnapshot(baseUrl, pairingToken, origin, handleId, rowIndex, columnIndex) {
  const baseArgs = { handleId, rowIndex, columnIndex };
  const textExecution = await executeCommand(
    baseUrl,
    pairingToken,
    origin,
    "kompas-pages-executor",
    "xlsx-to-kompas-tbl.table-cell-get-text",
    baseArgs,
    15000,
  );
  const oneLineExecution = await executeCommand(
    baseUrl,
    pairingToken,
    origin,
    "kompas-pages-executor",
    "xlsx-to-kompas-tbl.table-cell-get-one-line",
    baseArgs,
    15000,
  );
  const lineCountExecution = await executeCommand(
    baseUrl,
    pairingToken,
    origin,
    "kompas-pages-executor",
    "xlsx-to-kompas-tbl.table-cell-get-line-count",
    baseArgs,
    15000,
  );
  const lineCount = Number(lineCountExecution.result) || 0;
  const lines = [];
  for (let lineIndex = 0; lineIndex < lineCount; lineIndex += 1) {
    const alignExecution = await executeCommand(
      baseUrl,
      pairingToken,
      origin,
      "kompas-pages-executor",
      "xlsx-to-kompas-tbl.table-cell-get-line-align",
      { ...baseArgs, lineIndex },
      15000,
    );
    const itemCountExecution = await executeCommand(
      baseUrl,
      pairingToken,
      origin,
      "kompas-pages-executor",
      "xlsx-to-kompas-tbl.table-cell-get-line-item-count",
      { ...baseArgs, lineIndex },
      15000,
    );
    const itemCount = Number(itemCountExecution.result) || 0;
    const items = [];
    for (let itemIndex = 0; itemIndex < itemCount; itemIndex += 1) {
      const itemArgs = { ...baseArgs, lineIndex, itemIndex };
      const [
        textResult,
        fontNameResult,
        heightResult,
        boldResult,
        italicResult,
        underlineResult,
        colorResult,
        widthFactorResult,
      ] = await Promise.all([
        executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-text", itemArgs, 15000),
        executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-font-name", itemArgs, 15000),
        executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-height", itemArgs, 15000),
        executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-bold", itemArgs, 15000),
        executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-italic", itemArgs, 15000),
        executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-underline", itemArgs, 15000),
        executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-color", itemArgs, 15000),
        executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-width-factor", itemArgs, 15000),
      ]);
      items.push({
        text: String(textResult.result || ""),
        fontName: String(fontNameResult.result || ""),
        heightMm: Number(heightResult.result) || 0,
        bold: coerceBooleanLike(boldResult.result),
        italic: coerceBooleanLike(italicResult.result),
        underline: coerceBooleanLike(underlineResult.result),
        color: Number(colorResult.result) || 0,
        widthFactor: Number(widthFactorResult.result) || 0,
      });
    }
    lines.push({
      align: Number(alignExecution.result) || 0,
      items,
    });
  }
  return {
    text: String(textExecution.result || ""),
    oneLine: coerceBooleanLike(oneLineExecution.result),
    lineCount,
    lines,
  };
}

const POINTS_TO_MM = 25.4 / 72;
const STYLE_PROOF_COLUMN_COUNT = 5;
const STYLE_PROOF_INDEXED_COLORS = [
  "000000",
  "FFFFFF",
  "FF0000",
  "00AA00",
  "0000FF",
  "FF6600",
  "800080",
  "008080",
  "666666",
  "800000",
  "006400",
  "1F497D",
  "C0504D",
  "4F81BD",
];

function recordCommandInvocation(commandId, source = "single") {
  if (!activeCommandStats || !commandId) {
    return;
  }
  const key = String(commandId);
  const bucket = activeCommandStats[key] || {
    count: 0,
    singleCount: 0,
    batchCount: 0,
  };
  bucket.count += 1;
  if (source === "batch") {
    bucket.batchCount += 1;
  } else {
    bucket.singleCount += 1;
  }
  activeCommandStats[key] = bucket;
}

function snapshotCommandCounts(commandStats = activeCommandStats) {
  return Object.fromEntries(
    Object.entries(commandStats || {}).map(([commandId, value]) => [commandId, Number(value?.count) || 0]),
  );
}
const DEFAULT_STYLE_PROOF_FONT = Object.freeze({
  fontName: "Calibri",
  heightPt: 11,
  bold: false,
  italic: false,
  underline: false,
  colorSpec: { type: "rgb", value: "000000" },
});
const STYLE_PROOF_ALIGN_CODE_BY_HORIZONTAL = {
  "": 0,
  left: 0,
  center: 1,
  right: 2,
};

function normalizeHex6(value) {
  const normalized = String(value || "").replace(/[^0-9a-f]/gi, "").toUpperCase();
  if (normalized.length === 8) {
    return normalized.slice(2);
  }
  if (normalized.length === 6) {
    return normalized;
  }
  return "000000";
}

function createRgbColorSpec(value) {
  return {
    type: "rgb",
    value: normalizeHex6(value),
  };
}

function createIndexedColorSpec(index) {
  return {
    type: "indexed",
    index: Number(index),
  };
}

function createAutoColorSpec() {
  return {
    type: "auto",
  };
}

function normalizeStyleProofColorSpec(colorSpec) {
  if (colorSpec?.type === "indexed") {
    return createIndexedColorSpec(colorSpec.index);
  }
  if (colorSpec?.type === "auto") {
    return createAutoColorSpec();
  }
  return createRgbColorSpec(colorSpec?.value || colorSpec?.rgb || "000000");
}

function resolveStyleProofColorHex(colorSpec) {
  const normalized = normalizeStyleProofColorSpec(colorSpec);
  if (normalized.type === "indexed") {
    return STYLE_PROOF_INDEXED_COLORS[normalized.index] || "000000";
  }
  if (normalized.type === "auto") {
    return "000000";
  }
  return normalized.value;
}

function createStyleProofRun(text, fontName, heightPt, options = {}) {
  const parsedHeight = Number.parseFloat(heightPt);
  return {
    text: String(text ?? ""),
    fontName: String(fontName || DEFAULT_STYLE_PROOF_FONT.fontName),
    heightPt: Number.isFinite(parsedHeight) && parsedHeight > 0 ? parsedHeight : DEFAULT_STYLE_PROOF_FONT.heightPt,
    bold: Boolean(options.bold),
    italic: Boolean(options.italic),
    underline: Boolean(options.underline),
    colorSpec: normalizeStyleProofColorSpec(options.colorSpec || DEFAULT_STYLE_PROOF_FONT.colorSpec),
  };
}

function createStyleProofPlainCase(label, text, fontName, heightPt, options = {}) {
  return {
    kind: "plain",
    label: String(label || ""),
    horizontal: String(options.horizontal || ""),
    wrapText: Boolean(options.wrapText),
    runs: [
      createStyleProofRun(text, fontName, heightPt, options),
    ],
  };
}

function createStyleProofRichCase(label, runs, options = {}) {
  return {
    kind: "rich",
    label: String(label || ""),
    horizontal: String(options.horizontal || ""),
    wrapText: Boolean(options.wrapText),
    runs: (Array.isArray(runs) ? runs : []).map((run) => createStyleProofRun(
      run?.text ?? "",
      run?.fontName,
      run?.heightPt,
      run,
    )),
  };
}

function buildStyleProofCaseCatalog() {
  return [
    createStyleProofRichCase("center-rich-primary", [
      { text: "Bold", fontName: "Arial", heightPt: 14, bold: true, colorSpec: createRgbColorSpec("FF0000") },
      { text: " + ", fontName: "Calibri", heightPt: 11, colorSpec: createRgbColorSpec("000000") },
      { text: "Blue\nItalic", fontName: "Arial", heightPt: 12, italic: true, underline: true, colorSpec: createRgbColorSpec("0000FF") },
    ], { horizontal: "center", wrapText: true }),
    createStyleProofPlainCase("right-plain-green", "Plain styled", "Arial", 13, {
      horizontal: "right",
      wrapText: false,
      bold: true,
      italic: true,
      underline: true,
      colorSpec: createRgbColorSpec("008000"),
    }),
    createStyleProofPlainCase("mono-wrap-triple", "Mono wrap\nSecond line\nThird", "Courier New", 10, {
      horizontal: "",
      wrapText: true,
      colorSpec: createRgbColorSpec("666666"),
    }),
    createStyleProofRichCase("right-rich-serif-mono", [
      { text: "Times", fontName: "Times New Roman", heightPt: 16, colorSpec: createRgbColorSpec("0000FF") },
      { text: " / ", fontName: "Calibri", heightPt: 11, colorSpec: createRgbColorSpec("000000") },
      { text: "Courier", fontName: "Courier New", heightPt: 11, bold: true, underline: true, colorSpec: createRgbColorSpec("800000") },
      { text: "\nTail", fontName: "Arial", heightPt: 9, italic: true, colorSpec: createRgbColorSpec("008000") },
    ], { horizontal: "right", wrapText: true }),
    createStyleProofPlainCase("tiny-calibri-auto", "Tiny Calibri", "Calibri", 8, {
      horizontal: "left",
      wrapText: false,
      colorSpec: createAutoColorSpec(),
    }),
    createStyleProofPlainCase("verdana-bold-orange", "Wide Verdana", "Verdana", 18, {
      horizontal: "center",
      wrapText: false,
      bold: true,
      colorSpec: createRgbColorSpec("FF6600"),
    }),
    createStyleProofPlainCase("segoe-italic-teal", "Segoe Italic", "Segoe UI", 12, {
      horizontal: "right",
      wrapText: false,
      italic: true,
      colorSpec: createRgbColorSpec("008080"),
    }),
    createStyleProofPlainCase("tahoma-wrap-flag", "Wrap flag only", "Tahoma", 11, {
      horizontal: "left",
      wrapText: true,
      underline: true,
      colorSpec: createIndexedColorSpec(6),
    }),
    createStyleProofPlainCase("georgia-two-line", "Georgia line\nSecond", "Georgia", 15, {
      horizontal: "center",
      wrapText: true,
      colorSpec: createIndexedColorSpec(13),
    }),
    createStyleProofPlainCase("consolas-code", "Code 12345", "Consolas", 10, {
      horizontal: "right",
      wrapText: false,
      bold: true,
      colorSpec: createIndexedColorSpec(2),
    }),
    createStyleProofRichCase("cambria-mix", [
      { text: "Cambria", fontName: "Cambria", heightPt: 13, bold: true, colorSpec: createIndexedColorSpec(11) },
      { text: " mix", fontName: "Calibri", heightPt: 11, colorSpec: createRgbColorSpec("000000") },
    ], { horizontal: "left", wrapText: false }),
    createStyleProofRichCase("gost-spec", [
      { text: "ГОСТ", fontName: "Arial", heightPt: 12, bold: true, colorSpec: createRgbColorSpec("C0504D") },
      { text: " / ", fontName: "Calibri", heightPt: 11, colorSpec: createRgbColorSpec("000000") },
      { text: "Spec", fontName: "Times New Roman", heightPt: 12, italic: true, colorSpec: createIndexedColorSpec(11) },
    ], { horizontal: "center", wrapText: false }),
    createStyleProofRichCase("triple-line-rich", [
      { text: "L1", fontName: "Segoe UI", heightPt: 11, colorSpec: createRgbColorSpec("008080") },
      { text: "\nL2", fontName: "Courier New", heightPt: 11, bold: true, colorSpec: createRgbColorSpec("800000") },
      { text: "\nL3", fontName: "Arial", heightPt: 10, italic: true, colorSpec: createRgbColorSpec("4F81BD") },
    ], { horizontal: "right", wrapText: true }),
    createStyleProofPlainCase("serif-underlined", "Underlined serif", "Times New Roman", 11, {
      horizontal: "left",
      wrapText: false,
      underline: true,
      colorSpec: createIndexedColorSpec(9),
    }),
    createStyleProofPlainCase("segoe-combo", "Segoe combo", "Segoe UI", 14, {
      horizontal: "center",
      wrapText: false,
      bold: true,
      italic: true,
      colorSpec: createIndexedColorSpec(10),
    }),
    createStyleProofPlainCase("cambria-wrap-flag", "Wrap still on", "Cambria", 16, {
      horizontal: "right",
      wrapText: true,
      colorSpec: createRgbColorSpec("1F497D"),
    }),
    createStyleProofRichCase("space-preserve", [
      { text: "A  ", fontName: "Arial", heightPt: 11, colorSpec: createRgbColorSpec("000000") },
      { text: "B  ", fontName: "Consolas", heightPt: 11, bold: true, colorSpec: createIndexedColorSpec(5) },
      { text: "C", fontName: "Arial", heightPt: 11, colorSpec: createRgbColorSpec("000000") },
    ], { horizontal: "left", wrapText: false }),
    createStyleProofPlainCase("arial-two-line", "Line 1\nLine 2", "Arial", 9, {
      horizontal: "center",
      wrapText: true,
      colorSpec: createRgbColorSpec("000000"),
    }),
    createStyleProofRichCase("rgb-runs", [
      { text: "R", fontName: "Arial", heightPt: 12, bold: true, colorSpec: createIndexedColorSpec(2) },
      { text: "G", fontName: "Arial", heightPt: 12, bold: true, colorSpec: createIndexedColorSpec(3) },
      { text: "B", fontName: "Arial", heightPt: 12, bold: true, colorSpec: createIndexedColorSpec(4) },
      { text: "K", fontName: "Arial", heightPt: 12, bold: true, colorSpec: createAutoColorSpec() },
    ], { horizontal: "right", wrapText: false }),
    createStyleProofPlainCase("mono-emphasis", "Mono emphasis", "Courier New", 14, {
      horizontal: "left",
      wrapText: false,
      bold: true,
      italic: true,
      colorSpec: createIndexedColorSpec(8),
    }),
    createStyleProofRichCase("size-status", [
      { text: "Размер 25", fontName: "Times New Roman", heightPt: 16, colorSpec: createIndexedColorSpec(13) },
      { text: "\nOK", fontName: "Arial", heightPt: 10, bold: true, colorSpec: createIndexedColorSpec(10) },
    ], { horizontal: "center", wrapText: true }),
    createStyleProofPlainCase("verdana-gray", "Light Verdana", "Verdana", 9, {
      horizontal: "right",
      wrapText: false,
      colorSpec: createIndexedColorSpec(8),
    }),
    createStyleProofPlainCase("cyrillic-plain", "Материал сталь 09Г2С", "Calibri", 12, {
      horizontal: "left",
      wrapText: false,
      bold: true,
      colorSpec: createIndexedColorSpec(11),
    }),
    createStyleProofRichCase("three-font-stack", [
      { text: "Top\n", fontName: "Tahoma", heightPt: 11, colorSpec: createIndexedColorSpec(7) },
      { text: "Mid\n", fontName: "Georgia", heightPt: 13, underline: true, colorSpec: createIndexedColorSpec(6) },
      { text: "Bot", fontName: "Consolas", heightPt: 11, bold: true, colorSpec: createIndexedColorSpec(5) },
    ], { horizontal: "center", wrapText: true }),
    createStyleProofPlainCase("final-tahoma", "Final case", "Tahoma", 13, {
      horizontal: "right",
      wrapText: false,
      bold: true,
      underline: true,
      colorSpec: createIndexedColorSpec(11),
    }),
  ];
}

function flattenCellMatrix(cellMatrix) {
  const cells = [];
  for (const row of Array.isArray(cellMatrix) ? cellMatrix : []) {
    for (const cell of Array.isArray(row) ? row : []) {
      if (cell?.address && cell.hasContent) {
        cells.push(cell);
      }
    }
  }
  return cells;
}

async function readSnapshotMap(baseUrl, pairingToken, origin, handleId, cells) {
  const raw = {};
  const normalized = {};
  for (const cell of cells) {
    const snapshot = await readCellSnapshot(
      baseUrl,
      pairingToken,
      origin,
      handleId,
      cell.rowIndex,
      cell.columnIndex,
    );
    raw[cell.address] = snapshot;
    normalized[cell.address] = normalizeCellSnapshot(snapshot);
  }
  return { raw, normalized };
}

async function readTableReference(baseUrl, pairingToken, origin, tableHandleId) {
  const prepareReferenceExecution = await executeCommand(
    baseUrl,
    pairingToken,
    origin,
    "kompas-pages-executor",
    "xlsx-to-kompas-tbl.table-update",
    { handleId: tableHandleId },
    30000,
  );
  assertCondition(prepareReferenceExecution.result !== false, "KOMPAS returned Update=false before reading table reference.", prepareReferenceExecution);
  const tableReferenceExecution = await executeCommand(
    baseUrl,
    pairingToken,
    origin,
    "kompas-pages-executor",
    "xlsx-to-kompas-tbl.table-get-reference",
    { handleId: tableHandleId },
    15000,
  );
  const tableReference = coerceNumberLike(tableReferenceExecution.result);
  assertCondition(Number.isFinite(tableReference) && tableReference > 0, "Invalid table reference", tableReferenceExecution);
  return tableReference;
}

async function reopenTableHandleByReference(baseUrl, pairingToken, origin, tableReference) {
  const reopenExecution = await executeCommand(
    baseUrl,
    pairingToken,
    origin,
    "kompas-pages-executor",
    "xlsx-to-kompas-tbl.table-open-by-reference",
    {
      tableReference,
      refresh: true,
    },
    30000,
  );
  return requireHandleId(reopenExecution, "table-open-by-reference");
}

async function readActiveDocumentInfo(baseUrl, pairingToken, origin) {
  const execution = await executeCommand(
    baseUrl,
    pairingToken,
    origin,
    "kompas-pages-executor",
    "xlsx-to-kompas-tbl.active-document",
    { refresh: true },
    15000,
  );
  return execution.result || null;
}

async function readDocumentHandleState(baseUrl, pairingToken, origin, handleId) {
  const [pathExecution, activeExecution, visibleExecution] = await Promise.all([
    executeCommand(
      baseUrl,
      pairingToken,
      origin,
      "kompas-pages-executor",
      "xlsx-to-kompas-tbl.document-get-path",
      { handleId },
      15000,
    ),
    executeCommand(
      baseUrl,
      pairingToken,
      origin,
      "kompas-pages-executor",
      "xlsx-to-kompas-tbl.document-get-active",
      { handleId },
      15000,
    ),
    executeCommand(
      baseUrl,
      pairingToken,
      origin,
      "kompas-pages-executor",
      "xlsx-to-kompas-tbl.document-get-visible",
      { handleId },
      15000,
    ),
  ]);
  return {
    path: String(pathExecution.result || ""),
    active: coerceBooleanLike(activeExecution.result),
    visible: coerceBooleanLike(visibleExecution.result),
  };
}

async function waitForActiveDocumentPath(baseUrl, pairingToken, origin, expectedPath, timeoutMs = 20000) {
  const expected = normalizeWindowsPathForComparison(expectedPath);
  const startedAt = Date.now();
  let lastDocument = null;
  while (Date.now() - startedAt <= timeoutMs) {
    lastDocument = await readActiveDocumentInfo(baseUrl, pairingToken, origin);
    if (normalizeWindowsPathForComparison(lastDocument?.path) === expected) {
      return lastDocument;
    }
    await new Promise((resolve) => setTimeout(resolve, 500));
  }
  throw new Error(`Active document path mismatch. Expected ${expectedPath}, got ${lastDocument?.path || "<empty>"}`);
}

function createApi7AdditionalLineCommands(handleId, cellMatrix) {
  const rows = Array.isArray(cellMatrix) ? cellMatrix.length : 0;
  const cols = Array.isArray(cellMatrix?.[0]) ? cellMatrix[0].length : 0;
  const commands = [];
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const row = Array.isArray(cellMatrix[rowIndex]) ? cellMatrix[rowIndex] : [];
    for (let columnIndex = 0; columnIndex < cols; columnIndex += 1) {
      const cell = row[columnIndex];
      if (!cell?.hasContent || !Array.isArray(cell.lines) || !cell.lines.length) {
        continue;
      }
      commands.push({
        commandId: "xlsx-to-kompas-tbl.table-cell-set-one-line",
        arguments: {
          handleId,
          rowIndex: cell.rowIndex,
          columnIndex: cell.columnIndex,
          oneLine: false,
        },
        timeoutMilliseconds: 30000,
        profileId: "kompas-pages-executor",
      });
      cell.lines.forEach((line, extraLineIndex) => {
        const lineIndex = extraLineIndex + 1;
        commands.push({
          commandId: "xlsx-to-kompas-tbl.table-cell-add-line",
          arguments: {
            handleId,
            rowIndex: cell.rowIndex,
            columnIndex: cell.columnIndex,
            align: cell.alignCode,
          },
          timeoutMilliseconds: 30000,
          profileId: "kompas-pages-executor",
        });
        line.items.forEach((item, itemIndex) => {
          commands.push({
            commandId: itemIndex === 0
              ? "xlsx-to-kompas-tbl.table-cell-add-item-before"
              : "xlsx-to-kompas-tbl.table-cell-add-item",
            arguments: {
              handleId,
              rowIndex: cell.rowIndex,
              columnIndex: cell.columnIndex,
              lineIndex,
              itemIndex,
              value: item.text,
              fontName: item.fontName,
              heightMm: item.heightMm,
              bold: item.bold,
              italic: item.italic,
              underline: item.underline,
              color: item.color,
              widthFactor: item.widthFactor,
            },
            timeoutMilliseconds: 30000,
            profileId: "kompas-pages-executor",
          });
        });
      });
    }
  }
  return commands;
}

function createApi7AdditionalLineStyleCommands(handleId, cellMatrix) {
  const rows = Array.isArray(cellMatrix) ? cellMatrix.length : 0;
  const cols = Array.isArray(cellMatrix?.[0]) ? cellMatrix[0].length : 0;
  const commands = [];
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const row = Array.isArray(cellMatrix[rowIndex]) ? cellMatrix[rowIndex] : [];
    for (let columnIndex = 0; columnIndex < cols; columnIndex += 1) {
      const cell = row[columnIndex];
      if (!cell?.hasContent || !Array.isArray(cell.lines) || !cell.lines.length) {
        continue;
      }
      cell.lines.forEach((line, extraLineIndex) => {
        const lineIndex = extraLineIndex + 1;
        line.items.forEach((item, itemIndex) => {
          commands.push({
            commandId: "xlsx-to-kompas-tbl.table-cell-set-item",
            arguments: {
              handleId,
              rowIndex: cell.rowIndex,
              columnIndex: cell.columnIndex,
              lineIndex,
              itemIndex,
              value: item.text,
              fontName: item.fontName,
              heightMm: item.heightMm,
              bold: item.bold,
              italic: item.italic,
              underline: item.underline,
              color: item.color,
              widthFactor: item.widthFactor,
            },
            timeoutMilliseconds: 30000,
            profileId: "kompas-pages-executor",
          });
        });
      });
    }
  }
  return commands;
}

async function appendAdditionalLinesViaApi7(baseUrl, pairingToken, origin, tableHandleId, cellMatrix) {
  const commands = createApi7AdditionalLineCommands(tableHandleId, cellMatrix);
  if (!commands.length) {
    return;
  }
  await executeBatch(baseUrl, pairingToken, origin, commands, {
    reportVerbosity: "Compact",
    stopOnError: true,
  });
  const styleCommands = createApi7AdditionalLineStyleCommands(tableHandleId, cellMatrix);
  if (styleCommands.length) {
    await executeBatch(baseUrl, pairingToken, origin, styleCommands, {
      reportVerbosity: "Compact",
      stopOnError: true,
    });
  }
}

async function waitForResultBoxState(page, expectedText, options = {}) {
  const timeout = Math.max(1000, Number(options.timeout) || 15000);
  const progressPrefix = String(options.progressPrefix || "").trim();
  const allowedTexts = new Set(
    Array.isArray(options.allowedTexts)
      ? options.allowedTexts.map((value) => String(value || "").trim()).filter(Boolean)
      : [],
  );
  const resultBox = page.locator("#xlsx-result-box");
  const startedAt = Date.now();
  while ((Date.now() - startedAt) < timeout) {
    const text = String(await resultBox.textContent() || "").trim();
    if (text.includes(expectedText)) {
      return text;
    }
    if (text) {
      const isBusy = progressPrefix
        && (text === `${progressPrefix} is running...` || text.startsWith(`${progressPrefix}: `));
      if (!isBusy && !allowedTexts.has(text)) {
        throw new Error(`Expected result containing '${expectedText}', got '${text}'`);
      }
    }
    await page.waitForTimeout(500);
  }
  const actual = String(await resultBox.textContent() || "").trim();
  throw new Error(`Timed out waiting for '${expectedText}'. Current result box: '${actual}'`);
}

async function waitForFreshResultBoxState(page, previousText, expectedText, options = {}) {
  const timeout = Math.max(1000, Number(options.timeout) || 15000);
  const progressPrefix = String(options.progressPrefix || "").trim();
  const allowedTexts = new Set(
    Array.isArray(options.allowedTexts)
      ? options.allowedTexts.map((value) => String(value || "").trim()).filter(Boolean)
      : [],
  );
  const previous = String(previousText || "").trim();
  const resultBox = page.locator("#xlsx-result-box");
  const startedAt = Date.now();
  while ((Date.now() - startedAt) < timeout) {
    const text = String(await resultBox.textContent() || "").trim();
    if (text && text !== previous && text.includes(expectedText)) {
      return text;
    }
    if (text && text !== previous) {
      const isBusy = progressPrefix
        && (text === `${progressPrefix} is running...` || text.startsWith(`${progressPrefix}: `));
      if (!isBusy && !allowedTexts.has(text)) {
        throw new Error(`Expected fresh result containing '${expectedText}', got '${text}'`);
      }
    }
    await page.waitForTimeout(500);
  }
  const actual = String(await resultBox.textContent() || "").trim();
  throw new Error(`Timed out waiting for fresh '${expectedText}'. Current result box: '${actual}'`);
}

async function waitForInputValue(page, selector, expectedValue, timeoutMs = 60000) {
  const locator = page.locator(selector);
  const expected = String(expectedValue ?? "");
  const startedAt = Date.now();
  let lastValue = "";
  while ((Date.now() - startedAt) < timeoutMs) {
    lastValue = String(await locator.inputValue() || "");
    if (lastValue === expected) {
      return lastValue;
    }
    await page.waitForTimeout(250);
  }
  throw new Error(`Timed out waiting for ${selector}='${expected}'. Current value: '${lastValue}'`);
}

async function waitForTextContent(page, selector, expectedSubstring, timeoutMs = 60000) {
  const locator = page.locator(selector);
  const expected = String(expectedSubstring || "");
  const startedAt = Date.now();
  let lastValue = "";
  while ((Date.now() - startedAt) < timeoutMs) {
    lastValue = String(await locator.textContent() || "").trim();
    if (lastValue.includes(expected)) {
      return lastValue;
    }
    await page.waitForTimeout(250);
  }
  throw new Error(`Timed out waiting for ${selector} containing '${expected}'. Current value: '${lastValue}'`);
}

async function populateTableHandleDirect(baseUrl, pairingToken, origin, tableHandleId, cellMatrix) {
  const rows = Array.isArray(cellMatrix) ? cellMatrix.length : 0;
  const cols = Array.isArray(cellMatrix?.[0]) ? cellMatrix[0].length : 0;
  const writes = [];
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const row = Array.isArray(cellMatrix[rowIndex]) ? cellMatrix[rowIndex] : [];
    for (let columnIndex = 0; columnIndex < cols; columnIndex += 1) {
      if (row[columnIndex]?.hasContent) {
        writes.push(row[columnIndex]);
      }
    }
  }
  if (writes.length === 0) {
    return;
  }

  const api5DocumentHandle = requireHandleId(
    await executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-active-document2d", {}, 15000),
    "api5-active-document2d",
  );
  const tableReference = await readTableReference(baseUrl, pairingToken, origin, tableHandleId);
  expectApi5Success(
    (await executeCommand(
      baseUrl,
      pairingToken,
      origin,
      "kompas-pages-executor",
      "xlsx-to-kompas-tbl.api5-document-open-table",
      {
        handleId: api5DocumentHandle,
        tableReference,
      },
      30000,
    )).result,
    "api5-document-open-table",
    { tableReference },
  );

  try {
    for (const cell of writes) {
      const textParamHandle = requireHandleId(
        await executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-create-text-param", {}, 15000),
        "api5-text-param",
      );
      await executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-object-init", { handleId: textParamHandle }, 15000);
      const lineArrayHandle = requireHandleId(
        await executeCommand(
          baseUrl,
          pairingToken,
          origin,
          "kompas-pages-executor",
          "xlsx-to-kompas-tbl.api5-text-param-get-line-array",
          { handleId: textParamHandle },
          15000,
        ),
        "api5-line-array",
      );

      for (const line of cell.lines) {
        const lineHandle = requireHandleId(
          await executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-create-text-line-param", {}, 15000),
          "api5-text-line-param",
        );
        await executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-object-init", { handleId: lineHandle }, 15000);
        const itemArrayHandle = requireHandleId(
          await executeCommand(
            baseUrl,
            pairingToken,
            origin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.api5-text-line-param-get-item-array",
            { handleId: lineHandle },
            15000,
          ),
          "api5-item-array",
        );

        for (const item of line.items) {
          const itemHandle = requireHandleId(
            await executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-create-text-item-param", {}, 15000),
            "api5-text-item-param",
          );
          await executeCommand(baseUrl, pairingToken, origin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-object-init", { handleId: itemHandle }, 15000);
          await executeCommand(
            baseUrl,
            pairingToken,
            origin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.api5-text-item-param-set-basic",
            {
              handleId: itemHandle,
              value: item.text,
              itemType: API5_TEXT_ITEM_STRING,
            },
            15000,
          );
          const fontHandle = requireHandleId(
            await executeCommand(
              baseUrl,
              pairingToken,
              origin,
              "kompas-pages-executor",
              "xlsx-to-kompas-tbl.api5-text-item-param-get-font",
              { handleId: itemHandle },
              15000,
            ),
            "api5-text-item-font",
          );
          await executeCommand(
            baseUrl,
            pairingToken,
            origin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.api5-text-item-font-set",
            {
              handleId: fontHandle,
              fontName: item.fontName,
              heightMm: item.heightMm,
              color: item.color,
              bitVector: toApi5TextFlags(item),
            },
            15000,
          );
          expectApi5Success(
            (await executeCommand(
              baseUrl,
              pairingToken,
              origin,
              "kompas-pages-executor",
              "xlsx-to-kompas-tbl.api5-text-item-param-set-font",
              {
                handleId: itemHandle,
                fontHandle,
              },
              15000,
            )).result,
            "api5-text-item-param-set-font",
          );
          expectApi5Success(
            (await executeCommand(
              baseUrl,
              pairingToken,
              origin,
              "kompas-pages-executor",
              "xlsx-to-kompas-tbl.api5-dynamic-array-add-item",
              {
                handleId: itemArrayHandle,
                index: -1,
                itemHandle,
              },
              15000,
            )).result,
            "api5-item-array-add-item",
          );
        }
        expectApi5Success(
          (await executeCommand(
            baseUrl,
            pairingToken,
            origin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.api5-dynamic-array-add-item",
            {
              handleId: lineArrayHandle,
              index: -1,
              itemHandle: lineHandle,
            },
            15000,
            )).result,
            "api5-line-array-add-item",
          );
      }
      expectApi5Success(
        (await executeCommand(
          baseUrl,
          pairingToken,
          origin,
          "kompas-pages-executor",
          "xlsx-to-kompas-tbl.api5-document-set-table-cell-text",
          {
            handleId: api5DocumentHandle,
            cellNumber: toApi5CellNumber(cell.rowIndex, cell.columnIndex, cols),
            textHandle: textParamHandle,
          },
          30000,
        )).result,
        "api5-document-set-table-cell-text",
      );
    }
  } finally {
    expectApi5Success(
      (await executeCommand(
        baseUrl,
        pairingToken,
        origin,
        "kompas-pages-executor",
        "xlsx-to-kompas-tbl.api5-document-end-obj",
        { handleId: api5DocumentHandle },
        30000,
      )).result,
      "api5-document-end-obj",
    );
  }

  const refreshedTableHandleId = await reopenTableHandleByReference(baseUrl, pairingToken, origin, tableReference);
  const updateExecution = await executeCommand(
    baseUrl,
    pairingToken,
    origin,
    "kompas-pages-executor",
    "xlsx-to-kompas-tbl.table-update",
    { handleId: refreshedTableHandleId },
    30000,
  );

  const alignCommands = [];
  for (const cell of writes) {
    alignCommands.push({
      commandId: "xlsx-to-kompas-tbl.table-cell-set-one-line",
      arguments: {
        handleId: refreshedTableHandleId,
        rowIndex: cell.rowIndex,
        columnIndex: cell.columnIndex,
        oneLine: Boolean(cell.oneLine),
      },
      timeoutMilliseconds: 15000,
      profileId: "kompas-pages-executor",
    });
    for (let lineIndex = 0; lineIndex < cell.lines.length; lineIndex += 1) {
      if (lineIndex > 0) {
        continue;
      }
      alignCommands.push({
        commandId: "xlsx-to-kompas-tbl.table-cell-set-line-align",
        arguments: {
          handleId: refreshedTableHandleId,
          rowIndex: cell.rowIndex,
          columnIndex: cell.columnIndex,
          lineIndex,
          align: cell.alignCode,
        },
        timeoutMilliseconds: 15000,
        profileId: "kompas-pages-executor",
      });
    }
  }
  if (alignCommands.length) {
    await executeBatch(baseUrl, pairingToken, origin, alignCommands, {
      reportVerbosity: "Compact",
      stopOnError: true,
    });
  }
  return {
    handleId: refreshedTableHandleId,
    updateResult: updateExecution.result,
  };
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
  activeCommandStats = {};
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
  const kompasSampleCopyPath = path.join(
    workspaceRoot,
    `${path.parse(kompasSamplePath).name}-${stamp}${path.extname(kompasSamplePath)}`,
  );
  fs.copyFileSync(kompasSamplePath, kompasSampleCopyPath);
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
    kompasSampleTemplatePath: kompasSamplePath,
    kompasOpenedDocumentPath: kompasSampleCopyPath,
    outputTablePath,
  };

  try {
    let headedKompasProcess = null;
    if (args.headed) {
      headedKompasProcess = assertSingleKompasProcessForHeadedRun();
      report.headedKompasProcess = headedKompasProcess;
      report.headedKompasFocusBeforeStart = focusKompasWindow(headedKompasProcess.Id);
    }

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
      { path: kompasSampleCopyPath },
      60000,
    );
    const openedDocumentHandleId = String(openExecution.result?.handleId || "");
    if (!openedDocumentHandleId) {
      throw new Error(`Failed to open sample drawing: ${JSON.stringify(openExecution)}`);
    }
    await executeCommand(
      utilityUrl,
      pairingToken,
      pagesOrigin,
      "kompas-pages-executor",
      "xlsx-to-kompas-tbl.document-set-active",
      {
        handleId: openedDocumentHandleId,
        active: true,
      },
      15000,
    );
    report.openedDocumentHandleId = openedDocumentHandleId;
    report.openedDocumentHandleState = await readDocumentHandleState(
      utilityUrl,
      pairingToken,
      pagesOrigin,
      openedDocumentHandleId,
    );
    report.activeDocumentAfterOpen = await waitForActiveDocumentPath(
      utilityUrl,
      pairingToken,
      pagesOrigin,
      kompasSampleCopyPath,
      20000,
    );
    if (args.headed && report.headedKompasProcess?.Id) {
      report.headedKompasFocusAfterOpen = focusKompasWindow(report.headedKompasProcess.Id);
    }

    await page.locator("#module-badge").filter({ hasText: "doc ok" }).waitFor();
    await page.screenshot({ path: path.join(screenshotsRoot, "10-status.png"), fullPage: true });
    const scenarios = [];

    async function activateXlsxModule() {
      await page.locator('.rail__tab[data-module-id="xlsx-to-kompas-tbl"]').click();
      await page.locator('.module-host[data-module-id="xlsx-to-kompas-tbl"].is-active').waitFor();
    }

    async function resetXlsxModuleIfNeeded() {
      await activateXlsxModule();
      const matrixText = normalizeWhitespace(await page.locator("#xlsx-matrix-size").textContent());
      if (matrixText !== "0 x 0") {
        await page.locator("#xlsx-reset-button").click();
        await page.locator("#xlsx-matrix-size").filter({ hasText: "0 x 0" }).waitFor();
      }
    }

    async function loadFixtureViaUi(inputPath, matrixText, previewNeedle, screenshotName) {
      await activateXlsxModule();
      await page.locator("#xlsx-file-input").setInputFiles(inputPath);
      await page.locator("#xlsx-matrix-size").filter({ hasText: matrixText }).waitFor();
      if (previewNeedle) {
        await page.locator("#xlsx-preview-table").filter({ hasText: previewNeedle }).waitFor();
      }
      await page.screenshot({ path: path.join(screenshotsRoot, screenshotName), fullPage: true });
    }

    async function readTableCount() {
      const execution = await executeCommand(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        "kompas-pages-executor",
        "xlsx-to-kompas-tbl.view-table-count",
        { refresh: true },
        15000,
      );
      return Number(execution.result) || 0;
    }

    async function insertTableDirect(tablePath) {
      return executeCommand(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        "kompas-pages-executor",
        "xlsx-to-kompas-tbl.insert-table",
        { path: tablePath },
        60000,
      );
    }

    async function runWorkflowScenario() {
      await resetXlsxModuleIfNeeded();
      await loadFixtureViaUi(fixturePath, "8 x 13", "M2.2", "20-workflow-uploaded.png");
      const initialAutoFollowPath = await page.locator("#xlsx-output-path").inputValue();
      const expectedAutoFollowPath = path.join(
        path.dirname(String(report.activeDocumentAfterOpen?.path || "")),
        "table_M2.tbl",
      );
      assertCondition(
        initialAutoFollowPath && initialAutoFollowPath.toLowerCase().endsWith(".tbl"),
        "Expected auto-follow output path",
        { initialAutoFollowPath },
      );

      await page.locator("#xlsx-cell-width").fill("10");
      await page.locator("#xlsx-cell-height").fill("5");
      await waitForInputValue(page, "#xlsx-table-width", "130");
      await waitForInputValue(page, "#xlsx-table-height", "40");

      await page.locator("#xlsx-output-path").fill(outputTablePath);
      await page.locator("#xlsx-follow-button").click();
      await waitForInputValue(page, "#xlsx-output-path", expectedAutoFollowPath);
      await page.locator("#xlsx-output-path").fill(outputTablePath);

      const inlineTempArtifactsBefore = listInlineTempArtifacts();
      await page.locator("#xlsx-inline-button").click();
      const inlineResultText = await waitForResultBoxState(page, "Inline |", {
        timeout: 180000,
        progressPrefix: "Inline",
        allowedTexts: ["Matrix loaded. Export or Inline is ready."],
      });
      const inlineReportedHandleId = parseInlineHandleId(inlineResultText);
      const inlineTableReference = parseInlineTableReference(inlineResultText);
      const inlineHandleId = Number.isFinite(inlineTableReference) && inlineTableReference > 0
        ? await reopenTableHandleByReference(utilityUrl, pairingToken, pagesOrigin, inlineTableReference)
        : inlineReportedHandleId;
      const inlineTempArtifactsAfter = listInlineTempArtifacts();
      await page.screenshot({ path: path.join(screenshotsRoot, "25-workflow-inline.png"), fullPage: true });
      assertCondition(!fs.existsSync(outputTablePath), "Inline must not create the manual output path", { outputTablePath });
      assertCondition(inlineResultText.includes("Inline | no .tbl |"), "Inline must report no .tbl", { inlineResultText });
      assertCondition(Boolean(inlineReportedHandleId), "Inline did not report a direct handle", { inlineResultText });
      assertCondition(Number.isFinite(inlineTableReference) && inlineTableReference > 0, "Inline did not report a valid table reference", { inlineResultText });
      assertCondition(Boolean(inlineHandleId), "Inline direct table handle is not accessible by reference", {
        inlineReportedHandleId,
        inlineTableReference,
      });
      assertCondition(
        JSON.stringify(inlineTempArtifactsBefore) === JSON.stringify(inlineTempArtifactsAfter),
        "Inline created temp .tbl artifacts",
        { inlineTempArtifactsBefore, inlineTempArtifactsAfter },
      );

      if (args.pauseAfterInlineMs > 0) {
        if (args.headed && report.headedKompasProcess?.Id) {
          report.headedKompasFocusAfterInline = focusKompasWindow(report.headedKompasProcess.Id);
        }
        await new Promise((resolve) => setTimeout(resolve, args.pauseAfterInlineMs));
        if (!page.isClosed()) {
          await page.screenshot({ path: path.join(screenshotsRoot, "26-inline-paused.png"), fullPage: true });
        }
      }

      if (args.stopAfterInline) {
        return {
          initialAutoFollowPath,
          expectedAutoFollowPath,
          inlineHandleId,
          inlineReportedHandleId,
          inlineTableReference,
          inlineResultText,
          inlineTempArtifactsBefore,
          inlineTempArtifactsAfter,
          stopAfterInline: true,
          pauseAfterInlineMs: args.pauseAfterInlineMs,
        };
      }

      await page.locator("#xlsx-export-button").click();
      await page.locator("#xlsx-result-box").filter({ hasText: "OK" }).waitFor({ timeout: 180000 });
      await page.screenshot({ path: path.join(screenshotsRoot, "30-workflow-exported.png"), fullPage: true });

      assertCondition(fs.existsSync(outputTablePath), "Exported table was not created", { outputTablePath });
      const outputStat = fs.statSync(outputTablePath);
      assertCondition(outputStat.size > 0, "Exported table is empty", { outputTablePath, size: outputStat.size });

      await page.locator("#xlsx-insert-button").click();
      await page.locator("#xlsx-result-box").filter({ hasText: "Inserted" }).waitFor({ timeout: 60000 });
      await page.screenshot({ path: path.join(screenshotsRoot, "40-workflow-inserted.png"), fullPage: true });
      const insertResultText = String(await page.locator("#xlsx-result-box").textContent() || "");

      const [download] = await Promise.all([
        page.waitForEvent("download"),
        page.locator("#xlsx-download-button").click(),
      ]);
      const downloadedPath = path.join(workspaceRoot, await download.suggestedFilename());
      await download.saveAs(downloadedPath);
      assertCondition(fs.existsSync(downloadedPath), "Downloaded file was not created", { downloadedPath });
      assertCondition(fs.statSync(downloadedPath).size > 0, "Downloaded table artifact is empty", { downloadedPath });
      assertCondition(compareFiles(downloadedPath, outputTablePath), "Downloaded file must match exported file", { downloadedPath, outputTablePath });

      await page.locator('.rail__tab[data-module-id="kompas-text-sync"]').click();
      await page.locator('.module-host[data-module-id="kompas-text-sync"].is-active').waitFor();
      await activateXlsxModule();
      await page.locator("#xlsx-matrix-size").filter({ hasText: "8 x 13" }).waitFor();
      await page.screenshot({ path: path.join(screenshotsRoot, "50-workflow-tab-switch.png"), fullPage: true });

      await page.locator("#xlsx-reset-button").click();
      await page.locator("#xlsx-matrix-size").filter({ hasText: "0 x 0" }).waitFor();
      assertCondition(await page.locator("#xlsx-download-button").isDisabled(), "Download button should be disabled after reset");
      assertCondition(await page.locator("#xlsx-inline-button").isDisabled(), "Inline button should be disabled after reset");
      assertCondition(await page.locator("#xlsx-export-button").isDisabled(), "Export button should be disabled after reset");
      assertCondition(normalizeWhitespace(await page.locator("#xlsx-file-name").textContent()) === "not loaded", "File name should reset to not loaded");
      await page.screenshot({ path: path.join(screenshotsRoot, "60-workflow-reset.png"), fullPage: true });

      return {
        initialAutoFollowPath,
        expectedAutoFollowPath,
        inlineHandleId,
        inlineReportedHandleId,
        inlineTableReference,
        inlineResultText,
        inlineTempArtifactsBefore,
        inlineTempArtifactsAfter,
        outputTablePath,
        outputTableSize: outputStat.size,
        insertResultText,
        downloadedPath,
        downloadedSize: fs.statSync(downloadedPath).size,
      };
    }

    async function runRichProofScenario() {
      await resetXlsxModuleIfNeeded();
      const richLayout = buildRichProofLayout();
      const richFixturePath = createRichFixture(workspaceRoot);
      const richOutputPath = path.join(workspaceRoot, "rich-proof.tbl");
      const directRichPath = path.join(workspaceRoot, "rich-direct.tbl");
      const expectedRichSnapshots = richLayout.expectedSnapshots;
      const expectedRichCellMatrix = richLayout.cellMatrix;
      const richCells = flattenCellMatrix(expectedRichCellMatrix);
      const richRows = richLayout.rows;
      const richCols = richLayout.cols;
      const firstLineOnlyRichCellMatrix = createFirstLineOnlyCellMatrix(expectedRichCellMatrix);
      const additionalLinesRichCellMatrix = createAdditionalLinesOnlyCellMatrix(expectedRichCellMatrix);
      assertCondition(richLayout.cases.length >= 25, "Rich proof must contain at least 25 style cases", { count: richLayout.cases.length });

      let directHandleId = "";
      let directSavedHandleId = "";
      const directCreateExecution = await executeCommand(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        "kompas-pages-executor",
        "xlsx-to-kompas-tbl.create-table",
        {
          rows: richRows,
          cols: richCols,
          cellWidthMm: 20,
          cellHeightMm: 8,
        },
        60000,
      );
      directHandleId = String(directCreateExecution.result?.handleId || "");
      assertCondition(directHandleId, "Direct rich proof create-table did not return a handle");
      const directPopulateResult = await populateTableHandleDirect(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        directHandleId,
        firstLineOnlyRichCellMatrix,
      );
      directHandleId = directPopulateResult.handleId;
      await appendAdditionalLinesViaApi7(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        directHandleId,
        additionalLinesRichCellMatrix,
      );
      const directUpdateExecution = await executeCommand(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        "kompas-pages-executor",
        "xlsx-to-kompas-tbl.table-update",
        {
          handleId: directHandleId,
        },
        30000,
      );
      const directLiveSnapshots = await readSnapshotMap(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        directHandleId,
        richCells,
      );
      for (const cell of richCells) {
        assertCondition(
          JSON.stringify(directLiveSnapshots.normalized[cell.address]) === JSON.stringify(expectedRichSnapshots[cell.address]),
          `Direct in-memory rich ${cell.address} snapshot mismatch before Save`,
          {
            expected: expectedRichSnapshots[cell.address],
            actual: directLiveSnapshots.normalized[cell.address],
            raw: directLiveSnapshots.raw[cell.address],
          },
        );
      }
      const directSaveExecution = await executeCommand(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        "kompas-pages-executor",
        "xlsx-to-kompas-tbl.table-save",
        {
          handleId: directHandleId,
          path: directRichPath,
        },
        60000,
      );
      assertCondition(directSaveExecution.result !== false, "KOMPAS returned Save=false for direct rich proof handle");
      assertCondition(fs.existsSync(directRichPath), "Direct rich proof file missing after Save", { directRichPath });
      const directSavedInsert = await insertTableDirect(directRichPath);
      directSavedHandleId = String(directSavedInsert.result?.handleId || "");
      assertCondition(directSavedHandleId, "Direct rich proof saved insert did not return a handle");
      const directSavedSnapshots = await readSnapshotMap(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        directSavedHandleId,
        richCells,
      );

      await loadFixtureViaUi(richFixturePath, `${richRows} x ${richCols}`, "Wrap flag only", "70-rich-uploaded.png");
      await page.locator("#xlsx-output-path").fill(richOutputPath);

      const inlineCountBefore = await readTableCount();
      const inlineTempArtifactsBefore = listInlineTempArtifacts();
      await page.locator("#xlsx-inline-button").click();
      const inlineResultText = await waitForResultBoxState(page, "Inline |", {
        timeout: 180000,
        progressPrefix: "Inline",
        allowedTexts: ["Matrix loaded. Export or Inline is ready."],
      });
      const inlineReportedHandleId = parseInlineHandleId(inlineResultText);
      const inlineTableReference = parseInlineTableReference(inlineResultText);
      await page.screenshot({ path: path.join(screenshotsRoot, "75-rich-inline.png"), fullPage: true });
      const inlineCountAfter = await readTableCount();
      const inlineTempArtifactsAfter = listInlineTempArtifacts();
      assertCondition(inlineResultText.includes("Inline | no .tbl |"), "Rich inline must report no .tbl", { inlineResultText });
      assertCondition(Boolean(inlineReportedHandleId), "Rich inline did not report a handle", { inlineResultText });
      assertCondition(Number.isFinite(inlineTableReference) && inlineTableReference > 0, "Rich inline did not report a valid table reference", { inlineResultText });
      assertCondition(
        JSON.stringify(inlineTempArtifactsBefore) === JSON.stringify(inlineTempArtifactsAfter),
        "Rich inline created temp .tbl artifacts",
        { inlineTempArtifactsBefore, inlineTempArtifactsAfter },
      );

      await page.locator("#xlsx-export-button").click();
      await page.locator("#xlsx-result-box").filter({ hasText: "OK" }).waitFor({ timeout: 180000 });
      await page.screenshot({ path: path.join(screenshotsRoot, "80-rich-exported.png"), fullPage: true });
      assertCondition(fs.existsSync(richOutputPath), "Rich export file missing", { richOutputPath });
      assertCondition(fs.statSync(richOutputPath).size > 0, "Rich export file is empty", { richOutputPath });

      const exportedInsert = await insertTableDirect(richOutputPath);
      const exportedHandleId = String(exportedInsert.result?.handleId || "");
      assertCondition(exportedHandleId, "Insert of exported rich table did not return handle");

      const inlineHandleId = await reopenTableHandleByReference(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        inlineTableReference,
      );
      assertCondition(inlineHandleId, "Reopen of inline rich table by reference did not return handle");

      const exportedSnapshots = await readSnapshotMap(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        exportedHandleId,
        richCells,
      );
      const inlineSnapshots = await readSnapshotMap(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        inlineHandleId,
        richCells,
      );

      for (const cell of richCells) {
        assertCondition(
          JSON.stringify(directSavedSnapshots.normalized[cell.address]) === JSON.stringify(expectedRichSnapshots[cell.address]),
          `Direct saved rich ${cell.address} snapshot mismatch after Save/Insert`,
          {
            expected: expectedRichSnapshots[cell.address],
            actual: directSavedSnapshots.normalized[cell.address],
            raw: directSavedSnapshots.raw[cell.address],
          },
        );
        assertCondition(
          JSON.stringify(exportedSnapshots.normalized[cell.address]) === JSON.stringify(expectedRichSnapshots[cell.address]),
          `Exported rich ${cell.address} snapshot mismatch`,
          {
            expected: expectedRichSnapshots[cell.address],
            actual: exportedSnapshots.normalized[cell.address],
            raw: exportedSnapshots.raw[cell.address],
          },
        );
        assertCondition(
          JSON.stringify(exportedSnapshots.normalized[cell.address]) === JSON.stringify(inlineSnapshots.normalized[cell.address]),
          `Inline rich ${cell.address} snapshot differs from export artifact`,
          {
            exported: exportedSnapshots.raw[cell.address],
            inline: inlineSnapshots.raw[cell.address],
            normalizedExported: exportedSnapshots.normalized[cell.address],
            normalizedInline: inlineSnapshots.normalized[cell.address],
          },
        );
      }

      return {
        richFixturePath,
        styleCaseCount: richLayout.cases.length,
        styleCases: richLayout.cases.map((styleCase) => ({
          label: styleCase.label,
          kind: styleCase.kind,
          address: styleCase.address,
          horizontal: styleCase.horizontal,
          wrapText: styleCase.wrapText,
          oneLine: styleCase.oneLine,
          lineCount: styleCase.lineCount,
          itemCount: styleCase.itemCount,
          text: styleCase.text,
        })),
        expectedAddresses: richCells.map((cell) => cell.address),
        directPopulateUpdateResult: directPopulateResult.updateResult,
        directFinalUpdateResult: directUpdateExecution.result,
        inlineReportedHandleId,
        inlineTableReference,
        inlineResultText,
        inlineTempArtifactsBefore,
        inlineTempArtifactsAfter,
        inlineCountBefore,
        inlineCountAfter,
        richOutputPath,
        richOutputSize: fs.statSync(richOutputPath).size,
        exportedHandleId,
        inlineHandleId,
        directRichPath,
        directRichSize: fs.existsSync(directRichPath) ? fs.statSync(directRichPath).size : 0,
        directHandleId,
        directSavedHandleId,
        directLiveSnapshots: directLiveSnapshots.raw,
        directSavedSnapshots: directSavedSnapshots.raw,
        exportedSnapshots: exportedSnapshots.raw,
        inlineSnapshots: inlineSnapshots.raw,
        normalizedDirectLiveSnapshots: directLiveSnapshots.normalized,
        normalizedDirectSavedSnapshots: directSavedSnapshots.normalized,
        normalizedExportedSnapshots: exportedSnapshots.normalized,
        normalizedInlineSnapshots: inlineSnapshots.normalized,
      };
    }

    async function runAutoFitProofScenario() {
      await resetXlsxModuleIfNeeded();
      const richLayout = buildRichProofLayout();
      const richFixturePath = createRichFixture(workspaceRoot);
      const richCells = flattenCellMatrix(richLayout.cellMatrix);
      const richRows = richLayout.rows;
      const richCols = richLayout.cols;
      const expectedComparableSnapshots = {};
      for (const cell of richCells) {
        expectedComparableSnapshots[cell.address] = stripHeightMmFromNormalizedSnapshot(richLayout.expectedSnapshots[cell.address]);
      }

      await loadFixtureViaUi(richFixturePath, `${richRows} x ${richCols}`, "Wrap flag only", "81-autofit-loaded.png");

      async function configureAutoFitPhase(cellWidthMm, cellHeightMm, autoFitEnabled, screenshotName) {
        await page.locator("#xlsx-cell-width").fill(String(cellWidthMm));
        await page.locator("#xlsx-cell-height").fill(String(cellHeightMm));
        await waitForInputValue(page, "#xlsx-cell-width", String(cellWidthMm));
        await waitForInputValue(page, "#xlsx-cell-height", String(cellHeightMm));
        await waitForInputValue(page, "#xlsx-table-width", String(cellWidthMm * richCols));
        await waitForInputValue(page, "#xlsx-table-height", String(cellHeightMm * richRows));
        await page.locator("#xlsx-font-autofit").setChecked(autoFitEnabled);
        const summaryText = await waitForTextContent(
          page,
          "#xlsx-layout-summary",
          autoFitEnabled ? "text=auto-fit" : "text=excel 1:1",
          60000,
        );
        await page.screenshot({ path: path.join(screenshotsRoot, screenshotName), fullPage: true });
        return summaryText;
      }

      async function inlineAndReadSnapshots(screenshotName) {
        const previousResultText = String(await page.locator("#xlsx-result-box").textContent() || "").trim();
        const inlineTempArtifactsBefore = listInlineTempArtifacts();
        await page.locator("#xlsx-inline-button").click();
        const inlineResultText = await waitForFreshResultBoxState(page, previousResultText, "Inline |", {
          timeout: 180000,
          progressPrefix: "Inline",
          allowedTexts: ["Matrix loaded. Export or Inline is ready."],
        });
        const inlineReportedHandleId = parseInlineHandleId(inlineResultText);
        const inlineTableReference = parseInlineTableReference(inlineResultText);
        await page.screenshot({ path: path.join(screenshotsRoot, screenshotName), fullPage: true });
        const inlineTempArtifactsAfter = listInlineTempArtifacts();
        assertCondition(inlineResultText.includes("Inline | no .tbl |"), "Auto-fit inline must report no .tbl", { inlineResultText });
        assertCondition(Boolean(inlineReportedHandleId), "Auto-fit inline did not report a handle", { inlineResultText });
        assertCondition(Number.isFinite(inlineTableReference) && inlineTableReference > 0, "Auto-fit inline did not report a valid table reference", { inlineResultText });
        assertCondition(
          JSON.stringify(inlineTempArtifactsBefore) === JSON.stringify(inlineTempArtifactsAfter),
          "Auto-fit inline created temp .tbl artifacts",
          { inlineTempArtifactsBefore, inlineTempArtifactsAfter },
        );
        const handleId = await reopenTableHandleByReference(
          utilityUrl,
          pairingToken,
          pagesOrigin,
          inlineTableReference,
        );
        assertCondition(handleId, "Auto-fit inline reopen by reference did not return handle", { inlineTableReference });
        const snapshots = await readSnapshotMap(
          utilityUrl,
          pairingToken,
          pagesOrigin,
          handleId,
          richCells,
        );
        return {
          handleId,
          inlineReportedHandleId,
          inlineTableReference,
          inlineResultText,
          inlineTempArtifactsBefore,
          inlineTempArtifactsAfter,
          snapshots,
        };
      }

      const baselineSummaryText = await configureAutoFitPhase(20, 8, false, "82-autofit-baseline-ready.png");
      const baselinePhase = await inlineAndReadSnapshots("83-autofit-baseline-inline.png");
      const shrinkSummaryText = await configureAutoFitPhase(12, 4, true, "84-autofit-shrink-ready.png");
      const shrinkPhase = await inlineAndReadSnapshots("85-autofit-shrink-inline.png");
      const growSummaryText = await configureAutoFitPhase(45, 14, true, "86-autofit-grow-ready.png");
      const growPhase = await inlineAndReadSnapshots("87-autofit-grow-inline.png");

      const baselineComparableSnapshots = {};
      const shrinkComparableSnapshots = {};
      const growComparableSnapshots = {};
      for (const cell of richCells) {
        baselineComparableSnapshots[cell.address] = stripHeightMmFromNormalizedSnapshot(baselinePhase.snapshots.normalized[cell.address]);
        shrinkComparableSnapshots[cell.address] = stripHeightMmFromNormalizedSnapshot(shrinkPhase.snapshots.normalized[cell.address]);
        growComparableSnapshots[cell.address] = stripHeightMmFromNormalizedSnapshot(growPhase.snapshots.normalized[cell.address]);
      }

      for (const cell of richCells) {
        assertCondition(
          JSON.stringify(baselineComparableSnapshots[cell.address]) === JSON.stringify(expectedComparableSnapshots[cell.address]),
          `Auto-fit baseline ${cell.address} changed text/style content`,
          {
            expected: expectedComparableSnapshots[cell.address],
            actual: baselineComparableSnapshots[cell.address],
          },
        );
        assertCondition(
          JSON.stringify(shrinkComparableSnapshots[cell.address]) === JSON.stringify(expectedComparableSnapshots[cell.address]),
          `Auto-fit shrink ${cell.address} changed text/style content`,
          {
            expected: expectedComparableSnapshots[cell.address],
            actual: shrinkComparableSnapshots[cell.address],
          },
        );
        assertCondition(
          JSON.stringify(growComparableSnapshots[cell.address]) === JSON.stringify(expectedComparableSnapshots[cell.address]),
          `Auto-fit grow ${cell.address} changed text/style content`,
          {
            expected: expectedComparableSnapshots[cell.address],
            actual: growComparableSnapshots[cell.address],
          },
        );
      }

      const baselineHeightMetrics = buildSnapshotHeightMetrics(baselinePhase.snapshots.normalized);
      const shrinkHeightMetrics = buildSnapshotHeightMetrics(shrinkPhase.snapshots.normalized);
      const growHeightMetrics = buildSnapshotHeightMetrics(growPhase.snapshots.normalized);
      const heightDeltaEpsilon = 0.05;
      const trackedAddresses = richCells.map((cell) => cell.address);
      const provenShrinkAddresses = trackedAddresses.filter((address) => (
        Number(shrinkHeightMetrics[address]?.maxItemHeightMm || 0) < (Number(baselineHeightMetrics[address]?.maxItemHeightMm || 0) - heightDeltaEpsilon)
        || Number(shrinkHeightMetrics[address]?.totalLineHeightMm || 0) < (Number(baselineHeightMetrics[address]?.totalLineHeightMm || 0) - heightDeltaEpsilon)
      ));
      const provenGrowAddresses = trackedAddresses.filter((address) => (
        Number(growHeightMetrics[address]?.maxItemHeightMm || 0) > (Number(baselineHeightMetrics[address]?.maxItemHeightMm || 0) + heightDeltaEpsilon)
        || Number(growHeightMetrics[address]?.totalLineHeightMm || 0) > (Number(baselineHeightMetrics[address]?.totalLineHeightMm || 0) + heightDeltaEpsilon)
      ));

      assertCondition(parseAdjustedCellCount(shrinkSummaryText) > 0, "Auto-fit shrink summary did not report adjusted cells", { shrinkSummaryText });
      assertCondition(parseAdjustedCellCount(growSummaryText) > 0, "Auto-fit grow summary did not report adjusted cells", { growSummaryText });
      assertCondition(provenShrinkAddresses.includes("A1"), "Auto-fit shrink did not reduce A1 height in KOMPAS readback", {
        baseline: baselineHeightMetrics.A1,
        shrink: shrinkHeightMetrics.A1,
      });
      assertCondition(provenGrowAddresses.includes("E1"), "Auto-fit grow did not enlarge E1 height in KOMPAS readback", {
        baseline: baselineHeightMetrics.E1,
        grow: growHeightMetrics.E1,
      });
      assertCondition(provenShrinkAddresses.length >= 5, "Auto-fit shrink proof is too weak", { provenShrinkAddresses });
      assertCondition(provenGrowAddresses.length >= 5, "Auto-fit grow proof is too weak", { provenGrowAddresses });

      return {
        richFixturePath,
        trackedAddresses,
        baselineSummaryText,
        shrinkSummaryText,
        growSummaryText,
        baselineAdjustedCellCount: parseAdjustedCellCount(baselineSummaryText),
        shrinkAdjustedCellCount: parseAdjustedCellCount(shrinkSummaryText),
        growAdjustedCellCount: parseAdjustedCellCount(growSummaryText),
        baselineCellSize: { widthMm: 20, heightMm: 8, autoFitEnabled: false },
        shrinkCellSize: { widthMm: 12, heightMm: 4, autoFitEnabled: true },
        growCellSize: { widthMm: 45, heightMm: 14, autoFitEnabled: true },
        baselineHandleId: baselinePhase.handleId,
        baselineInlineReportedHandleId: baselinePhase.inlineReportedHandleId,
        baselineInlineTableReference: baselinePhase.inlineTableReference,
        baselineInlineResultText: baselinePhase.inlineResultText,
        shrinkHandleId: shrinkPhase.handleId,
        shrinkInlineReportedHandleId: shrinkPhase.inlineReportedHandleId,
        shrinkInlineTableReference: shrinkPhase.inlineTableReference,
        shrinkInlineResultText: shrinkPhase.inlineResultText,
        growHandleId: growPhase.handleId,
        growInlineReportedHandleId: growPhase.inlineReportedHandleId,
        growInlineTableReference: growPhase.inlineTableReference,
        growInlineResultText: growPhase.inlineResultText,
        normalizedBaselineSnapshots: baselinePhase.snapshots.normalized,
        normalizedShrinkSnapshots: shrinkPhase.snapshots.normalized,
        normalizedGrowSnapshots: growPhase.snapshots.normalized,
        baselineHeightMetrics,
        shrinkHeightMetrics,
        growHeightMetrics,
        provenShrinkAddresses,
        provenGrowAddresses,
      };
    }

    async function runCommandProofScenario() {
      const iterationCount = 10;
      const rows = 2;
      const cols = 2;
      const proofRoot = path.join(workspaceRoot, "command-proof");
      ensureDir(proofRoot);

      const sampleIterations = [];
      for (let iterationIndex = 0; iterationIndex < iterationCount; iterationIndex += 1) {
        const iteration = iterationIndex + 1;
        const commandProofDocumentPath = path.join(
          workspaceRoot,
          `Plug-command-proof-${stamp}-${String(iteration).padStart(2, "0")}.frw`,
        );
        fs.copyFileSync(kompasSamplePath, commandProofDocumentPath);
        const trackedDocumentHandleId = requireHandleId(
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.open-document",
            { path: commandProofDocumentPath },
            60000,
          ),
          "command-proof.open-document",
        );
        assertCondition(trackedDocumentHandleId, "Command proof requires an opened document handle", {
          iteration,
          commandProofDocumentPath,
        });
        await executeCommand(
          utilityUrl,
          pairingToken,
          pagesOrigin,
          "kompas-pages-executor",
          "xlsx-to-kompas-tbl.document-set-active",
          { handleId: trackedDocumentHandleId, active: true },
          15000,
        );
        await waitForActiveDocumentPath(
          utilityUrl,
          pairingToken,
          pagesOrigin,
          commandProofDocumentPath,
          30000,
        );
        let tableHandleId = "";
        let reopenedHandleId = "";
        let insertedHandleId = "";
        let api5ProofHandleId = "";
        let api5DocumentHandle = "";
        let api5OpenedTable = false;
        try {
          const applicationInfo = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.application.info",
            { refresh: true },
            15000,
          );
          assertCondition(Boolean(applicationInfo.result), "Command proof application.info returned empty payload", { iteration });

          const activeDocumentInfo = await readActiveDocumentInfo(utilityUrl, pairingToken, pagesOrigin);
          assertCondition(
            normalizeWindowsPathForComparison(activeDocumentInfo?.path) === normalizeWindowsPathForComparison(commandProofDocumentPath),
            "Command proof active document path mismatch",
            { iteration, actual: activeDocumentInfo?.path, expected: commandProofDocumentPath },
          );

          const activeViewInfo = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.active-view",
            { refresh: true },
            15000,
          );
          assertCondition(Boolean(activeViewInfo.result?.handleId), "Command proof active view handle is missing", { iteration, activeViewInfo });

          const [viewXExecution, viewYExecution, frameXExecution, frameYExecution] = await Promise.all([
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.active-view-x", {}, 15000),
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.active-view-y", {}, 15000),
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.active-frame-center-x", {}, 15000),
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.active-frame-center-y", {}, 15000),
          ]);
          const frameCenterX = Number(frameXExecution.result);
          const frameCenterY = Number(frameYExecution.result);
          assertCondition(Number.isFinite(frameCenterX) && Number.isFinite(frameCenterY), "Command proof frame center is invalid", {
            iteration,
            frameCenterX: frameXExecution.result,
            frameCenterY: frameYExecution.result,
          });

          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.document-set-active",
            {
              handleId: trackedDocumentHandleId,
              active: true,
            },
            15000,
          );
          const documentState = await readDocumentHandleState(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            trackedDocumentHandleId,
          );
          assertCondition(documentState.active === true, "Command proof tracked document is not active", { iteration, documentState });
          assertCondition(documentState.visible === true, "Command proof tracked document is not visible", { iteration, documentState });

          const tableCountBefore = Number((await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.view-table-count",
            { refresh: true },
            15000,
          )).result) || 0;

          tableHandleId = requireHandleId(
            await executeCommand(
              utilityUrl,
              pairingToken,
              pagesOrigin,
              "kompas-pages-executor",
              "xlsx-to-kompas-tbl.create-table",
              {
                rows,
                cols,
                cellWidthMm: 18,
                cellHeightMm: 7,
              },
              60000,
            ),
            "command-proof.create-table",
          );

          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-set-text",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
              value: `seed-${iteration}`,
            },
            15000,
          );
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-clear-text",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
            },
            15000,
          );
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-set-text",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
              value: `body-${iteration}`,
            },
            15000,
          );
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-set-one-line",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
              oneLine: false,
            },
            15000,
          );
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-set-line-align",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
              lineIndex: 0,
              align: iterationIndex % 3,
            },
            15000,
          );
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-add-line",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
              align: (iterationIndex + 1) % 3,
            },
            30000,
          );
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-add-item-before",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
              lineIndex: 0,
              itemIndex: 0,
              value: `P${iteration}`,
              fontName: "Tahoma",
              heightMm: Number((2.8 + (iterationIndex * 0.1)).toFixed(4)),
              bold: true,
              italic: false,
              underline: false,
              color: 0x0000FF,
              widthFactor: 1,
            },
            30000,
          );
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-add-item",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
              lineIndex: 0,
              itemIndex: 1,
              value: `S${iteration}`,
              fontName: "Calibri",
              heightMm: Number((2.6 + (iterationIndex * 0.1)).toFixed(4)),
              bold: false,
              italic: Boolean(iterationIndex % 2),
              underline: false,
              color: 0xAA5500,
              widthFactor: 1,
            },
            30000,
          );
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-add-item",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
              lineIndex: 1,
              itemIndex: 0,
              value: `L${iteration}`,
              fontName: "Arial",
              heightMm: Number((2.5 + (iterationIndex * 0.1)).toFixed(4)),
              bold: false,
              italic: false,
              underline: true,
              color: 0x008800,
              widthFactor: 1,
            },
            30000,
          );
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-set-item",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
              lineIndex: 0,
              itemIndex: 0,
              value: `PX${iteration}`,
              fontName: "Verdana",
              heightMm: Number((3 + (iterationIndex * 0.1)).toFixed(4)),
              bold: Boolean(iterationIndex % 2 === 0),
              italic: Boolean(iterationIndex % 2 === 1),
              underline: true,
              color: 0x006600,
              widthFactor: 1,
            },
            30000,
          );

          const cellTextExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-get-text",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
            },
            15000,
          );
          const oneLineExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-get-one-line",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
            },
            15000,
          );
          const lineCountExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-get-line-count",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
            },
            15000,
          );
          const lineAlignExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-get-line-align",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
              lineIndex: 0,
            },
            15000,
          );
          const itemCountExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-get-line-item-count",
            {
              handleId: tableHandleId,
              rowIndex: 0,
              columnIndex: 0,
              lineIndex: 0,
            },
            15000,
          );
          const [
            itemTextExecution,
            itemFontExecution,
            itemHeightExecution,
            itemBoldExecution,
            itemItalicExecution,
            itemUnderlineExecution,
            itemColorExecution,
            itemWidthFactorExecution,
          ] = await Promise.all([
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-text", { handleId: tableHandleId, rowIndex: 0, columnIndex: 0, lineIndex: 0, itemIndex: 0 }, 15000),
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-font-name", { handleId: tableHandleId, rowIndex: 0, columnIndex: 0, lineIndex: 0, itemIndex: 0 }, 15000),
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-height", { handleId: tableHandleId, rowIndex: 0, columnIndex: 0, lineIndex: 0, itemIndex: 0 }, 15000),
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-bold", { handleId: tableHandleId, rowIndex: 0, columnIndex: 0, lineIndex: 0, itemIndex: 0 }, 15000),
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-italic", { handleId: tableHandleId, rowIndex: 0, columnIndex: 0, lineIndex: 0, itemIndex: 0 }, 15000),
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-underline", { handleId: tableHandleId, rowIndex: 0, columnIndex: 0, lineIndex: 0, itemIndex: 0 }, 15000),
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-color", { handleId: tableHandleId, rowIndex: 0, columnIndex: 0, lineIndex: 0, itemIndex: 0 }, 15000),
            executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-cell-get-item-width-factor", { handleId: tableHandleId, rowIndex: 0, columnIndex: 0, lineIndex: 0, itemIndex: 0 }, 15000),
          ]);

          assertCondition(String(cellTextExecution.result || "").includes(`body-${iteration}`) && String(cellTextExecution.result || "").includes(`L${iteration}`), "Command proof cell text mismatch after API7 edit", {
            iteration,
            cellText: cellTextExecution.result,
          });
          assertCondition(coerceBooleanLike(oneLineExecution.result) === false, "Command proof cell should be multiline", { iteration, oneLine: oneLineExecution.result });
          assertCondition((Number(lineCountExecution.result) || 0) >= 2, "Command proof line count is too small", { iteration, lineCount: lineCountExecution.result });
          assertCondition((Number(itemCountExecution.result) || 0) >= 2, "Command proof line item count is too small", { iteration, itemCount: itemCountExecution.result });
          assertCondition(String(itemTextExecution.result || "").includes(`PX${iteration}`), "Command proof item text mismatch", { iteration, itemText: itemTextExecution.result });
          assertCondition(String(itemFontExecution.result || "").length > 0, "Command proof item font is empty", { iteration });
          assertCondition((Number(itemHeightExecution.result) || 0) > 0, "Command proof item height is invalid", { iteration, height: itemHeightExecution.result });
          assertCondition(Number(itemWidthFactorExecution.result) === 1, "Command proof item width factor mismatch", { iteration, widthFactor: itemWidthFactorExecution.result });

          const tableUpdateExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-update",
            { handleId: tableHandleId },
            30000,
          );
          assertCondition(tableUpdateExecution.result !== false, "Command proof table-update failed before reference read", { iteration });

          const referencePrepareExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-update",
            { handleId: tableHandleId },
            30000,
          );
          const referenceExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-get-reference",
            { handleId: tableHandleId },
            15000,
          );
          const tableReference = coerceNumberLike(referenceExecution.result);
          assertCondition(Number.isFinite(tableReference) && tableReference > 0, "Command proof table reference is invalid", {
            iteration,
            referencePrepareResult: referencePrepareExecution.result,
            referenceResult: referenceExecution.result,
          });
          const tempExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-get-temp",
            { handleId: tableHandleId },
            15000,
          );
          const validExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-get-valid",
            { handleId: tableHandleId },
            15000,
          );
          reopenedHandleId = await reopenTableHandleByReference(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            tableReference,
          );
          const targetX = Number((frameCenterX + (iteration * 1.5)).toFixed(4));
          const targetY = Number((frameCenterY + (iteration * 1.25)).toFixed(4));
          const positionExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-set-position",
            {
              handleId: reopenedHandleId,
              x: targetX,
              y: targetY,
            },
            30000,
          );
          assertCondition(positionExecution.result !== false, "Command proof table-set-position failed", { iteration, targetX, targetY });
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.active-frame-refresh",
            { refresh: true },
            30000,
          );
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.active-view-update",
            { refresh: true },
            30000,
          );
          const tableXExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-get-x",
            { handleId: reopenedHandleId },
            15000,
          );
          const tableYExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-get-y",
            { handleId: reopenedHandleId },
            15000,
          );
          assertCondition(
            Math.abs((Number(tableXExecution.result) || 0) - targetX) < 0.01 && Math.abs((Number(tableYExecution.result) || 0) - targetY) < 0.01,
            "Command proof table coordinates mismatch after move",
            {
              iteration,
              targetX,
              targetY,
              actualX: tableXExecution.result,
              actualY: tableYExecution.result,
            },
          );

          const savePath = path.join(proofRoot, `command-proof-${String(iteration).padStart(2, "0")}.tbl`);
          const saveExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-save",
            {
              handleId: reopenedHandleId,
              path: savePath,
            },
            60000,
          );
          assertCondition(saveExecution.result !== false, "Command proof table-save failed", { iteration, savePath });
          assertCondition(fs.existsSync(savePath) && fs.statSync(savePath).size > 0, "Command proof saved .tbl is missing or empty", { iteration, savePath });

          insertedHandleId = requireHandleId(
            await executeCommand(
              utilityUrl,
              pairingToken,
              pagesOrigin,
              "kompas-pages-executor",
              "xlsx-to-kompas-tbl.insert-table",
              { path: savePath },
              60000,
            ),
            "command-proof.insert-table",
          );
          const insertedValidExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-get-valid",
            { handleId: insertedHandleId },
            15000,
          );
          const insertedTempExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-get-temp",
            { handleId: insertedHandleId },
            15000,
          );
          assertCondition(coerceBooleanLike(insertedValidExecution.result) === true, "Command proof inserted table is invalid", { iteration });
          assertCondition(coerceBooleanLike(insertedTempExecution.result) === false, "Command proof inserted table must not be temporary", { iteration });

          const dummyTextParamHandle = requireHandleId(
            await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-create-text-param", {}, 15000),
            "command-proof.api5-dummy-text-param",
          );
          await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-object-init", { handleId: dummyTextParamHandle }, 15000);
          const dummyLineArrayFromText = requireHandleId(
            await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-text-param-get-line-array", { handleId: dummyTextParamHandle }, 15000),
            "command-proof.api5-dummy-line-array",
          );
          expectApi5Success(
            (await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-text-param-set-line-array", { handleId: dummyTextParamHandle, lineArrayHandle: dummyLineArrayFromText }, 15000)).result,
            "command-proof.api5-text-param-set-line-array",
          );
          const dummyLineHandle = requireHandleId(
            await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-create-text-line-param", {}, 15000),
            "command-proof.api5-dummy-line-param",
          );
          await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-object-init", { handleId: dummyLineHandle }, 15000);
          const dummyItemArrayFromLine = requireHandleId(
            await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-text-line-param-get-item-array", { handleId: dummyLineHandle }, 15000),
            "command-proof.api5-dummy-item-array",
          );
          expectApi5Success(
            (await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-text-line-param-set-item-array", { handleId: dummyLineHandle, itemArrayHandle: dummyItemArrayFromLine }, 15000)).result,
            "command-proof.api5-text-line-param-set-item-array",
          );
          const dummyItemHandle = requireHandleId(
            await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-create-text-item-param", {}, 15000),
            "command-proof.api5-dummy-item-param",
          );
          await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-object-init", { handleId: dummyItemHandle }, 15000);
          await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-text-item-param-set-basic", { handleId: dummyItemHandle, value: `dummy-${iteration}`, itemType: API5_TEXT_ITEM_STRING }, 15000);
          const dummyFontHandle = requireHandleId(
            await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-text-item-param-get-font", { handleId: dummyItemHandle }, 15000),
            "command-proof.api5-dummy-font",
          );
          await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-text-item-font-set", { handleId: dummyFontHandle, fontName: "Arial", heightMm: Number((2.4 + (iterationIndex * 0.1)).toFixed(4)), color: 0xFF0000, bitVector: toApi5TextFlags({ bold: true, italic: false, underline: true }) }, 15000);
          expectApi5Success(
            (await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-text-item-param-set-font", { handleId: dummyItemHandle, fontHandle: dummyFontHandle }, 15000)).result,
            "command-proof.api5-text-item-param-set-font",
          );
          const dummyDynamicLineArray = requireHandleId(
            await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-create-dynamic-array", { arrayType: API5_TEXT_LINE_ARRAY_TYPE }, 15000),
            "command-proof.api5-dynamic-line-array",
          );
          const dummyDynamicItemArray = requireHandleId(
            await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-create-dynamic-array", { arrayType: API5_TEXT_ITEM_ARRAY_TYPE }, 15000),
            "command-proof.api5-dynamic-item-array",
          );
          expectApi5Success(
            (await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-dynamic-array-add-item", { handleId: dummyDynamicItemArray, index: -1, itemHandle: dummyItemHandle }, 15000)).result,
            "command-proof.api5-dynamic-array-add-item.item",
          );
          assertCondition((Number((await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-dynamic-array-count", { handleId: dummyDynamicItemArray }, 15000)).result) || 0) >= 1, "Command proof API5 item dynamic array count is invalid", { iteration });
          expectApi5Success(
            (await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-dynamic-array-add-item", { handleId: dummyDynamicLineArray, index: -1, itemHandle: dummyLineHandle }, 15000)).result,
            "command-proof.api5-dynamic-array-add-item.line",
          );
          assertCondition((Number((await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-dynamic-array-count", { handleId: dummyDynamicLineArray }, 15000)).result) || 0) >= 1, "Command proof API5 line dynamic array count is invalid", { iteration });

          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.active-document",
            { refresh: true },
            15000,
          );
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.active-view",
            { refresh: true },
            15000,
          );

          api5ProofHandleId = requireHandleId(
            await executeCommand(
              utilityUrl,
              pairingToken,
              pagesOrigin,
              "kompas-pages-executor",
              "xlsx-to-kompas-tbl.create-table",
              { rows: 1, cols: 1, cellWidthMm: 14, cellHeightMm: 6 },
              60000,
            ),
            "command-proof.api5-proof-table",
          );
          const api5ProofMatrix = [[{
            address: "A1",
            rowIndex: 0,
            columnIndex: 0,
            text: `api5-${iteration}`,
            oneLine: true,
            alignCode: 0,
            hasContent: true,
            lines: [{
              items: [{
                text: `api5-${iteration}`,
                fontName: "Arial",
                heightMm: Number((2.8 + (iterationIndex * 0.1)).toFixed(4)),
                bold: Boolean(iterationIndex % 2 === 0),
                italic: Boolean(iterationIndex % 2 === 1),
                underline: true,
                color: 0xFF0000,
                widthFactor: 1,
              }],
            }],
          }]];
          const api5PopulateResult = await populateTableHandleDirect(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            api5ProofHandleId,
            api5ProofMatrix,
          );
          api5ProofHandleId = api5PopulateResult.handleId;
          const api5ProofTextExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-get-text",
            { handleId: api5ProofHandleId, rowIndex: 0, columnIndex: 0 },
            15000,
          );
          assertCondition(String(api5ProofTextExecution.result || "").includes(`api5-${iteration}`), "Command proof API5 helper write did not reach readback", {
            iteration,
            api5ProofText: api5ProofTextExecution.result,
          });

          const api5ProofReference = await readTableReference(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            api5ProofHandleId,
          );
          api5DocumentHandle = requireHandleId(
            await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-active-document2d", {}, 15000),
            "command-proof.api5-active-document2d",
          );
          expectApi5Success(
            (await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-document-open-table", { handleId: api5DocumentHandle, tableReference: api5ProofReference }, 30000)).result,
            "command-proof.api5-document-open-table",
          );
          api5OpenedTable = true;
          expectApi5Success(
            (await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-document-clear-table-cell-text", { handleId: api5DocumentHandle, cellNumber: toApi5CellNumber(0, 0, 1) }, 30000)).result,
            "command-proof.api5-document-clear-table-cell-text",
          );
          expectApi5Success(
            (await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.api5-document-end-obj", { handleId: api5DocumentHandle }, 30000)).result,
            "command-proof.api5-document-end-obj",
          );
          api5OpenedTable = false;
          api5ProofHandleId = await reopenTableHandleByReference(utilityUrl, pairingToken, pagesOrigin, api5ProofReference);
          await executeCommand(utilityUrl, pairingToken, pagesOrigin, "kompas-pages-executor", "xlsx-to-kompas-tbl.table-update", { handleId: api5ProofHandleId }, 30000);
          const api5ClearedTextExecution = await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-cell-get-text",
            { handleId: api5ProofHandleId, rowIndex: 0, columnIndex: 0 },
            15000,
          );
          assertCondition(String(api5ClearedTextExecution.result || "") === "", "Command proof API5 clear did not empty the proof cell", {
            iteration,
            api5ClearedText: api5ClearedTextExecution.result,
          });

          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-delete",
            { handleId: insertedHandleId },
            30000,
          );
          insertedHandleId = "";
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-delete",
            { handleId: reopenedHandleId },
            30000,
          );
          reopenedHandleId = "";
          tableHandleId = "";
          await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.table-delete",
            { handleId: api5ProofHandleId },
            30000,
          );
          api5ProofHandleId = "";

          const tableCountAfter = Number((await executeCommand(
            utilityUrl,
            pairingToken,
            pagesOrigin,
            "kompas-pages-executor",
            "xlsx-to-kompas-tbl.view-table-count",
            { refresh: true },
            15000,
          )).result) || 0;
          assertCondition(tableCountAfter === tableCountBefore, "Command proof table count did not return to baseline after cleanup", {
            iteration,
            tableCountBefore,
            tableCountAfter,
          });

          sampleIterations.push({
            iteration,
            documentPath: commandProofDocumentPath,
            viewX: Number(viewXExecution.result) || 0,
            viewY: Number(viewYExecution.result) || 0,
            frameCenterX,
            frameCenterY,
            tableCountBefore,
            tableCountAfter,
            lineAlign: Number(lineAlignExecution.result) || 0,
            itemBold: coerceBooleanLike(itemBoldExecution.result),
            itemItalic: coerceBooleanLike(itemItalicExecution.result),
            itemUnderline: coerceBooleanLike(itemUnderlineExecution.result),
            itemColor: Number(itemColorExecution.result) || 0,
            temp: coerceBooleanLike(tempExecution.result),
            valid: coerceBooleanLike(validExecution.result),
            savePath,
            api5CellAddress: "A1",
            api5CellText: String(api5ProofTextExecution.result || ""),
          });
        } finally {
          if (api5OpenedTable && api5DocumentHandle) {
            try {
              await executeCommand(
                utilityUrl,
                pairingToken,
                pagesOrigin,
                "kompas-pages-executor",
                "xlsx-to-kompas-tbl.api5-document-end-obj",
                { handleId: api5DocumentHandle },
                30000,
              );
            } catch {}
          }
          if (insertedHandleId) {
            try {
              await executeCommand(
                utilityUrl,
                pairingToken,
                pagesOrigin,
                "kompas-pages-executor",
                "xlsx-to-kompas-tbl.table-delete",
                { handleId: insertedHandleId },
                30000,
              );
            } catch {}
          }
          if (api5ProofHandleId) {
            try {
              await executeCommand(
                utilityUrl,
                pairingToken,
                pagesOrigin,
                "kompas-pages-executor",
                "xlsx-to-kompas-tbl.table-delete",
                { handleId: api5ProofHandleId },
                30000,
              );
            } catch {}
          }
          const cleanupHandleId = reopenedHandleId || tableHandleId;
          if (cleanupHandleId) {
            try {
              await executeCommand(
                utilityUrl,
                pairingToken,
                pagesOrigin,
                "kompas-pages-executor",
                "xlsx-to-kompas-tbl.table-delete",
                { handleId: cleanupHandleId },
                30000,
              );
            } catch {}
          }
        }
      }

      const trackedCommandCounts = Object.fromEntries(
        Object.keys(COMMAND_PROOF_THRESHOLDS)
          .sort((left, right) => left.localeCompare(right))
          .map((commandId) => [commandId, snapshotCommandCounts(activeCommandStats)[commandId] || 0]),
      );
      const missingCommands = Object.entries(COMMAND_PROOF_THRESHOLDS)
        .filter(([commandId, minimumCount]) => (trackedCommandCounts[commandId] || 0) < minimumCount)
        .map(([commandId, minimumCount]) => ({
          commandId,
          minimumCount,
          actualCount: trackedCommandCounts[commandId] || 0,
        }));
      assertCondition(missingCommands.length === 0, "Command proof did not reach required live invocation thresholds", { missingCommands });

      return {
        iterationCount,
        minimumCommandThresholds: COMMAND_PROOF_THRESHOLDS,
        trackedCommandCounts,
        missingCommands,
        sampleIterations,
      };
    }

    if (args.scenario === "workflow" || args.scenario === "all") {
      const workflow = await runWorkflowScenario();
      scenarios.push({
        name: "workflow",
        ...workflow,
      });
      if (workflow.stopAfterInline) {
        Object.assign(report, {
          success: true,
          scenarios,
          stopAfterInline: true,
        });
        fs.writeFileSync(path.join(artifactRoot, "report.json"), `${JSON.stringify(report, null, 2)}\n`, "utf8");
        process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
        return;
      }
    }

    if (args.scenario === "rich-proof" || args.scenario === "all") {
      const richProof = await runRichProofScenario();
      scenarios.push({
        name: "rich-proof",
        ...richProof,
      });
    }

    if (args.scenario === "autofit-proof" || args.scenario === "all") {
      const autoFitProof = await runAutoFitProofScenario();
      scenarios.push({
        name: "autofit-proof",
        ...autoFitProof,
      });
    }

    if (args.scenario === "command-proof" || args.scenario === "all") {
      const commandProof = await runCommandProofScenario();
      scenarios.push({
        name: "command-proof",
        ...commandProof,
      });
    }

    Object.assign(report, {
      success: true,
      scenario: args.scenario,
      scenarios,
      commandInvocationCounts: snapshotCommandCounts(activeCommandStats),
    });
    fs.writeFileSync(path.join(artifactRoot, "report.json"), `${JSON.stringify(report, null, 2)}\n`, "utf8");
    process.stdout.write(`${JSON.stringify(report, null, 2)}\n`);
  } catch (error) {
    report.error = String(error.stack || error);
    fs.writeFileSync(path.join(artifactRoot, "report.json"), `${JSON.stringify(report, null, 2)}\n`, "utf8");
    throw error;
  } finally {
    activeCommandStats = null;
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
