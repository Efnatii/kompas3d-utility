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
        if (!["all", "workflow", "rich-proof"].includes(value)) {
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

function parseInlineOutputPath(resultText) {
  const match = String(resultText || "").match(/^Inline \| (.+?) \| anchor=/);
  return match ? match[1].trim() : "";
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
  return [[
    {
      address: "A1",
      rowIndex: 0,
      columnIndex: 0,
      text: "Bold + Blue\nItalic",
      horizontal: "center",
      alignCode: 1,
      wrapText: true,
      oneLine: false,
      hasContent: true,
      lines: [
        {
          items: [
            {
              text: "Bold",
              fontName: "Arial",
              heightMm: 4.9389,
              bold: true,
              italic: false,
              underline: false,
              color: 16711680,
              widthFactor: 1,
            },
            {
              text: " + ",
              fontName: "Calibri",
              heightMm: 3.8806,
              bold: false,
              italic: false,
              underline: false,
              color: 0,
              widthFactor: 1,
            },
            {
              text: "Blue",
              fontName: "Arial",
              heightMm: 4.2333,
              bold: false,
              italic: true,
              underline: true,
              color: 255,
              widthFactor: 1,
            },
          ],
        },
        {
          items: [
            {
              text: "Italic",
              fontName: "Arial",
              heightMm: 4.2333,
              bold: false,
              italic: true,
              underline: true,
              color: 255,
              widthFactor: 1,
            },
          ],
        },
      ],
    },
    {
      address: "B1",
      rowIndex: 0,
      columnIndex: 1,
      text: "Plain styled",
      horizontal: "right",
      alignCode: 2,
      wrapText: false,
      oneLine: true,
      hasContent: true,
      lines: [
        {
          items: [
            {
              text: "Plain styled",
              fontName: "Arial",
              heightMm: 4.5861,
              bold: true,
              italic: true,
              underline: true,
              color: 32768,
              widthFactor: 1,
            },
          ],
        },
      ],
    },
  ]];
}

function createExpectedRichSnapshots() {
  return {
    a1: normalizeCellSnapshot({
      text: "Bold + Blue\nItalic",
      oneLine: false,
      lineCount: 2,
      lines: [
        {
          align: 1,
          items: createExpectedRichCellMatrix()[0][0].lines[0].items,
        },
        {
          align: 1,
          items: createExpectedRichCellMatrix()[0][0].lines[1].items,
        },
      ],
    }),
    b1: normalizeCellSnapshot({
      text: "Plain styled",
      oneLine: true,
      lineCount: 1,
      lines: [
        {
          align: 2,
          items: createExpectedRichCellMatrix()[0][1].lines[0].items,
        },
      ],
    }),
  };
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

function createRichFixture(workspaceRoot) {
  const fixturePath = path.join(workspaceRoot, "fixture-rich.xlsx");
  const script = `
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment
from openpyxl.cell.rich_text import CellRichText, TextBlock, InlineFont

path = r"""${fixturePath}"""
wb = Workbook()
ws = wb.active
ws.title = "RichProof"
ws["A1"] = CellRichText(
    TextBlock(InlineFont(rFont="Arial", b=True, sz=14, color="FF0000"), "Bold"),
    " + ",
    TextBlock(InlineFont(rFont="Arial", i=True, u="single", sz=12, color="0000FF"), "Blue\\nItalic"),
)
ws["A1"].alignment = Alignment(horizontal="center", wrap_text=True)
ws["B1"] = "Plain styled"
ws["B1"].font = Font(name="Arial", bold=True, italic=True, underline="single", size=13, color="008000")
ws["B1"].alignment = Alignment(horizontal="right", wrap_text=False)
ws.column_dimensions["A"].width = 24
ws.column_dimensions["B"].width = 18
wb.save(path)
print(path)
`;
  runPython(script);
  assertCondition(fs.existsSync(fixturePath), "Rich fixture was not created", { fixturePath });
  return fixturePath;
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
          "xlsx-to-kompas-tbl.api5-create-dynamic-array",
          { arrayType: API5_TEXT_LINE_ARRAY_TYPE },
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
            "xlsx-to-kompas-tbl.api5-create-dynamic-array",
            { arrayType: API5_TEXT_ITEM_ARRAY_TYPE },
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
            "xlsx-to-kompas-tbl.api5-text-line-param-set-item-array",
            {
              handleId: lineHandle,
              itemArrayHandle,
            },
            15000,
          )).result,
          "api5-text-line-param-set-item-array",
        );
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
          "xlsx-to-kompas-tbl.api5-text-param-set-line-array",
          {
            handleId: textParamHandle,
            lineArrayHandle,
          },
          15000,
        )).result,
        "api5-text-param-set-line-array",
      );
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

  const updateExecution = await executeCommand(
    baseUrl,
    pairingToken,
    origin,
    "kompas-pages-executor",
    "xlsx-to-kompas-tbl.table-update",
    { handleId: tableHandleId },
    30000,
  );
  assertCondition(updateExecution.result !== false, "KOMPAS returned Update=false after API5 table population.", updateExecution);

  const alignCommands = [];
  for (const cell of writes) {
    alignCommands.push({
      commandId: "xlsx-to-kompas-tbl.table-cell-set-one-line",
      arguments: {
        handleId: tableHandleId,
        rowIndex: cell.rowIndex,
        columnIndex: cell.columnIndex,
        oneLine: Boolean(cell.oneLine),
      },
      timeoutMilliseconds: 15000,
      profileId: "kompas-pages-executor",
    });
    for (let lineIndex = 0; lineIndex < cell.lines.length; lineIndex += 1) {
      alignCommands.push({
        commandId: "xlsx-to-kompas-tbl.table-cell-set-line-align",
        arguments: {
          handleId: tableHandleId,
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
      const autoFollowPath = await page.locator("#xlsx-output-path").inputValue();
      assertCondition(autoFollowPath && autoFollowPath.toLowerCase().endsWith(".tbl"), "Expected auto-follow output path", { autoFollowPath });

      await page.locator("#xlsx-cell-width").fill("10");
      await page.locator("#xlsx-cell-height").fill("5");
      await page.waitForFunction(() => document.querySelector("#xlsx-table-width")?.value === "130");
      await page.waitForFunction(() => document.querySelector("#xlsx-table-height")?.value === "40");

      await page.locator("#xlsx-output-path").fill(outputTablePath);
      await page.locator("#xlsx-follow-button").click();
      await page.waitForFunction((expected) => document.querySelector("#xlsx-output-path")?.value === expected, autoFollowPath);
      await page.locator("#xlsx-output-path").fill(outputTablePath);

      await page.locator("#xlsx-inline-button").click();
      const inlineResultText = await waitForResultBoxState(page, "Inline |", {
        timeout: 180000,
        progressPrefix: "Inline",
        allowedTexts: ["Matrix loaded. Export or Inline is ready."],
      });
      const inlineArtifactPath = parseInlineOutputPath(inlineResultText);
      await page.screenshot({ path: path.join(screenshotsRoot, "25-workflow-inline.png"), fullPage: true });
      assertCondition(!fs.existsSync(outputTablePath), "Inline must not create the manual output path", { outputTablePath });
      assertCondition(fs.existsSync(inlineArtifactPath), "Inline temp .tbl was not created", { inlineArtifactPath });
      assertCondition(fs.statSync(inlineArtifactPath).size > 0, "Inline temp .tbl is empty", { inlineArtifactPath });

      if (args.pauseAfterInlineMs > 0) {
        await page.waitForTimeout(args.pauseAfterInlineMs);
        await page.screenshot({ path: path.join(screenshotsRoot, "26-inline-paused.png"), fullPage: true });
      }

      if (args.stopAfterInline) {
        return {
          autoFollowPath,
          inlineArtifactPath,
          inlineArtifactSize: fs.statSync(inlineArtifactPath).size,
          inlineResultText,
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
        autoFollowPath,
        inlineArtifactPath,
        inlineArtifactSize: fs.statSync(inlineArtifactPath).size,
        inlineResultText,
        outputTablePath,
        outputTableSize: outputStat.size,
        insertResultText,
        downloadedPath,
        downloadedSize: fs.statSync(downloadedPath).size,
      };
    }

    async function runRichProofScenario() {
      await resetXlsxModuleIfNeeded();
      const richFixturePath = createRichFixture(workspaceRoot);
      const richOutputPath = path.join(workspaceRoot, "rich-proof.tbl");
      const directRichPath = path.join(workspaceRoot, "rich-direct.tbl");
      const expectedRichSnapshots = createExpectedRichSnapshots();
      const expectedRichCellMatrix = createExpectedRichCellMatrix();

      let directHandleId = "";
      let directSavedHandleId = "";
      const directCreateExecution = await executeCommand(
        utilityUrl,
        pairingToken,
        pagesOrigin,
        "kompas-pages-executor",
        "xlsx-to-kompas-tbl.create-table",
        {
          rows: 1,
          cols: 2,
          cellWidthMm: 20,
          cellHeightMm: 8,
        },
        60000,
      );
      directHandleId = String(directCreateExecution.result?.handleId || "");
      assertCondition(directHandleId, "Direct rich proof create-table did not return a handle");
      await populateTableHandleDirect(utilityUrl, pairingToken, pagesOrigin, directHandleId, expectedRichCellMatrix);
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
      assertCondition(directUpdateExecution.result !== false, "KOMPAS returned Update=false for direct rich proof handle");
      const directLiveA1 = await readCellSnapshot(utilityUrl, pairingToken, pagesOrigin, directHandleId, 0, 0);
      const directLiveB1 = await readCellSnapshot(utilityUrl, pairingToken, pagesOrigin, directHandleId, 0, 1);
      const normalizedDirectLiveA1 = normalizeCellSnapshot(directLiveA1);
      const normalizedDirectLiveB1 = normalizeCellSnapshot(directLiveB1);
      assertCondition(
        JSON.stringify(normalizedDirectLiveA1) === JSON.stringify(expectedRichSnapshots.a1),
        "Direct in-memory rich A1 snapshot mismatch before Save",
        { expected: expectedRichSnapshots.a1, actual: normalizedDirectLiveA1, raw: directLiveA1 },
      );
      assertCondition(
        JSON.stringify(normalizedDirectLiveB1) === JSON.stringify(expectedRichSnapshots.b1),
        "Direct in-memory rich B1 snapshot mismatch before Save",
        { expected: expectedRichSnapshots.b1, actual: normalizedDirectLiveB1, raw: directLiveB1 },
      );
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
      const directSavedA1 = await readCellSnapshot(utilityUrl, pairingToken, pagesOrigin, directSavedHandleId, 0, 0);
      const directSavedB1 = await readCellSnapshot(utilityUrl, pairingToken, pagesOrigin, directSavedHandleId, 0, 1);
      const normalizedDirectSavedA1 = normalizeCellSnapshot(directSavedA1);
      const normalizedDirectSavedB1 = normalizeCellSnapshot(directSavedB1);

      await loadFixtureViaUi(richFixturePath, "1 x 2", "Plain styled", "70-rich-uploaded.png");
      await page.locator("#xlsx-output-path").fill(richOutputPath);

      const inlineCountBefore = await readTableCount();
      await page.locator("#xlsx-inline-button").click();
      const inlineResultText = await waitForResultBoxState(page, "Inline |", {
        timeout: 180000,
        progressPrefix: "Inline",
        allowedTexts: ["Matrix loaded. Export or Inline is ready."],
      });
      const inlineArtifactPath = parseInlineOutputPath(inlineResultText);
      await page.screenshot({ path: path.join(screenshotsRoot, "75-rich-inline.png"), fullPage: true });
      const inlineCountAfter = await readTableCount();
      assertCondition(fs.existsSync(inlineArtifactPath), "Rich inline temp artifact missing", { inlineArtifactPath });
      assertCondition(fs.statSync(inlineArtifactPath).size > 0, "Rich inline temp artifact is empty", { inlineArtifactPath });

      await page.locator("#xlsx-export-button").click();
      await page.locator("#xlsx-result-box").filter({ hasText: "OK" }).waitFor({ timeout: 180000 });
      await page.screenshot({ path: path.join(screenshotsRoot, "80-rich-exported.png"), fullPage: true });
      assertCondition(fs.existsSync(richOutputPath), "Rich export file missing", { richOutputPath });
      assertCondition(fs.statSync(richOutputPath).size > 0, "Rich export file is empty", { richOutputPath });

      const exportedInsert = await insertTableDirect(richOutputPath);
      const exportedHandleId = String(exportedInsert.result?.handleId || "");
      assertCondition(exportedHandleId, "Insert of exported rich table did not return handle");

      const inlineInsert = await insertTableDirect(inlineArtifactPath);
      const inlineHandleId = String(inlineInsert.result?.handleId || "");
      assertCondition(inlineHandleId, "Insert of inline rich artifact did not return handle");

      const exportedA1 = await readCellSnapshot(utilityUrl, pairingToken, pagesOrigin, exportedHandleId, 0, 0);
      const exportedB1 = await readCellSnapshot(utilityUrl, pairingToken, pagesOrigin, exportedHandleId, 0, 1);
      const inlineA1 = await readCellSnapshot(utilityUrl, pairingToken, pagesOrigin, inlineHandleId, 0, 0);
      const inlineB1 = await readCellSnapshot(utilityUrl, pairingToken, pagesOrigin, inlineHandleId, 0, 1);
      const normalizedExportedA1 = normalizeCellSnapshot(exportedA1);
      const normalizedExportedB1 = normalizeCellSnapshot(exportedB1);
      const normalizedInlineA1 = normalizeCellSnapshot(inlineA1);
      const normalizedInlineB1 = normalizeCellSnapshot(inlineB1);

      assertCondition(
        JSON.stringify(normalizedDirectSavedA1) === JSON.stringify(expectedRichSnapshots.a1),
        "Direct saved rich A1 snapshot mismatch after Save/Insert",
        { expected: expectedRichSnapshots.a1, actual: normalizedDirectSavedA1, raw: directSavedA1 },
      );
      assertCondition(
        JSON.stringify(normalizedDirectSavedB1) === JSON.stringify(expectedRichSnapshots.b1),
        "Direct saved rich B1 snapshot mismatch after Save/Insert",
        { expected: expectedRichSnapshots.b1, actual: normalizedDirectSavedB1, raw: directSavedB1 },
      );

      assertCondition(
        JSON.stringify(normalizedExportedA1) === JSON.stringify(expectedRichSnapshots.a1),
        "Exported rich A1 snapshot mismatch",
        { expected: expectedRichSnapshots.a1, actual: normalizedExportedA1, raw: exportedA1 },
      );
      assertCondition(
        JSON.stringify(normalizedExportedB1) === JSON.stringify(expectedRichSnapshots.b1),
        "Exported rich B1 snapshot mismatch",
        { expected: expectedRichSnapshots.b1, actual: normalizedExportedB1, raw: exportedB1 },
      );

      assertCondition(
        JSON.stringify(normalizedExportedA1) === JSON.stringify(normalizedInlineA1),
        "Inline rich artifact A1 snapshot differs from export artifact",
        { exportedA1, inlineA1, normalizedExportedA1, normalizedInlineA1 },
      );
      assertCondition(
        JSON.stringify(normalizedExportedB1) === JSON.stringify(normalizedInlineB1),
        "Inline rich artifact B1 snapshot differs from export artifact",
        { exportedB1, inlineB1, normalizedExportedB1, normalizedInlineB1 },
      );

      return {
        richFixturePath,
        inlineArtifactPath,
        inlineArtifactSize: fs.statSync(inlineArtifactPath).size,
        inlineCountBefore,
        inlineCountAfter,
        richOutputPath,
        richOutputSize: fs.statSync(richOutputPath).size,
        exportedHandleId,
        inlineHandleId,
        exportedA1,
        exportedB1,
        inlineA1,
        inlineB1,
        directRichPath,
        directRichSize: fs.existsSync(directRichPath) ? fs.statSync(directRichPath).size : 0,
        directHandleId,
        directSavedHandleId,
        directLiveA1,
        directLiveB1,
        directSavedA1,
        directSavedB1,
        normalizedExportedA1,
        normalizedExportedB1,
        normalizedInlineA1,
        normalizedInlineB1,
        normalizedDirectLiveA1,
        normalizedDirectLiveB1,
        normalizedDirectSavedA1,
        normalizedDirectSavedB1,
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

    Object.assign(report, {
      success: true,
      scenario: args.scenario,
      scenarios,
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
