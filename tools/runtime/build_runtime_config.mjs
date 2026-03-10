import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..", "..");

function uniqueStrings(values) {
  return [...new Set((values || []).filter(Boolean).map((value) => String(value).trim()).filter(Boolean))];
}

function buildRuntimeConfig(options = {}) {
  const outputPath = options.outputPath || path.join(repoRoot, "out", "web-pages-runtime", "config.bootstrap.json");
  const listenUrl = options.listenUrl || "http://127.0.0.1:38741";
  const uiUrl = options.uiUrl || "http://127.0.0.1:5511/index.html";
  const pairingToken = options.pairingToken || "replace-this-token";
  const allowedOrigins = uniqueStrings(options.allowedOrigins || ["http://127.0.0.1:5511", "http://localhost:5511"]);
  const logFilePath = options.logFilePath || path.join(repoRoot, "out", "web-pages-runtime", "utility.log");
  const diagnosticsDirectory = options.diagnosticsDirectory || path.join(repoRoot, "out", "web-pages-runtime", "diagnostics");
  const profileDirectory = options.profileDirectory || path.join(repoRoot, "out", "web-pages-runtime", "profiles");
  const cacheDirectory = options.cacheDirectory || path.join(repoRoot, "out", "web-pages-runtime", "cache");

  const config = {
    Versions: {
      UtilityVersion: "1.0.0",
      ConfigVersion: options.configVersion || `kompas-pages-bootstrap-${new Date().toISOString()}`,
      ConfigSchemaVersion: 2,
    },
    Metadata: {
      ProductName: "KOMPAS Pages Executor",
      ProductCode: "kompas3d-utility.pages.executor",
      Author: "Codex",
      Description: "Bootstrap config for static Pages executor over WebBridge.Utility.",
      RepositoryUrl: "https://github.com/Efnatii/kompas3d-utility",
    },
    Runtime: {
      EnvironmentName: options.environmentName || "PagesBootstrap",
      DevMode: false,
      NoBrowser: true,
    },
    Server: {
      ListenUrl: listenUrl,
    },
    Ui: {
      Url: uiUrl,
      OpenMode: "Never",
      SessionWaitSeconds: 5,
    },
    Lifecycle: {
      ShutdownPolicy: "WhenIdle",
      IdleSeconds: Number.isFinite(options.idleSeconds) ? Number(options.idleSeconds) : 180,
    },
    Logging: {
      Level: options.logLevel || "Information",
      DebugMode: Boolean(options.debugMode),
      FilePath: logFilePath,
    },
    Storage: {
      ProfileDirectory: profileDirectory,
      CacheDirectory: cacheDirectory,
      DiagnosticsDirectory: diagnosticsDirectory,
    },
    Catalog: {
      Profiles: [],
    },
    Adapters: {
      Com: [],
      System: {},
    },
    Security: {
      LoopbackOnly: true,
      PairingToken: pairingToken,
      AllowedOrigins: allowedOrigins,
    },
    Session: {
      HeartbeatIntervalSeconds: 10,
      HeartbeatTimeoutSeconds: 30,
      PresenceTimeoutSeconds: 60,
      SuppressAutoOpenOnPresenceSessions: true,
      SweepIntervalSeconds: 2,
    },
  };

  return { outputPath, config };
}

function parseArgs(argv) {
  const parsed = {
    allowedOrigins: [],
  };

  for (let index = 2; index < argv.length; index += 1) {
    const token = argv[index];
    const value = argv[index + 1];
    switch (token) {
      case "--output":
        parsed.outputPath = value;
        index += 1;
        break;
      case "--listen-url":
        parsed.listenUrl = value;
        index += 1;
        break;
      case "--ui-url":
        parsed.uiUrl = value;
        index += 1;
        break;
      case "--pairing-token":
        parsed.pairingToken = value;
        index += 1;
        break;
      case "--origin":
        parsed.allowedOrigins.push(value);
        index += 1;
        break;
      case "--log-file":
        parsed.logFilePath = value;
        index += 1;
        break;
      case "--diagnostics-dir":
        parsed.diagnosticsDirectory = value;
        index += 1;
        break;
      case "--profile-dir":
        parsed.profileDirectory = value;
        index += 1;
        break;
      case "--cache-dir":
        parsed.cacheDirectory = value;
        index += 1;
        break;
      case "--idle-seconds":
        parsed.idleSeconds = Number.parseInt(value, 10);
        index += 1;
        break;
      case "--environment":
        parsed.environmentName = value;
        index += 1;
        break;
      case "--config-version":
        parsed.configVersion = value;
        index += 1;
        break;
      case "--debug":
        parsed.debugMode = true;
        break;
      default:
        throw new Error(`Unknown argument: ${token}`);
    }
  }

  return parsed;
}

function writeConfig(outputPath, config) {
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, `${JSON.stringify(config, null, 2)}\n`, "utf8");
}

function main() {
  const options = parseArgs(process.argv);
  const runtime = buildRuntimeConfig(options);
  writeConfig(runtime.outputPath, runtime.config);
  process.stdout.write(`${runtime.outputPath}\n`);
}

export { buildRuntimeConfig };

if (process.argv[1] && path.resolve(process.argv[1]) === __filename) {
  main();
}
