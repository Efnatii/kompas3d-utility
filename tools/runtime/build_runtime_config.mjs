import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const repoRoot = path.resolve(__dirname, "..", "..");
const defaultBridgeScriptPath = path.join(repoRoot, "scripts", "webbridge_xlsx_to_kompas_tbl.ps1");

function arg(name, converter = undefined) {
  const payload = { FromArgument: name };
  if (converter) {
    payload.Converter = converter;
  }
  return payload;
}

function literal(value, converter = undefined) {
  const payload = { Literal: value };
  if (converter) {
    payload.Converter = converter;
  }
  return payload;
}

function step(operation, member = "", options = {}) {
  const payload = { Operation: operation, Member: member };
  if (options.args?.length) {
    payload.Args = options.args;
  }
  if (options.valueArgument) {
    payload.ValueArgument = options.valueArgument;
  }
  if (options.storeAs) {
    payload.StoreAs = options.storeAs;
  }
  return payload;
}

function command(adapter, root, chain, options = {}) {
  const payload = {
    Adapter: adapter,
    Invoke: {
      Root: root,
      Chain: chain,
    },
  };
  if (options.defaultArguments) {
    payload.DefaultArguments = options.defaultArguments;
  }
  if (options.returnPath) {
    payload.Invoke.ReturnPath = options.returnPath;
  }
  return payload;
}

function powershellArguments(scriptPath, action) {
  return `-NoProfile -ExecutionPolicy Bypass -File "${scriptPath}" -Action ${action}`;
}

function bridgeCommand(scriptPath, action, timeoutMilliseconds) {
  return command(
    "system",
    "command",
    [
      step("call", "Run", {
        args: [
          literal("powershell.exe", "string"),
          literal(powershellArguments(scriptPath, action), "string"),
          literal(repoRoot, "path"),
          literal(timeoutMilliseconds, "int"),
          arg("stdin", "string"),
          literal(null),
        ],
      }),
    ],
    {
      defaultArguments: {
        stdin: null,
      },
    },
  );
}

function buildRuntimeConfig(options = {}) {
  const listenUrl = options.listenUrl || "http://127.0.0.1:38741";
  const uiUrl = options.uiUrl || "http://127.0.0.1:5511/index.html";
  const pairingToken = options.pairingToken || "kompas-pages-local";
  const allowedOrigins = Array.from(new Set(options.allowedOrigins || ["http://127.0.0.1:5511"]));
  const outputPath = options.outputPath || path.join(repoRoot, "out", "web-pages-runtime", "config.runtime.json");
  const logFilePath = options.logFilePath || path.join(repoRoot, "out", "web-pages-runtime", "utility.log");
  const diagnosticsDirectory = options.diagnosticsDirectory || path.join(repoRoot, "out", "web-pages-runtime", "diagnostics");
  const profileDirectory = options.profileDirectory || path.join(repoRoot, "out", "web-pages-runtime", "profiles");
  const cacheDirectory = options.cacheDirectory || path.join(repoRoot, "out", "web-pages-runtime", "cache");
  const bridgeScriptPath = options.bridgeScriptPath || defaultBridgeScriptPath;
  const profileId = options.profileId || "kompas-pages";
  const idleSeconds = Number.isFinite(options.idleSeconds) ? Number(options.idleSeconds) : 180;

  const config = {
    Versions: {
      UtilityVersion: "1.0.0",
      ConfigVersion: options.configVersion || `kompas-pages-${new Date().toISOString()}`,
      ConfigSchemaVersion: 2,
    },
    Metadata: {
      ProductName: "KOMPAS Pages",
      ProductCode: "kompas3d-utility.pages",
      Author: "Codex",
      Description: "Static Pages UI for browser-side XLSX to KOMPAS TBL export via WebBridge.Utility.",
      RepositoryUrl: "https://github.com/",
    },
    Runtime: {
      EnvironmentName: options.environmentName || "Pages",
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
      IdleSeconds: idleSeconds,
    },
    Logging: {
      Level: "Debug",
      DebugMode: true,
      FilePath: logFilePath,
    },
    Storage: {
      ProfileDirectory: profileDirectory,
      CacheDirectory: cacheDirectory,
      DiagnosticsDirectory: diagnosticsDirectory,
    },
    Catalog: {
      Profiles: [
        {
          ProfileId: profileId,
          ConfigSchemaVersion: 1,
          Description: "Browser-driven Pages profile for xlsx-to-kompas-tbl.",
          Checksum: `kompas-pages-${Date.now()}`,
          Commands: {
            "kompas.pages.status": {
              CommandId: "kompas.pages.status",
              ...bridgeCommand(bridgeScriptPath, "status", 30000),
            },
            "kompas.pages.open-document": {
              CommandId: "kompas.pages.open-document",
              ...bridgeCommand(bridgeScriptPath, "open-document", 60000),
            },
            "kompas.pages.export": {
              CommandId: "kompas.pages.export",
              ...bridgeCommand(bridgeScriptPath, "export", 180000),
            },
            "system.file.exists": {
              CommandId: "system.file.exists",
              ...command("system", "type:System.IO.File", [step("call", "Exists", { args: [arg("path", "path")] })]),
            },
            "system.file.read-bytes": {
              CommandId: "system.file.read-bytes",
              ...command("system", "type:System.IO.File", [step("call", "ReadAllBytes", { args: [arg("path", "path")] })]),
            },
            "system.generic.file-info.length": {
              CommandId: "system.generic.file-info.length",
              ...command("system", "type:System.IO.FileInfo", [
                step("new", "", { args: [arg("path", "path")] }),
                step("get", "Length"),
              ]),
            },
          },
        },
      ],
    },
    Adapters: {
      Com: [],
      System: {
        AllowedTypeNames: [
          "System.IO.File",
          "System.IO.FileInfo",
        ],
        AllowedProcessExecutables: [
          "powershell.exe",
        ],
      },
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

  return {
    outputPath,
    config,
  };
}

function parseCliArgs(argv) {
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
      case "--bridge-script":
        parsed.bridgeScriptPath = value;
        index += 1;
        break;
      case "--profile-id":
        parsed.profileId = value;
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
      default:
        throw new Error(`Unknown argument: ${token}`);
    }
  }

  return parsed;
}

function writeRuntimeConfig(outputPath, config) {
  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, `${JSON.stringify(config, null, 2)}\n`, "utf8");
}

function main() {
  const options = parseCliArgs(process.argv);
  const runtime = buildRuntimeConfig(options);
  writeRuntimeConfig(runtime.outputPath, runtime.config);
  process.stdout.write(`${runtime.outputPath}\n`);
}

export { buildRuntimeConfig };

if (process.argv[1] && path.resolve(process.argv[1]) === __filename) {
  main();
}
