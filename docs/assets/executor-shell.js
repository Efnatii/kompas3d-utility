const STORAGE_KEY = "kompas-pages.executor.settings";

const EXECUTOR_PROFILE_ID = "kompas-pages-executor";

const DEFAULT_EXECUTOR_SETTINGS = {
  utilityUrl: "http://127.0.0.1:38741",
  pairingToken: "replace-this-token",
  workspaceRoot: "C:\\_GIT_\\kompas3d-utility",
  clientName: "kompas-pages-executor",
  uiVersion: "1.0.0",
};

function cloneJson(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function toArray(value) {
  return Array.isArray(value) ? value : [];
}

function uniqueStrings(values) {
  const seen = new Set();
  const result = [];
  for (const value of values) {
    if (typeof value !== "string") {
      continue;
    }
    const trimmed = value.trim();
    if (!trimmed) {
      continue;
    }
    const key = trimmed.toLowerCase();
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(trimmed);
  }
  return result;
}

function findCaseInsensitiveKey(source, key) {
  if (!source || typeof source !== "object") {
    return undefined;
  }
  if (Object.prototype.hasOwnProperty.call(source, key)) {
    return key;
  }
  const wanted = String(key).toLowerCase();
  return Object.keys(source).find((entry) => entry.toLowerCase() === wanted);
}

function getCaseInsensitive(source, key, fallback = undefined) {
  const matchedKey = findCaseInsensitiveKey(source, key);
  return matchedKey ? source[matchedKey] : fallback;
}

function ensureObject(value) {
  return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}

function formatStamp() {
  return new Date().toLocaleTimeString("ru-RU", { hour12: false });
}

function basePageUrl() {
  return `${window.location.origin}${window.location.pathname}`;
}

function buildPowerShellArguments(scriptPath, action) {
  return `-NoProfile -ExecutionPolicy Bypass -File "${scriptPath}" -Action ${action}`;
}

const runtimeDsl = {
  arg(name, converter) {
    const payload = { FromArgument: name };
    if (converter) {
      payload.Converter = converter;
    }
    return payload;
  },

  literal(value, converter) {
    const payload = { Literal: value };
    if (converter) {
      payload.Converter = converter;
    }
    return payload;
  },

  step(operation, member = "", options = {}) {
    const payload = {
      Operation: operation,
      Member: member,
    };
    if (options.args && options.args.length) {
      payload.Args = options.args;
    }
    if (options.valueArgument) {
      payload.ValueArgument = options.valueArgument;
    }
    if (options.storeAs) {
      payload.StoreAs = options.storeAs;
    }
    return payload;
  },

  command(adapter, root, chain, options = {}) {
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
  },

  createProcessCommand({ scriptPath, action, workingDirectory, timeoutMilliseconds = 60000 }) {
    return runtimeDsl.command(
      "system",
      "command",
      [
        runtimeDsl.step("call", "Run", {
          args: [
            runtimeDsl.literal("powershell.exe", "string"),
            runtimeDsl.literal(buildPowerShellArguments(scriptPath, action), "string"),
            runtimeDsl.literal(workingDirectory, "path"),
            runtimeDsl.literal(timeoutMilliseconds, "int"),
            runtimeDsl.arg("stdin", "string"),
            runtimeDsl.literal(null),
          ],
        }),
      ],
      {
        defaultArguments: {
          stdin: null,
        },
      },
    );
  },
};

function unwrapEnvelope(payload) {
  if (payload && typeof payload === "object" && Object.prototype.hasOwnProperty.call(payload, "payload")) {
    return payload.payload;
  }
  return payload;
}

function extractApiError(payload) {
  if (!payload || typeof payload !== "object") {
    return null;
  }
  if (payload.error && typeof payload.error === "object") {
    return payload.error;
  }
  if (payload.payload && typeof payload.payload === "object" && payload.payload.error) {
    return payload.payload.error;
  }
  return null;
}

function formatHttpError(action, response, payload) {
  const apiError = extractApiError(payload);
  if (apiError?.code === "pairing_token_invalid") {
    return `${action}: invalid pairing token`;
  }
  if (apiError?.code === "origin_not_allowed") {
    return `${action}: origin is not allowed`;
  }
  if (apiError?.code === "origin_missing") {
    return `${action}: origin header is missing`;
  }
  if (apiError?.message) {
    return `${action}: ${apiError.message}`;
  }
  return `${action} ${response.status}`;
}

class WebBridgeClient {
  constructor({ baseUrl, pairingToken, clientName, uiVersion }) {
    this.baseUrl = String(baseUrl || "").replace(/\/+$/, "");
    this.pairingToken = pairingToken;
    this.clientName = clientName;
    this.uiVersion = uiVersion;
    this.sessionId = null;
    this.heartbeatIntervalMs = 10000;
    this.socket = null;
    this.heartbeatTimer = null;
    this.helloReceived = false;
  }

  headers(extra = {}) {
    return {
      "Content-Type": "application/json",
      "X-KWB-Pairing-Token": this.pairingToken,
      ...extra,
    };
  }

  async fetchJson(pathname, init = {}) {
    const response = await fetch(`${this.baseUrl}${pathname}`, {
      ...init,
      headers: this.headers(init.headers || {}),
    });
    const text = await response.text();
    let payload = null;
    try {
      payload = text ? JSON.parse(text) : null;
    } catch {
      payload = text;
    }
    return { response, payload, text };
  }

  async connect() {
    const { response, payload } = await this.fetchJson("/session/register", {
      method: "POST",
      body: JSON.stringify({
        clientName: this.clientName,
        uiVersion: this.uiVersion,
        desiredSessionId: this.sessionId,
      }),
    });
    if (!response.ok) {
      throw new Error(formatHttpError("session/register", response, payload));
    }

    const registered = unwrapEnvelope(payload);
    if (!registered?.sessionId || !registered?.wsUrl) {
      throw new Error("session/register returned incomplete payload");
    }

    this.sessionId = registered.sessionId;
    this.heartbeatIntervalMs = Math.max(
      1000,
      Number(registered.heartbeatIntervalSeconds || 10) * 1000,
    );

    let wsUrl = registered.wsUrl;
    wsUrl += wsUrl.includes("?") ? "&" : "?";
    wsUrl += `token=${encodeURIComponent(this.pairingToken)}`;
    await this.openSocket(wsUrl);
    return registered;
  }

  async disconnect(reason = "manual-close") {
    this.stopHeartbeat();

    if (this.sessionId) {
      try {
        await this.fetchJson("/session/closing", {
          method: "POST",
          body: JSON.stringify({
            sessionId: this.sessionId,
            reason,
          }),
        });
      } catch {
        // Best effort.
      }
    }

    if (this.socket) {
      try {
        this.socket.close(1000, reason);
      } catch {
        // Ignore close failures.
      }
      this.socket = null;
    }
  }

  stopHeartbeat() {
    if (this.heartbeatTimer !== null) {
      window.clearInterval(this.heartbeatTimer);
      this.heartbeatTimer = null;
    }
  }

  send(type, payload = null) {
    if (!this.socket || this.socket.readyState !== WebSocket.OPEN) {
      throw new Error("WebSocket is not open");
    }
    this.socket.send(JSON.stringify({
      id: `${Date.now()}-${Math.random().toString(16).slice(2)}`,
      type,
      ts: new Date().toISOString(),
      correlationId: null,
      payload,
      error: null,
    }));
  }

  async execute(profileId, commandId, args = {}, timeoutMilliseconds = 60000) {
    const { response, payload } = await this.fetchJson("/commands/execute", {
      method: "POST",
      body: JSON.stringify({
        profileId,
        commandId,
        arguments: args,
        reportVerbosity: "Compact",
        timeoutMilliseconds,
      }),
    });

    const execution = unwrapEnvelope(payload);
    if (!response.ok || !execution?.success) {
      const errorMessage = execution?.error?.message
        || execution?.error?.code
        || extractApiError(payload)?.message
        || extractApiError(payload)?.code
        || response.status;
      throw new Error(`${commandId} failed: ${errorMessage}`);
    }
    return execution;
  }

  async openSocket(wsUrl) {
    this.stopHeartbeat();
    this.helloReceived = false;
    this.socket = new WebSocket(wsUrl);

    this.socket.addEventListener("message", (event) => {
      try {
        const envelope = JSON.parse(event.data);
        if (envelope?.type === "hello") {
          this.helloReceived = true;
        }
      } catch {
        // Ignore malformed telemetry frames.
      }
    });

    await new Promise((resolve, reject) => {
      const timeoutId = window.setTimeout(() => {
        reject(new Error("ws hello timeout"));
      }, 20000);

      const cleanup = () => {
        window.clearTimeout(timeoutId);
      };

      this.socket.addEventListener("open", () => {
        try {
          this.send("hello", {
            clientName: this.clientName,
            sessionId: this.sessionId,
          });
        } catch (error) {
          cleanup();
          reject(error);
        }
      });

      this.socket.addEventListener("error", () => {
        cleanup();
        reject(new Error("ws connection failed"));
      });

      const pollId = window.setInterval(() => {
        if (this.helloReceived) {
          window.clearInterval(pollId);
          cleanup();
          this.heartbeatTimer = window.setInterval(() => {
            try {
              this.send("heartbeat", { sessionId: this.sessionId });
            } catch {
              this.stopHeartbeat();
            }
          }, this.heartbeatIntervalMs);
          resolve();
          return;
        }

        if (!this.socket || this.socket.readyState === WebSocket.CLOSED) {
          window.clearInterval(pollId);
          cleanup();
          reject(new Error("ws closed before hello"));
        }
      }, 100);
    });
  }
}

function parseQuerySettings() {
  const params = new URLSearchParams(window.location.search);
  return {
    utilityUrl: params.get("utilityUrl") || "",
    pairingToken: params.get("pairingToken") || "",
    workspaceRoot: params.get("workspaceRoot") || "",
    autoConnect: params.get("autoConnect") === "1",
  };
}

function restoreStoredSettings() {
  const raw = window.localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return {};
  }
  try {
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

function createShellContribution() {
  return {
    commands: {
      "system.file.exists": runtimeDsl.command(
        "system",
        "type:System.IO.File",
        [
          runtimeDsl.step("call", "Exists", {
            args: [runtimeDsl.arg("path", "path")],
          }),
        ],
      ),
      "system.file.read-bytes": runtimeDsl.command(
        "system",
        "type:System.IO.File",
        [
          runtimeDsl.step("call", "ReadAllBytes", {
            args: [runtimeDsl.arg("path", "path")],
          }),
        ],
      ),
    },
    allowedTypes: [
      "System.IO.File",
      "System.IO.FileInfo",
    ],
    allowedProcesses: [],
  };
}

function createExecutorShell({ modules }) {
  if (!Array.isArray(modules) || modules.length === 0) {
    throw new Error("Executor shell requires at least one module.");
  }

  const ui = {
    tabs: document.getElementById("module-tabs"),
    bridgeBadge: document.getElementById("bridge-badge"),
    runtimeBadge: document.getElementById("runtime-badge"),
    moduleBadge: document.getElementById("module-badge"),
    utilityUrl: document.getElementById("utility-url"),
    pairingToken: document.getElementById("pairing-token"),
    workspaceRoot: document.getElementById("workspace-root"),
    connectButton: document.getElementById("connect-button"),
    disconnectButton: document.getElementById("disconnect-button"),
    reloadRuntimeButton: document.getElementById("reload-runtime-button"),
    bridgeMeta: document.getElementById("bridge-meta"),
    runtimeMeta: document.getElementById("runtime-meta"),
    moduleMeta: document.getElementById("module-meta"),
    moduleTitle: document.getElementById("module-title"),
    moduleSubtitle: document.getElementById("module-subtitle"),
    moduleHosts: document.getElementById("module-hosts"),
    clearLogButton: document.getElementById("clear-log-button"),
    logOutput: document.getElementById("log-output"),
  };

  const orderedModules = modules.map((module) => ({ ...module }));
  const moduleById = new Map(orderedModules.map((module) => [module.id, module]));
  const hostsById = new Map();
  const handlesById = new Map();
  const displayById = new Map();
  const events = new EventTarget();

  const state = {
    bridge: null,
    runtimeReady: false,
    runtimeVersion: "",
    activeModuleId: orderedModules[0].id,
    connectPromise: null,
  };

  function setBadge(element, label, active) {
    element.textContent = label;
    element.className = active ? "badge" : "badge badge--dim";
  }

  function appendLog(scope, message, detail = "") {
    const suffix = detail ? ` ${detail}` : "";
    ui.logOutput.textContent += `\n[${formatStamp()}] ${scope} ${message}${suffix}`;
    ui.logOutput.scrollTop = ui.logOutput.scrollHeight;
  }

  function replaceLog(message) {
    ui.logOutput.textContent = message;
  }

  function emit(name, detail = {}) {
    events.dispatchEvent(new CustomEvent(name, { detail }));
  }

  function persistSettings() {
    window.localStorage.setItem(STORAGE_KEY, JSON.stringify({
      utilityUrl: ui.utilityUrl.value.trim(),
      workspaceRoot: ui.workspaceRoot.value.trim(),
    }));
  }

  function loadSettings() {
    const query = parseQuerySettings();
    const stored = restoreStoredSettings();
    ui.utilityUrl.value = query.utilityUrl || stored.utilityUrl || DEFAULT_EXECUTOR_SETTINGS.utilityUrl;
    ui.pairingToken.value = query.pairingToken || DEFAULT_EXECUTOR_SETTINGS.pairingToken;
    ui.workspaceRoot.value = query.workspaceRoot || stored.workspaceRoot || DEFAULT_EXECUTOR_SETTINGS.workspaceRoot;
    return query.autoConnect;
  }

  function getEffectivePairingToken() {
    return ui.pairingToken.value.trim() || DEFAULT_EXECUTOR_SETTINGS.pairingToken;
  }

  function getBridgeState() {
    return {
      connected: Boolean(state.bridge),
      interactive: Boolean(state.bridge?.helloReceived),
      runtimeReady: state.runtimeReady,
      runtimeVersion: state.runtimeVersion,
      sessionId: state.bridge?.sessionId || "",
    };
  }

  function normalizeWindowsPath(value) {
    return String(value || "")
      .replace(/\//g, "\\")
      .replace(/\\{2,}/g, "\\");
  }

  function resolveWorkspacePath(...parts) {
    const base = normalizeWindowsPath(ui.workspaceRoot.value.trim());
    const cleaned = [base, ...parts]
      .map((part) => normalizeWindowsPath(part))
      .filter(Boolean);
    if (cleaned.length === 0) {
      return "";
    }
    return cleaned
      .map((part, index) => {
        if (index === 0) {
          return part.replace(/[\\]+$/, "");
        }
        return part.replace(/^[\\]+/, "").replace(/[\\]+$/, "");
      })
      .join("\\");
  }

  function downloadBytes(bytes, fileName, mimeType = "application/octet-stream") {
    const blob = new Blob([bytes], { type: mimeType });
    const url = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = url;
    link.download = fileName;
    link.click();
    URL.revokeObjectURL(url);
  }

  function setModuleDisplay(moduleId, patch) {
    const current = displayById.get(moduleId) || {
      badgeLabel: "module idle",
      badgeActive: false,
      metaText: "Модуль не активирован.",
    };
    const next = { ...current, ...patch };
    displayById.set(moduleId, next);

    if (moduleId === state.activeModuleId) {
      setBadge(ui.moduleBadge, next.badgeLabel, next.badgeActive);
      ui.moduleMeta.textContent = next.metaText;
    }
  }

  function makeLogger(scope) {
    return {
      info(message, detail = "") {
        appendLog(scope, message, detail);
      },
      error(message, detail = "") {
        appendLog(scope, `ERROR ${message}`, detail);
      },
    };
  }

  async function executeCommand(commandId, args = {}, timeoutMilliseconds = 60000) {
    if (!state.bridge) {
      throw new Error("Bridge is not connected.");
    }
    if (!state.runtimeReady) {
      throw new Error("Runtime is not loaded.");
    }
    return state.bridge.execute(EXECUTOR_PROFILE_ID, commandId, args, timeoutMilliseconds);
  }

  async function fileExists(path) {
    const execution = await executeCommand("system.file.exists", { path }, 15000);
    return Boolean(execution.result);
  }

  async function readFileBytes(path) {
    const execution = await executeCommand("system.file.read-bytes", { path }, 30000);
    const payload = Array.isArray(execution.result) ? execution.result : [];
    return new Uint8Array(payload);
  }

  function createModuleContext(module) {
    const moduleLogger = makeLogger(module.id);
    return {
      id: module.id,
      events,
      dsl: runtimeDsl,
      logger: moduleLogger,
      storage: {
        get(key, fallback = null) {
          const raw = window.localStorage.getItem(`${module.id}:${key}`);
          if (raw === null) {
            return fallback;
          }
          try {
            return JSON.parse(raw);
          } catch {
            return fallback;
          }
        },
        set(key, value) {
          window.localStorage.setItem(`${module.id}:${key}`, JSON.stringify(value));
        },
      },
      getBridgeState,
      getUtilityUrl: () => ui.utilityUrl.value.trim(),
      getPairingToken: () => getEffectivePairingToken(),
      getWorkspaceRoot: () => ui.workspaceRoot.value.trim(),
      resolveWorkspacePath,
      getProfileId: () => EXECUTOR_PROFILE_ID,
      executeCommand,
      fileExists,
      readFileBytes,
      downloadBytes,
      setModuleBadge(label, active = false) {
        setModuleDisplay(module.id, { badgeLabel: label, badgeActive: active });
      },
      setModuleMeta(text) {
        setModuleDisplay(module.id, { metaText: text });
      },
      requestRuntimeReload() {
        return loadRuntime();
      },
    };
  }

  function applyModuleHeader(module) {
    ui.moduleTitle.textContent = module.title;
    ui.moduleSubtitle.textContent = module.subtitle || "Модуль WebBridge.Utility.";
    const display = displayById.get(module.id) || {
      badgeLabel: "module idle",
      badgeActive: false,
      metaText: "Модуль не активирован.",
    };
    setBadge(ui.moduleBadge, display.badgeLabel, display.badgeActive);
    ui.moduleMeta.textContent = display.metaText;
  }

  function mountModule(module) {
    if (hostsById.has(module.id)) {
      return hostsById.get(module.id);
    }

    const host = document.createElement("section");
    host.className = "module-host";
    host.dataset.moduleId = module.id;
    ui.moduleHosts.append(host);
    hostsById.set(module.id, host);

    const context = createModuleContext(module);
    const handle = module.mount(host, context) || {};
    handlesById.set(module.id, handle);
    return host;
  }

  async function runLifecycle(target, name) {
    if (target && typeof target[name] === "function") {
      await target[name]();
    }
  }

  async function activateModule(moduleId) {
    const module = moduleById.get(moduleId);
    if (!module) {
      throw new Error(`Unknown module: ${moduleId}`);
    }

    const previousId = state.activeModuleId;
    if (previousId && previousId !== moduleId) {
      const previousHost = hostsById.get(previousId);
      const previousHandle = handlesById.get(previousId);
      if (previousHost) {
        previousHost.classList.remove("is-active");
      }
      await runLifecycle(previousHandle, "deactivate");
      await runLifecycle(moduleById.get(previousId), "deactivate");
    }

    const host = mountModule(module);
    state.activeModuleId = moduleId;

    for (const tabButton of ui.tabs.querySelectorAll(".rail__tab")) {
      tabButton.classList.toggle("is-active", tabButton.dataset.moduleId === moduleId);
    }
    for (const moduleHost of ui.moduleHosts.querySelectorAll(".module-host")) {
      moduleHost.classList.toggle("is-active", moduleHost.dataset.moduleId === moduleId);
    }

    host.classList.add("is-active");
    applyModuleHeader(module);
    emit("module-activated", { moduleId });

    const handle = handlesById.get(moduleId);
    await runLifecycle(module, "activate");
    await runLifecycle(handle, "activate");
  }

  function renderTabs() {
    ui.tabs.replaceChildren();
    for (const module of orderedModules) {
      displayById.set(module.id, {
        badgeLabel: "module idle",
        badgeActive: false,
        metaText: module.subtitle || "Модуль ещё не активирован.",
      });

      const button = document.createElement("button");
      button.type = "button";
      button.className = "rail__tab";
      button.dataset.moduleId = module.id;
      button.innerHTML = `<span>${module.tabLabel}</span><small>${module.tabDetail || module.id}</small>`;
      button.addEventListener("click", () => {
        activateModule(module.id).catch((error) => {
          appendLog("shell", "ERROR activate-module", String(error.message || error));
        });
      });
      ui.tabs.append(button);
    }
  }

  function collectRuntimeContribution() {
    const combined = {
      commands: {},
      allowedTypes: [],
      allowedProcesses: [],
    };

    const contributions = [createShellContribution()];
    for (const module of orderedModules) {
      if (typeof module.getRuntimeContribution === "function") {
        contributions.push(module.getRuntimeContribution(createModuleContext(module)) || {});
      }
    }

    for (const contribution of contributions) {
      const commands = ensureObject(contribution.commands);
      for (const [commandId, definition] of Object.entries(commands)) {
        combined.commands[commandId] = {
          CommandId: commandId,
          ...cloneJson(definition),
        };
      }
      combined.allowedTypes.push(...toArray(contribution.allowedTypes));
      combined.allowedProcesses.push(...toArray(contribution.allowedProcesses));
    }

    combined.allowedTypes = uniqueStrings(combined.allowedTypes);
    combined.allowedProcesses = uniqueStrings(combined.allowedProcesses);
    return combined;
  }

  function buildRuntimeSettings(effectiveSettings) {
    const base = cloneJson(ensureObject(effectiveSettings));
    const versions = ensureObject(getCaseInsensitive(base, "Versions"));
    const metadata = ensureObject(getCaseInsensitive(base, "Metadata"));
    const runtime = ensureObject(getCaseInsensitive(base, "Runtime"));
    const server = ensureObject(getCaseInsensitive(base, "Server"));
    const lifecycle = ensureObject(getCaseInsensitive(base, "Lifecycle"));
    const logging = ensureObject(getCaseInsensitive(base, "Logging"));
    const storage = ensureObject(getCaseInsensitive(base, "Storage"));
    const catalog = ensureObject(getCaseInsensitive(base, "Catalog"));
    const adapters = ensureObject(getCaseInsensitive(base, "Adapters"));
    const security = ensureObject(getCaseInsensitive(base, "Security"));
    const session = ensureObject(getCaseInsensitive(base, "Session"));
    const baseSystem = ensureObject(getCaseInsensitive(adapters, "System"));
    const baseCom = toArray(getCaseInsensitive(adapters, "Com"));

    const contribution = collectRuntimeContribution();
    const allowedOrigins = uniqueStrings([
      ...toArray(getCaseInsensitive(security, "AllowedOrigins")),
      window.location.origin,
    ]);
    const allowedTypes = uniqueStrings([
      ...toArray(getCaseInsensitive(baseSystem, "AllowedTypeNames")),
      ...contribution.allowedTypes,
    ]);
    const allowedProcesses = uniqueStrings([
      ...toArray(getCaseInsensitive(baseSystem, "AllowedProcessExecutables")),
      ...contribution.allowedProcesses,
    ]);

    return {
      Versions: {
        UtilityVersion: getCaseInsensitive(versions, "UtilityVersion") || "1.0.0",
        ConfigVersion: `kompas-pages-executor-${Date.now()}`,
        ConfigSchemaVersion: 2,
      },
      Metadata: {
        ProductName: getCaseInsensitive(metadata, "ProductName") || "KOMPAS Pages Executor",
        ProductCode: getCaseInsensitive(metadata, "ProductCode") || "kompas3d-utility.pages.executor",
        Author: getCaseInsensitive(metadata, "Author") || "Codex",
        Company: getCaseInsensitive(metadata, "Company") || null,
        Description: getCaseInsensitive(metadata, "Description")
          || "Static Pages UI Executor for KOMPAS utilities via WebBridge.Utility.",
        RepositoryUrl: getCaseInsensitive(metadata, "RepositoryUrl") || null,
      },
      Runtime: {
        EnvironmentName: getCaseInsensitive(runtime, "EnvironmentName") || "PagesExecutor",
        DevMode: Boolean(getCaseInsensitive(runtime, "DevMode")),
        NoBrowser: true,
      },
      Server: {
        ListenUrl: getCaseInsensitive(server, "ListenUrl") || ui.utilityUrl.value.trim(),
      },
      Ui: {
        Url: basePageUrl(),
        OpenMode: "Never",
        SessionWaitSeconds: Number(getCaseInsensitive(getCaseInsensitive(base, "Ui"), "SessionWaitSeconds") || 5),
      },
      Lifecycle: {
        ShutdownPolicy: getCaseInsensitive(lifecycle, "ShutdownPolicy") || "WhenIdle",
        IdleSeconds: Number(getCaseInsensitive(lifecycle, "IdleSeconds") || 180),
      },
      Logging: {
        Level: getCaseInsensitive(logging, "Level") || "Information",
        DebugMode: Boolean(getCaseInsensitive(logging, "DebugMode")),
        FilePath: getCaseInsensitive(logging, "FilePath") || null,
      },
      Storage: {
        ProfileDirectory: getCaseInsensitive(storage, "ProfileDirectory") || null,
        CacheDirectory: getCaseInsensitive(storage, "CacheDirectory") || null,
        DiagnosticsDirectory: getCaseInsensitive(storage, "DiagnosticsDirectory") || null,
      },
      Catalog: {
        Url: getCaseInsensitive(catalog, "Url") || null,
        Manifest: getCaseInsensitive(catalog, "Manifest") || null,
        Profiles: [
          {
            ProfileId: EXECUTOR_PROFILE_ID,
            ConfigSchemaVersion: 1,
            Description: "Browser-loaded runtime profile for tabbed KOMPAS Pages executor.",
            Checksum: `kompas-pages-executor-${Date.now()}`,
            Commands: contribution.commands,
          },
        ],
      },
      Adapters: {
        Com: cloneJson(baseCom),
        System: {
          ...cloneJson(baseSystem),
          AllowedTypeNames: allowedTypes,
          AllowedProcessExecutables: allowedProcesses,
        },
      },
      Security: {
        LoopbackOnly: getCaseInsensitive(security, "LoopbackOnly") !== false,
        PairingToken: ui.pairingToken.value.trim(),
        AllowedOrigins: allowedOrigins,
      },
      Session: {
        HeartbeatIntervalSeconds: Number(getCaseInsensitive(session, "HeartbeatIntervalSeconds") || 10),
        HeartbeatTimeoutSeconds: Number(getCaseInsensitive(session, "HeartbeatTimeoutSeconds") || 30),
        PresenceTimeoutSeconds: Number(getCaseInsensitive(session, "PresenceTimeoutSeconds") || 60),
        SuppressAutoOpenOnPresenceSessions:
          getCaseInsensitive(session, "SuppressAutoOpenOnPresenceSessions") !== false,
        SweepIntervalSeconds: Number(getCaseInsensitive(session, "SweepIntervalSeconds") || 2),
      },
    };
  }

  async function loadRuntime() {
    if (!state.bridge) {
      throw new Error("Bridge is not connected.");
    }

    setBadge(ui.runtimeBadge, "runtime load", false);
    ui.runtimeMeta.textContent = "Чтение effective config и загрузка runtime profile.";

    const effective = await state.bridge.fetchJson("/config/effective");
    if (!effective.response.ok) {
      throw new Error(formatHttpError("/config/effective", effective.response, effective.payload));
    }

    const effectiveResponse = unwrapEnvelope(effective.payload);
    const effectiveSettings = effectiveResponse?.settings || effectiveResponse?.Settings || {};
    const settings = buildRuntimeSettings(effectiveSettings);

    const applied = await state.bridge.fetchJson("/config/load", {
      method: "POST",
      body: JSON.stringify({
        settings,
        persist: false,
      }),
    });
    const update = unwrapEnvelope(applied.payload);
    if (!applied.response.ok || !update?.applied) {
      throw new Error(update?.message || formatHttpError("/config/load", applied.response, applied.payload));
    }

    state.runtimeReady = true;
    state.runtimeVersion = update?.version?.configVersion || settings.Versions.ConfigVersion;
    setBadge(ui.runtimeBadge, "runtime ready", true);
    ui.runtimeMeta.textContent = `${state.runtimeVersion} | profile=${EXECUTOR_PROFILE_ID}`;
    appendLog("shell", "runtime-loaded", state.runtimeVersion);
    emit("runtime-loaded", { settings, response: update });
    return update;
  }

  async function connectBridge() {
    if (state.connectPromise) {
      return state.connectPromise;
    }

    state.connectPromise = (async () => {
      persistSettings();

      if (state.bridge) {
        await disconnectBridge();
      }

      setBadge(ui.bridgeBadge, "bridge connect", false);
      setBadge(ui.runtimeBadge, "runtime idle", false);
      ui.bridgeMeta.textContent = "Регистрация UI-сессии...";
      ui.runtimeMeta.textContent = "Runtime ещё не загружен.";
      state.runtimeReady = false;
      ui.connectButton.disabled = true;

      const createBridge = (pairingToken) => new WebBridgeClient({
        baseUrl: ui.utilityUrl.value.trim(),
        pairingToken,
        clientName: DEFAULT_EXECUTOR_SETTINGS.clientName,
        uiVersion: DEFAULT_EXECUTOR_SETTINGS.uiVersion,
      });

      const requestedToken = getEffectivePairingToken();
      const retryTokens = uniqueStrings([
        requestedToken,
        DEFAULT_EXECUTOR_SETTINGS.pairingToken,
        "kompas-pages-local",
      ]);
      let bridge = createBridge(retryTokens[0]);
      let registration;

      try {
        registration = await bridge.connect();
      } catch (error) {
        const message = String(error?.message || error);
        if (!message.includes("invalid pairing token") || retryTokens.length < 2) {
          throw error;
        }
        let connected = false;

        for (const retryToken of retryTokens.slice(1)) {
          appendLog("shell", "bootstrap-token-retry", retryToken);
          ui.pairingToken.value = retryToken;
          bridge = createBridge(retryToken);
          try {
            registration = await bridge.connect();
            connected = true;
            break;
          } catch (retryError) {
            const retryMessage = String(retryError?.message || retryError);
            if (!retryMessage.includes("invalid pairing token")) {
              throw retryError;
            }
          }
        }

        if (!connected) {
          throw error;
        }
      }

      state.bridge = bridge;
      setBadge(ui.bridgeBadge, "bridge online", true);
      ui.bridgeMeta.textContent = `session=${registration.sessionId} | heartbeat=${registration.heartbeatIntervalSeconds}s`;
      appendLog("shell", "bridge-connected", registration.sessionId);
      emit("bridge-connected", { registration });
      await loadRuntime();
    })();

    try {
      await state.connectPromise;
    } finally {
      state.connectPromise = null;
      ui.connectButton.disabled = false;
    }
  }

  async function disconnectBridge() {
    if (!state.bridge) {
      return;
    }

    await state.bridge.disconnect();
    state.bridge = null;
    state.runtimeReady = false;
    state.runtimeVersion = "";
    setBadge(ui.bridgeBadge, "bridge offline", false);
    setBadge(ui.runtimeBadge, "runtime idle", false);
    ui.bridgeMeta.textContent = "Сессия не зарегистрирована.";
    ui.runtimeMeta.textContent = "Runtime ещё не загружен.";
    appendLog("shell", "bridge-disconnected");
    emit("bridge-disconnected");
  }

  function bindShellEvents() {
    ui.connectButton.addEventListener("click", () => {
      connectBridge().catch((error) => {
        setBadge(ui.bridgeBadge, "bridge error", false);
        ui.bridgeMeta.textContent = String(error.message || error);
        appendLog("shell", "ERROR bridge-connect", ui.bridgeMeta.textContent);
      });
    });

    ui.disconnectButton.addEventListener("click", () => {
      disconnectBridge().catch((error) => {
        appendLog("shell", "ERROR bridge-disconnect", String(error.message || error));
      });
    });

    ui.reloadRuntimeButton.addEventListener("click", () => {
      loadRuntime().catch((error) => {
        setBadge(ui.runtimeBadge, "runtime error", false);
        ui.runtimeMeta.textContent = String(error.message || error);
        appendLog("shell", "ERROR runtime-load", ui.runtimeMeta.textContent);
      });
    });

    ui.clearLogButton.addEventListener("click", () => {
      replaceLog("ready");
    });

    ui.utilityUrl.addEventListener("change", persistSettings);
    ui.workspaceRoot.addEventListener("change", persistSettings);

    window.addEventListener("beforeunload", () => {
      if (state.bridge) {
        state.bridge.disconnect("page-close").catch(() => {});
      }
    });
  }

  async function init() {
    replaceLog("ready");
    renderTabs();
    bindShellEvents();
    const autoConnect = loadSettings();
    await activateModule(state.activeModuleId);
    setBadge(ui.bridgeBadge, "bridge offline", false);
    setBadge(ui.runtimeBadge, "runtime idle", false);
    applyModuleHeader(moduleById.get(state.activeModuleId));

    if (autoConnect) {
      connectBridge().catch((error) => {
        setBadge(ui.bridgeBadge, "bridge error", false);
        ui.bridgeMeta.textContent = String(error.message || error);
        appendLog("shell", "ERROR autoconnect", ui.bridgeMeta.textContent);
      });
    }
  }

  return {
    init,
    events,
    executeCommand,
    connectBridge,
    disconnectBridge,
    loadRuntime,
    getBridgeState,
  };
}

export {
  EXECUTOR_PROFILE_ID,
  DEFAULT_EXECUTOR_SETTINGS,
  createExecutorShell,
  runtimeDsl,
};
