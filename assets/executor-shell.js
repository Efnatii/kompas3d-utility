const EXECUTOR_PROFILE_ID = "kompas-pages-executor";
const DEFAULT_UTILITY_URL = "http://127.0.0.1:38741";
const DEFAULT_PAIRING_TOKENS = ["kompas-pages-local", "replace-this-token"];
const UI_FINGERPRINT_SALT = "kompas-pages-ui:v1";
const CONNECT_BACKOFF_MS = [1000, 2000, 5000];
const RUNTIME_PROFILE_DESCRIPTION = "Browser-managed runtime overlay for remote-updatable KOMPAS Pages UI.";

const DEFAULT_EXECUTOR_SETTINGS = {
  utilityUrl: DEFAULT_UTILITY_URL,
  clientName: "kompas-pages-executor",
  uiVersion: "2.0.0",
  defaultPairingTokens: [...DEFAULT_PAIRING_TOKENS],
};

const KOMPAS_COM_ADAPTER = {
  AdapterName: "kompas",
  DisplayName: "KOMPAS",
  DispatcherName: "KOMPAS COM",
  DefaultProgIds: ["KOMPAS.Application.7", "KOMPAS.Application.5"],
  ComErrorCode: "kompas_com_error",
  InvokeErrorCode: "kompas_invoke_failed",
  UnavailableErrorCode: "adapter_unavailable",
  UnavailableMessage: "KOMPAS COM runtime is not available.",
  ReuseApplication: true,
  ReuseDocumentContext: true,
  BusyRetryMaxAttempts: 8,
  BusyRetryDelaysMs: [25, 50, 100, 150, 250, 400, 600, 800],
  CompactReportDefault: false,
  InteropAssemblies: [
    "C:\\Program Files\\ASCON\\KOMPAS-3D v24\\Libs\\PolynomLib\\Bin\\Client\\Interop.KompasAPI7.dll",
    "C:\\Program Files\\ASCON\\KOMPAS-3D v24\\Bin\\Interop.KompasAPI7.dll",
  ],
  Surfaces: [
    {
      Name: "IApplication",
      ClrTypeName: "KompasAPI7.IApplication",
      Iid: "6a2efaf7-8254-45a5-9dc8-2213f16af5d7",
      Aliases: ["application"],
    },
    {
      Name: "IKompasDocument2D",
      ClrTypeName: "KompasAPI7.IKompasDocument2D",
      Iid: "096e62b3-7184-4998-9925-74bb710d8d8e",
      Aliases: ["document2d"],
    },
    {
      Name: "IViewsAndLayersManager",
      ClrTypeName: "KompasAPI7.IViewsAndLayersManager",
      Iid: "a4737593-578b-4187-8cad-e1056eb5404b",
      Aliases: ["viewsManager"],
    },
    {
      Name: "IViews",
      ClrTypeName: "KompasAPI7.IViews",
      Iid: "9cd1b5e6-c1a2-4910-8d0c-97080b14aa3d",
      Aliases: ["views"],
    },
    {
      Name: "IView",
      ClrTypeName: "KompasAPI7.IView",
      Iid: "21a7ba87-1c8b-41b4-8247-cdd593546f37",
      Aliases: ["view"],
    },
    {
      Name: "ISymbols2DContainer",
      ClrTypeName: "KompasAPI7.ISymbols2DContainer",
      Iid: "f46b0086-17f2-4489-a5a7-0aa677610afd",
      Aliases: ["symbols2d"],
    },
    {
      Name: "IDrawingTables",
      ClrTypeName: "KompasAPI7.IDrawingTables",
      Iid: "df92dace-bdc6-4341-86da-3a9c8dcfdefe",
      Aliases: ["drawingTables"],
    },
    {
      Name: "IDrawingTable",
      ClrTypeName: "KompasAPI7.IDrawingTable",
      Iid: "9b421bda-0444-4a68-b69c-1c05d05c9d28",
      Aliases: ["drawingTable"],
    },
    {
      Name: "ITable",
      ClrTypeName: "KompasAPI7.ITable",
      Iid: "d3715420-645e-435b-bb25-8e35ac570718",
      Aliases: ["table"],
    },
    {
      Name: "ITableCell",
      ClrTypeName: "KompasAPI7.ITableCell",
      Iid: "cf9150ba-0e3a-46de-8973-332a00361474",
      Aliases: ["tableCell"],
    },
    {
      Name: "IText",
      ClrTypeName: "KompasAPI7.IText",
      Iid: "99b840fc-0150-4dad-bc0e-ad481baab8c2",
      Aliases: ["text"],
    },
  ],
  RealtimeAssignments: [
    { Member: "Visible", Value: true },
  ],
  VisibleApplicationAssignments: [
    { Member: "Visible", Value: true },
  ],
  ResultHints: [
    {
      Name: "application",
      MatchCurrentApplicationReference: true,
      RequiredAllMembers: ["Documents", "Visible"],
      Fields: [
        { Name: "connected", StaticValue: true },
        { Name: "progId", UseRuntimeProgId: true },
        { Name: "visible", MemberPaths: ["Visible"] },
        { Name: "version", MemberPaths: ["Version"] },
      ],
      CompactFields: [
        { Name: "progId", UseRuntimeProgId: true },
        { Name: "visible", MemberPaths: ["Visible"] },
      ],
    },
    {
      Name: "view",
      RequiredAllMembers: ["Layers", "Name"],
      Fields: [
        { Name: "name", MemberPaths: ["Name"] },
        { Name: "scale", MemberPaths: ["Scale"] },
        { Name: "angle", MemberPaths: ["Angle"] },
      ],
      CompactFields: [
        { Name: "name", MemberPaths: ["Name"] },
      ],
    },
    {
      Name: "document",
      RequiredAllMembers: ["Type"],
      RequiredAnyMembers: ["PathName", "Name"],
      Fields: [
        { Name: "name", MemberPaths: ["Name"] },
        { Name: "path", MemberPaths: ["PathName"] },
        { Name: "type", MemberPaths: ["Type"] },
      ],
      CompactFields: [
        { Name: "path", MemberPaths: ["PathName", "Name"] },
      ],
    },
  ],
};

function cloneJson(value) {
  return value === undefined ? undefined : JSON.parse(JSON.stringify(value));
}

function toArray(value) {
  return Array.isArray(value) ? value : [];
}

function ensureObject(value) {
  return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function uniqueStrings(values) {
  const seen = new Set();
  const result = [];
  for (const value of toArray(values)) {
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

function uniqueByKey(values, keySelector) {
  const result = [];
  const seen = new Set();
  for (const value of values) {
    const key = String(keySelector(value) || "").trim().toLowerCase();
    if (!key || seen.has(key)) {
      continue;
    }
    seen.add(key);
    result.push(value);
  }
  return result;
}

function normalizeStringPart(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function normalizeNumberPart(value) {
  const numeric = Number(value);
  return Number.isFinite(numeric) ? String(Math.trunc(numeric)) : "0";
}

function normalizeArrayForCompare(values) {
  return uniqueStrings(toArray(values)).map((value) => value.toLowerCase()).sort();
}

function formatStamp() {
  return new Date().toLocaleTimeString("ru-RU", { hour12: false });
}

function basePageUrl(locationLike = globalThis.location) {
  if (!locationLike?.origin || !locationLike?.pathname) {
    return "";
  }
  return `${locationLike.origin}${locationLike.pathname}`;
}

function normalizeWindowsPath(value) {
  return String(value || "")
    .replace(/\//g, "\\")
    .replace(/\\{2,}/g, "\\");
}

function dirname(filePath) {
  const value = normalizeWindowsPath(filePath);
  const index = value.lastIndexOf("\\");
  return index >= 0 ? value.slice(0, index) : "";
}

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

function isInvalidPairingTokenError(error) {
  return String(error?.message || error).toLowerCase().includes("invalid pairing token");
}

function stableSortJson(value) {
  if (Array.isArray(value)) {
    return value.map((entry) => stableSortJson(entry));
  }
  if (value && typeof value === "object") {
    return Object.keys(value)
      .sort((left, right) => left.localeCompare(right))
      .reduce((accumulator, key) => {
        accumulator[key] = stableSortJson(value[key]);
        return accumulator;
      }, {});
  }
  return value;
}

function stableStringify(value) {
  return JSON.stringify(stableSortJson(value));
}

function camelCaseKey(value) {
  const text = String(value || "");
  return text ? `${text.slice(0, 1).toLowerCase()}${text.slice(1)}` : "";
}

function readNamedValue(source, name) {
  const object = ensureObject(source);
  if (Object.prototype.hasOwnProperty.call(object, name)) {
    return object[name];
  }
  const camelName = camelCaseKey(name);
  if (camelName && Object.prototype.hasOwnProperty.call(object, camelName)) {
    return object[camelName];
  }
  return undefined;
}

function clearNamedValue(target, name) {
  if (!target || typeof target !== "object") {
    return;
  }
  delete target[name];
  const camelName = camelCaseKey(name);
  if (camelName) {
    delete target[camelName];
  }
}

function pickFirstDefined(...values) {
  for (const value of values) {
    if (value !== undefined) {
      return value;
    }
  }
  return undefined;
}

function caseFoldKeys(value) {
  if (Array.isArray(value)) {
    return value.map((entry) => caseFoldKeys(entry));
  }
  if (value && typeof value === "object") {
    return Object.keys(value).reduce((accumulator, key) => {
      accumulator[String(key).toLowerCase()] = caseFoldKeys(value[key]);
      return accumulator;
    }, {});
  }
  return value;
}

function buildCanonicalBlock(source, keys) {
  const block = {};
  for (const key of keys) {
    const value = readNamedValue(source, key);
    if (value !== undefined) {
      block[key] = cloneJson(value);
    }
  }
  return block;
}

function normalizeFlatSettings(source) {
  return {
    AgentVersion: pickFirstDefined(
      readNamedValue(source, "AgentVersion"),
      readNamedValue(source, "UtilityVersion"),
    ) || "",
    ConfigVersion: readNamedValue(source, "ConfigVersion") || "",
    ConfigSchemaVersion: readNamedValue(source, "ConfigSchemaVersion") || 0,
    EnvironmentName: readNamedValue(source, "EnvironmentName") || "",
    ListenUrl: readNamedValue(source, "ListenUrl") || "",
    ManifestUrl: readNamedValue(source, "ManifestUrl") || "",
    UiUrl: readNamedValue(source, "UiUrl") || "",
    Profiles: cloneJson(toArray(readNamedValue(source, "Profiles"))),
    ComAdapters: cloneJson(toArray(readNamedValue(source, "ComAdapters"))),
    SystemAdapter: buildCanonicalBlock(readNamedValue(source, "SystemAdapter"), [
      "RootAliases",
      "AllowedRoots",
      "Members",
      "AllowedTypeNames",
      "DeniedTypeNames",
      "AllowedNamespaces",
      "DeniedNamespaces",
      "DeniedInvocations",
      "AllowedProcessExecutables",
      "DeniedProcessExecutables",
      "AllowedUrlPrefixes",
      "DeniedUrlPrefixes",
      "AllowedRegistryHives",
      "DeniedRegistryHives",
    ]),
    Security: buildCanonicalBlock(readNamedValue(source, "Security"), [
      "LoopbackOnly",
      "PairingToken",
      "AllowedOrigins",
    ]),
  };
}

function normalizeLegacySettings(source) {
  const versions = readNamedValue(source, "Versions");
  const runtime = readNamedValue(source, "Runtime");
  const server = readNamedValue(source, "Server");
  const ui = readNamedValue(source, "Ui");
  const catalog = readNamedValue(source, "Catalog");
  const adapters = readNamedValue(source, "Adapters");

  return {
    AgentVersion: readNamedValue(versions, "UtilityVersion") || "",
    ConfigVersion: readNamedValue(versions, "ConfigVersion") || "",
    ConfigSchemaVersion: readNamedValue(versions, "ConfigSchemaVersion") || 0,
    EnvironmentName: readNamedValue(runtime, "EnvironmentName") || "",
    ListenUrl: readNamedValue(server, "ListenUrl") || "",
    UiUrl: readNamedValue(ui, "Url") || "",
    Profiles: cloneJson(toArray(readNamedValue(catalog, "Profiles"))),
    ComAdapters: cloneJson(toArray(readNamedValue(adapters, "Com"))),
    SystemAdapter: buildCanonicalBlock(readNamedValue(adapters, "System"), [
      "RootAliases",
      "AllowedRoots",
      "Members",
      "AllowedTypeNames",
      "DeniedTypeNames",
      "AllowedNamespaces",
      "DeniedNamespaces",
      "DeniedInvocations",
      "AllowedProcessExecutables",
      "DeniedProcessExecutables",
      "AllowedUrlPrefixes",
      "DeniedUrlPrefixes",
      "AllowedRegistryHives",
      "DeniedRegistryHives",
    ]),
    Security: buildCanonicalBlock(readNamedValue(source, "Security"), [
      "LoopbackOnly",
      "PairingToken",
      "AllowedOrigins",
    ]),
  };
}

function normalizeRuntimeSettings(settings) {
  const rawSettings = cloneJson(ensureObject(settings));
  const schemaVariant = readNamedValue(rawSettings, "Versions") || readNamedValue(rawSettings, "Catalog") || readNamedValue(rawSettings, "Adapters")
    ? "legacy-nested"
    : "flat";

  return {
    schemaVariant,
    rawSettings,
    normalizedSettings: schemaVariant === "legacy-nested"
      ? normalizeLegacySettings(rawSettings)
      : normalizeFlatSettings(rawSettings),
  };
}

function serializeFlatRuntimeSettings(rawSettings, desiredSettings) {
  const serialized = cloneJson(ensureObject(rawSettings));
  const systemAdapter = cloneJson(ensureObject(desiredSettings.SystemAdapter));
  const security = cloneJson(ensureObject(desiredSettings.Security));

  clearNamedValue(serialized, "ConfigVersion");
  serialized.ConfigVersion = desiredSettings.ConfigVersion;

  clearNamedValue(serialized, "UiUrl");
  serialized.UiUrl = desiredSettings.UiUrl;

  clearNamedValue(serialized, "Profiles");
  serialized.Profiles = cloneJson(toArray(desiredSettings.Profiles));

  clearNamedValue(serialized, "ComAdapters");
  serialized.ComAdapters = cloneJson(toArray(desiredSettings.ComAdapters));

  clearNamedValue(serialized, "SystemAdapter");
  serialized.SystemAdapter = systemAdapter;

  clearNamedValue(serialized, "Security");
  serialized.Security = security;

  return serialized;
}

function serializeLegacyRuntimeSettings(rawSettings, desiredSettings) {
  const versions = buildCanonicalBlock(readNamedValue(rawSettings, "Versions"), [
    "UtilityVersion",
    "ConfigVersion",
    "ConfigSchemaVersion",
  ]);
  versions.ConfigVersion = desiredSettings.ConfigVersion;

  const metadata = buildCanonicalBlock(readNamedValue(rawSettings, "Metadata"), [
    "ProductName",
    "ProductCode",
    "Author",
    "Company",
    "Description",
    "RepositoryUrl",
  ]);
  const runtime = buildCanonicalBlock(readNamedValue(rawSettings, "Runtime"), [
    "EnvironmentName",
    "DevMode",
    "NoBrowser",
  ]);
  const server = buildCanonicalBlock(readNamedValue(rawSettings, "Server"), [
    "ListenUrl",
  ]);
  const ui = buildCanonicalBlock(readNamedValue(rawSettings, "Ui"), [
    "Url",
    "OpenMode",
    "SessionWaitSeconds",
  ]);
  ui.Url = desiredSettings.UiUrl;

  const lifecycle = buildCanonicalBlock(readNamedValue(rawSettings, "Lifecycle"), [
    "ShutdownPolicy",
    "IdleSeconds",
  ]);
  const logging = buildCanonicalBlock(readNamedValue(rawSettings, "Logging"), [
    "Level",
    "DebugMode",
    "FilePath",
  ]);
  const storage = buildCanonicalBlock(readNamedValue(rawSettings, "Storage"), [
    "ProfileDirectory",
    "CacheDirectory",
    "DiagnosticsDirectory",
  ]);
  const catalog = buildCanonicalBlock(readNamedValue(rawSettings, "Catalog"), [
    "Url",
    "Manifest",
  ]);
  catalog.Profiles = cloneJson(toArray(desiredSettings.Profiles));

  const adapters = {
    Com: cloneJson(toArray(desiredSettings.ComAdapters)),
    System: cloneJson(ensureObject(desiredSettings.SystemAdapter)),
  };
  const security = {
    ...buildCanonicalBlock(readNamedValue(rawSettings, "Security"), [
      "LoopbackOnly",
      "AllowedOrigins",
      "PairingToken",
    ]),
    ...cloneJson(ensureObject(desiredSettings.Security)),
  };
  const session = buildCanonicalBlock(readNamedValue(rawSettings, "Session"), [
    "HeartbeatIntervalSeconds",
    "HeartbeatTimeoutSeconds",
    "PresenceTimeoutSeconds",
    "SuppressAutoOpenOnPresenceSessions",
    "SweepIntervalSeconds",
  ]);

  return {
    Versions: versions,
    Metadata: metadata,
    Runtime: runtime,
    Server: server,
    Ui: ui,
    Lifecycle: lifecycle,
    Logging: logging,
    Storage: storage,
    Catalog: catalog,
    Adapters: adapters,
    Security: security,
    Session: session,
  };
}

function serializeRuntimeLoadSettings({ schemaVariant, rawSettings, desiredSettings }) {
  return schemaVariant === "legacy-nested"
    ? serializeLegacyRuntimeSettings(rawSettings, desiredSettings)
    : serializeFlatRuntimeSettings(rawSettings, desiredSettings);
}

async function sha256Hex(value) {
  if (!globalThis.crypto?.subtle) {
    throw new Error("crypto.subtle is unavailable");
  }
  const encoded = new TextEncoder().encode(String(value || ""));
  const digest = await globalThis.crypto.subtle.digest("SHA-256", encoded);
  return Array.from(new Uint8Array(digest), (byte) => byte.toString(16).padStart(2, "0")).join("");
}

function buildUiFingerprintSource(runtime = {}) {
  const navigatorLike = runtime.navigator ?? globalThis.navigator ?? {};
  const screenLike = runtime.screen ?? globalThis.screen ?? {};
  const intlLike = runtime.Intl ?? globalThis.Intl;
  const timezone = runtime.timezone || (() => {
    try {
      return intlLike?.DateTimeFormat?.().resolvedOptions?.().timeZone || "";
    } catch {
      return "";
    }
  })();

  const language = navigatorLike.language || toArray(navigatorLike.languages)[0] || "";
  const screenSignature = [
    normalizeNumberPart(screenLike.width),
    normalizeNumberPart(screenLike.height),
    normalizeNumberPart(screenLike.colorDepth),
  ].join("x");

  return [
    normalizeStringPart(navigatorLike.platform),
    normalizeStringPart(language),
    normalizeStringPart(timezone),
    normalizeNumberPart(navigatorLike.hardwareConcurrency),
    normalizeNumberPart(navigatorLike.maxTouchPoints),
    screenSignature,
  ].join("|");
}

async function deriveUiPairingToken(runtime = {}) {
  const source = typeof runtime === "string" ? runtime : buildUiFingerprintSource(runtime);
  return sha256Hex(`${UI_FINGERPRINT_SALT}|${source}`);
}

function createTokenCandidates(derivedToken) {
  return uniqueStrings([derivedToken, ...DEFAULT_PAIRING_TOKENS]);
}

function findProfile(settings, profileId) {
  return toArray(settings?.Profiles).find((profile) => String(readNamedValue(profile, "ProfileId") || "").toLowerCase() === profileId.toLowerCase()) || null;
}

function findComAdapter(settings, adapterName) {
  return toArray(settings?.ComAdapters).find((adapter) => String(readNamedValue(adapter, "AdapterName") || "").toLowerCase() === adapterName.toLowerCase()) || null;
}

function mergeComAdapters(existingAdapters, requiredAdapters) {
  const byName = new Map();
  for (const adapter of toArray(existingAdapters)) {
    const clone = cloneJson(adapter);
    const key = String(readNamedValue(clone, "AdapterName") || "").trim().toLowerCase();
    if (!key) {
      continue;
    }
    byName.set(key, clone);
  }
  for (const adapter of toArray(requiredAdapters)) {
    const clone = cloneJson(adapter);
    const key = String(readNamedValue(clone, "AdapterName") || "").trim().toLowerCase();
    if (!key) {
      continue;
    }
    const current = byName.get(key);
    byName.set(key, mergeNamedConfig(current, clone));
  }
  return [...byName.values()];
}

function mergeNamedConfig(baseValue, overrideValue) {
  if (overrideValue === undefined) {
    return cloneJson(baseValue);
  }
  if (!isPlainObject(baseValue) || !isPlainObject(overrideValue)) {
    return cloneJson(overrideValue);
  }

  const result = cloneJson(baseValue);
  for (const [key, value] of Object.entries(overrideValue)) {
    if (value === undefined) {
      continue;
    }
    if (isPlainObject(value) && isPlainObject(result[key])) {
      result[key] = mergeNamedConfig(result[key], value);
      continue;
    }
    result[key] = cloneJson(value);
  }
  return result;
}

function sleep(milliseconds) {
  return new Promise((resolve) => window.setTimeout(resolve, milliseconds));
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
};

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
      "system.file.delete": runtimeDsl.command(
        "system",
        "type:System.IO.File",
        [
          runtimeDsl.step("call", "Delete", {
            args: [runtimeDsl.arg("path", "path")],
          }),
        ],
      ),
      "system.directory.ensure": runtimeDsl.command(
        "system",
        "type:System.IO.Directory",
        [
          runtimeDsl.step("call", "CreateDirectory", {
            args: [runtimeDsl.arg("path", "path")],
          }),
        ],
      ),
      "system.path.get-temp-path": runtimeDsl.command(
        "system",
        "type:System.IO.Path",
        [
          runtimeDsl.step("call", "GetTempPath", {
            args: [],
          }),
        ],
      ),
    },
    allowedTypes: [
      "System.IO.Directory",
      "System.IO.File",
      "System.IO.Path",
    ],
    comAdapters: [],
  };
}

function buildRuntimeProfileCommands(commands) {
  const result = {};
  for (const [commandId, definition] of Object.entries(ensureObject(commands))) {
    result[commandId] = {
      CommandId: commandId,
      ...cloneJson(definition),
    };
  }
  return result;
}

async function buildDesiredRuntimeSettings({
  effectiveSettings,
  pageOrigin,
  pageUrl,
  sessionToken,
  contribution,
}) {
  const base = cloneJson(ensureObject(effectiveSettings));
  const security = ensureObject(base.Security);
  const systemAdapter = ensureObject(base.SystemAdapter);
  const commands = buildRuntimeProfileCommands(contribution.commands);
  const profiles = toArray(base.Profiles)
    .filter((profile) => String(readNamedValue(profile, "ProfileId") || "").toLowerCase() !== EXECUTOR_PROFILE_ID.toLowerCase());
  const allowedOrigins = uniqueStrings([
    ...toArray(security.AllowedOrigins),
    pageOrigin,
  ]).sort((left, right) => left.localeCompare(right));
  const allowedTypes = uniqueStrings([
    ...toArray(systemAdapter.AllowedTypeNames),
    ...toArray(contribution.allowedTypes),
  ]).sort((left, right) => left.localeCompare(right));
  const comAdapters = mergeComAdapters(base.ComAdapters, contribution.comAdapters);

  const checksum = await sha256Hex(stableStringify({
    commands,
    comAdapters,
    profileId: EXECUTOR_PROFILE_ID,
    uiUrl: pageUrl,
    allowedOrigins,
    allowedTypes,
  }));

  const desiredProfile = {
    ProfileId: EXECUTOR_PROFILE_ID,
    ConfigSchemaVersion: 1,
    Description: RUNTIME_PROFILE_DESCRIPTION,
    Checksum: checksum,
    Commands: commands,
  };

  return {
    ...base,
    ConfigVersion: `kompas-pages-overlay-${checksum.slice(0, 12)}`,
    UiUrl: pageUrl,
    Profiles: [...profiles, desiredProfile],
    ComAdapters: comAdapters,
    SystemAdapter: {
      ...cloneJson(systemAdapter),
      AllowedTypeNames: allowedTypes,
    },
    Security: {
      ...cloneJson(security),
      AllowedOrigins: allowedOrigins,
      PairingToken: sessionToken,
    },
  };
}

function assessRuntimeOverlay({
  effectiveSettings,
  desiredSettings,
  contribution,
}) {
  const reasons = [];
  const effectiveOrigins = normalizeArrayForCompare(ensureObject(effectiveSettings?.Security).AllowedOrigins);
  const desiredOrigins = normalizeArrayForCompare(ensureObject(desiredSettings?.Security).AllowedOrigins);
  if (stableStringify(effectiveOrigins) !== stableStringify(desiredOrigins)) {
    reasons.push("allowed-origins");
  }

  if (String(effectiveSettings?.UiUrl || "") !== String(desiredSettings?.UiUrl || "")) {
    reasons.push("ui-url");
  }

  const effectiveAllowedTypes = normalizeArrayForCompare(ensureObject(effectiveSettings?.SystemAdapter).AllowedTypeNames);
  const requiredAllowedTypes = normalizeArrayForCompare(contribution.allowedTypes);
  if (!requiredAllowedTypes.every((typeName) => effectiveAllowedTypes.includes(typeName))) {
    reasons.push("system-types");
  }

  for (const adapter of toArray(contribution.comAdapters)) {
    const effectiveAdapter = findComAdapter(effectiveSettings, adapter.AdapterName);
    if (!effectiveAdapter) {
      reasons.push(`adapter:${adapter.AdapterName}`);
      continue;
    }
    if (stableStringify(caseFoldKeys(effectiveAdapter)) !== stableStringify(caseFoldKeys(adapter))) {
      reasons.push(`adapter:${adapter.AdapterName}`);
    }
  }

  const effectiveProfile = findProfile(effectiveSettings, EXECUTOR_PROFILE_ID);
  const desiredProfile = findProfile(desiredSettings, EXECUTOR_PROFILE_ID);
  if (!effectiveProfile) {
    reasons.push("profile-missing");
  } else if (!desiredProfile || String(readNamedValue(effectiveProfile, "Checksum") || "") !== String(readNamedValue(desiredProfile, "Checksum") || "")) {
    reasons.push("profile-checksum");
  }

  return {
    reloadRequired: reasons.length > 0,
    reasons,
  };
}

async function ensureRuntimeOverlay({
  bridge,
  pageOrigin,
  pageUrl,
  sessionToken,
  contribution,
}) {
  const effective = await bridge.fetchJson("/config/effective");
  if (!effective.response.ok) {
    throw new Error(formatHttpError("/config/effective", effective.response, effective.payload));
  }

  const effectiveResponse = unwrapEnvelope(effective.payload);
  const effectiveConfig = normalizeRuntimeSettings(effectiveResponse?.settings || effectiveResponse?.Settings || effectiveResponse);
  const effectiveSettings = effectiveConfig.normalizedSettings;
  const desiredSettings = await buildDesiredRuntimeSettings({
    effectiveSettings,
    pageOrigin,
    pageUrl,
    sessionToken,
    contribution,
  });
  const assessment = assessRuntimeOverlay({
    effectiveSettings,
    desiredSettings,
    contribution,
  });

  if (!assessment.reloadRequired) {
    return {
      applied: false,
      assessment,
      effectiveSettings,
      desiredSettings,
      runtimeVersion: desiredSettings.ConfigVersion,
    };
  }

  const applied = await bridge.fetchJson("/config/load", {
    method: "POST",
    body: JSON.stringify({
      settings: serializeRuntimeLoadSettings({
        schemaVariant: effectiveConfig.schemaVariant,
        rawSettings: effectiveConfig.rawSettings,
        desiredSettings,
      }),
      persist: false,
    }),
  });
  const update = unwrapEnvelope(applied.payload);
  if (!applied.response.ok || !update?.applied) {
    throw new Error(update?.message || formatHttpError("/config/load", applied.response, applied.payload));
  }

  return {
    applied: true,
    assessment,
    effectiveSettings,
    desiredSettings,
    runtimeVersion: update?.version?.configVersion || desiredSettings.ConfigVersion,
  };
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
    this.manualClose = false;
    this.closePromise = Promise.resolve(null);
    this.closeResolver = null;
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
    this.manualClose = true;
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
    } else if (this.closeResolver) {
      this.closeResolver(null);
      this.closeResolver = null;
    }
  }

  waitForClose() {
    return this.closePromise;
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

  async executeBatch(profileId, commands, options = {}) {
    const requestCommands = toArray(commands).map((command) => ({
      profileId,
      commandId: command.commandId,
      arguments: ensureObject(command.arguments),
      reportVerbosity: command.reportVerbosity || options.reportVerbosity || "Compact",
      sharedContextId: command.sharedContextId || null,
      timeoutMilliseconds: command.timeoutMilliseconds || null,
    }));
    const { response, payload } = await this.fetchJson("/commands/execute-batch", {
      method: "POST",
      body: JSON.stringify({
        commands: requestCommands,
        sharedContextId: options.sharedContextId || null,
        reportVerbosity: options.reportVerbosity || "Compact",
        stopOnError: options.stopOnError !== false,
      }),
    });
    const batch = unwrapEnvelope(payload);
    if (!response.ok || !batch?.success) {
      const firstError = toArray(batch?.results).find((result) => !result?.success)?.error;
      const errorMessage = firstError?.message
        || firstError?.code
        || extractApiError(payload)?.message
        || extractApiError(payload)?.code
        || response.status;
      throw new Error(`execute-batch failed: ${errorMessage}`);
    }
    return batch;
  }

  async openSocket(wsUrl) {
    this.stopHeartbeat();
    this.helloReceived = false;
    this.manualClose = false;
    this.socket = new WebSocket(wsUrl);
    this.closePromise = new Promise((resolve) => {
      this.closeResolver = resolve;
    });

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

    this.socket.addEventListener("close", (event) => {
      this.stopHeartbeat();
      this.socket = null;
      if (this.closeResolver) {
        const resolve = this.closeResolver;
        this.closeResolver = null;
        resolve(this.manualClose ? null : new Error(`ws closed ${event.code} ${event.reason || "no-reason"}`));
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

function createExecutorShell({ modules }) {
  if (!Array.isArray(modules) || modules.length === 0) {
    throw new Error("Executor shell requires at least one module.");
  }

  const ui = {
    tabs: document.getElementById("module-tabs"),
    bridgeBadge: document.getElementById("bridge-badge"),
    runtimeBadge: document.getElementById("runtime-badge"),
    moduleBadge: document.getElementById("module-badge"),
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
  const contextById = new Map();
  const events = new EventTarget();

  const state = {
    bridge: null,
    runtimeReady: false,
    runtimeVersion: "",
    activeModuleId: orderedModules[0].id,
    sessionToken: "",
    supervisorPromise: null,
    stopRequested: false,
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

  function getBridgeState() {
    return {
      connected: Boolean(state.bridge),
      interactive: Boolean(state.bridge?.helloReceived),
      runtimeReady: state.runtimeReady,
      runtimeVersion: state.runtimeVersion,
      sessionId: state.bridge?.sessionId || "",
    };
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

  async function executeBatchCommand(commands, options = {}) {
    if (!state.bridge) {
      throw new Error("Bridge is not connected.");
    }
    if (!state.runtimeReady) {
      throw new Error("Runtime is not loaded.");
    }
    return state.bridge.executeBatch(EXECUTOR_PROFILE_ID, commands, options);
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

  async function deleteFile(path) {
    await executeCommand("system.file.delete", { path }, 15000);
  }

  async function ensureDirectory(path) {
    await executeCommand("system.directory.ensure", { path }, 15000);
  }

  async function getTempPath() {
    const execution = await executeCommand("system.path.get-temp-path", {}, 15000);
    return String(execution.result || "");
  }

  function createModuleContext(module) {
    if (contextById.has(module.id)) {
      return contextById.get(module.id);
    }

    const context = {
      id: module.id,
      events,
      dsl: runtimeDsl,
      logger: makeLogger(module.id),
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
      getProfileId: () => EXECUTOR_PROFILE_ID,
      executeCommand,
      executeBatchCommand,
      fileExists,
      readFileBytes,
      deleteFile,
      ensureDirectory,
      getTempPath,
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

    contextById.set(module.id, context);
    return context;
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
      comAdapters: [],
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
        combined.commands[commandId] = cloneJson(definition);
      }
      combined.allowedTypes.push(...toArray(contribution.allowedTypes));
      combined.comAdapters.push(...toArray(contribution.comAdapters));
    }

    combined.allowedTypes = uniqueStrings(combined.allowedTypes);
    combined.comAdapters = mergeComAdapters([], combined.comAdapters);
    return combined;
  }

  function setBridgeWaiting(message) {
    setBadge(ui.bridgeBadge, "bridge wait", false);
    setBadge(ui.runtimeBadge, "runtime idle", false);
    ui.bridgeMeta.textContent = message;
    ui.runtimeMeta.textContent = "Runtime overlay ещё не активен.";
    state.runtimeReady = false;
    state.runtimeVersion = "";
  }

  function setBridgeOnline(registration) {
    setBadge(ui.bridgeBadge, "bridge online", true);
    ui.bridgeMeta.textContent = `session=${registration.sessionId} | heartbeat=${registration.heartbeatIntervalSeconds}s`;
  }

  function setRuntimeReady(runtime) {
    state.runtimeReady = true;
    state.runtimeVersion = runtime.runtimeVersion;
    setBadge(ui.runtimeBadge, "runtime ready", true);
    ui.runtimeMeta.textContent = `${runtime.runtimeVersion} | profile=${EXECUTOR_PROFILE_ID}${runtime.applied ? " | overlay synced" : " | overlay ok"}`;
  }

  function resetBridgeState(message) {
    state.bridge = null;
    state.runtimeReady = false;
    state.runtimeVersion = "";
    state.sessionToken = "";
    setBadge(ui.bridgeBadge, "bridge retry", false);
    setBadge(ui.runtimeBadge, "runtime idle", false);
    ui.bridgeMeta.textContent = message;
    ui.runtimeMeta.textContent = "Runtime overlay ещё не активен.";
    emit("bridge-disconnected", { message });
  }

  async function loadRuntime() {
    if (!state.bridge) {
      throw new Error("Bridge is not connected.");
    }

    setBadge(ui.runtimeBadge, "runtime sync", false);
    ui.runtimeMeta.textContent = "Проверка effective config и runtime overlay.";
    const contribution = collectRuntimeContribution();
    const runtime = await ensureRuntimeOverlay({
      bridge: state.bridge,
      pageOrigin: window.location.origin,
      pageUrl: basePageUrl(),
      sessionToken: state.sessionToken,
      contribution,
    });
    setRuntimeReady(runtime);
    appendLog("shell", "runtime-ready", runtime.applied ? runtime.assessment.reasons.join(",") : runtime.runtimeVersion);
    emit("runtime-loaded", runtime);
    return runtime;
  }

  async function resolveTokenCandidates() {
    let derivedToken = "";
    try {
      derivedToken = await deriveUiPairingToken();
      appendLog("shell", "token-derived", derivedToken.slice(0, 12));
    } catch (error) {
      appendLog("shell", "token-derive-skipped", String(error.message || error));
    }
    return createTokenCandidates(derivedToken);
  }

  async function connectOnce() {
    setBridgeWaiting(`Автоподключение к ${DEFAULT_EXECUTOR_SETTINGS.utilityUrl}`);
    const tokens = await resolveTokenCandidates();
    if (tokens.length === 0) {
      throw new Error("No pairing tokens available.");
    }

    let lastError = null;
    for (const token of tokens) {
      const bridge = new WebBridgeClient({
        baseUrl: DEFAULT_EXECUTOR_SETTINGS.utilityUrl,
        pairingToken: token,
        clientName: DEFAULT_EXECUTOR_SETTINGS.clientName,
        uiVersion: DEFAULT_EXECUTOR_SETTINGS.uiVersion,
      });

      try {
        const registration = await bridge.connect();
        state.bridge = bridge;
        state.sessionToken = token;
        setBridgeOnline(registration);
        appendLog("shell", "bridge-connected", `${registration.sessionId} token=${token.slice(0, 12)}`);
        emit("bridge-connected", { registration, pairingToken: token });
        const runtime = await loadRuntime();
        return { bridge, registration, runtime };
      } catch (error) {
        lastError = error;
        try {
          await bridge.disconnect("connect-failed");
        } catch {
          // Ignore cleanup failures.
        }
        if (!isInvalidPairingTokenError(error)) {
          throw error;
        }
        appendLog("shell", "token-retry", token.slice(0, 12));
      }
    }

    throw lastError || new Error("Bridge connection failed.");
  }

  async function superviseConnection() {
    if (state.supervisorPromise) {
      return state.supervisorPromise;
    }

    state.stopRequested = false;
    state.supervisorPromise = (async () => {
      let attempt = 0;
      while (!state.stopRequested) {
        try {
          const result = await connectOnce();
          attempt = 0;
          const closeError = await result.bridge.waitForClose();
          if (state.stopRequested) {
            break;
          }
          throw closeError || new Error("Bridge socket closed.");
        } catch (error) {
          if (state.stopRequested) {
            break;
          }
          const waitMs = CONNECT_BACKOFF_MS[Math.min(attempt, CONNECT_BACKOFF_MS.length - 1)];
          attempt += 1;
          const message = String(error?.message || error);
          resetBridgeState(`${message} | retry ${Math.round(waitMs / 1000)}s`);
          appendLog("shell", "reconnect-wait", `${message} | ${waitMs}ms`);
          await sleep(waitMs);
        }
      }
    })();

    try {
      await state.supervisorPromise;
    } finally {
      state.supervisorPromise = null;
    }
  }

  function bindShellEvents() {
    ui.clearLogButton.addEventListener("click", () => {
      replaceLog("ready");
    });

    window.addEventListener("beforeunload", () => {
      state.stopRequested = true;
      if (state.bridge) {
        state.bridge.disconnect("page-close").catch(() => {});
      }
    });
  }

  async function init() {
    replaceLog("ready");
    renderTabs();
    bindShellEvents();
    await activateModule(state.activeModuleId);
    setBridgeWaiting(`Автоподключение к ${DEFAULT_EXECUTOR_SETTINGS.utilityUrl}`);
    applyModuleHeader(moduleById.get(state.activeModuleId));
    superviseConnection().catch((error) => {
      const message = String(error.message || error);
      resetBridgeState(message);
      appendLog("shell", "ERROR supervisor", message);
    });
  }

  return {
    init,
    events,
    executeCommand,
    executeBatchCommand,
    connectBridge: superviseConnection,
    loadRuntime,
    getBridgeState,
  };
}

export {
  EXECUTOR_PROFILE_ID,
  DEFAULT_EXECUTOR_SETTINGS,
  KOMPAS_COM_ADAPTER,
  assessRuntimeOverlay,
  buildDesiredRuntimeSettings,
  buildUiFingerprintSource,
  createExecutorShell,
  createTokenCandidates,
  deriveUiPairingToken,
  ensureRuntimeOverlay,
  normalizeRuntimeSettings,
  runtimeDsl,
  serializeRuntimeLoadSettings,
  stableStringify,
};
