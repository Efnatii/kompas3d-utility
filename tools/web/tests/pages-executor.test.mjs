import test from "node:test";
import assert from "node:assert/strict";

import {
  KOMPAS_COM_ADAPTER,
  buildDesiredRuntimeSettings,
  buildUiFingerprintSource,
  createTokenCandidates,
  deriveUiPairingToken,
  ensureRuntimeOverlay,
  normalizeRuntimeSettings,
  serializeRuntimeLoadSettings,
} from "../../../assets/executor-shell.js";
import {
  buildAutoOutputPath,
  createCellWriteBatches,
  reconcileLinkedLayout,
} from "../../../assets/modules/xlsx-to-kompas-tbl.js";

test("ui derived token is deterministic for a mocked browser fingerprint", async () => {
  const fingerprint = buildUiFingerprintSource({
    navigator: {
      platform: "Win32",
      language: "ru-RU",
      hardwareConcurrency: 16,
      maxTouchPoints: 0,
    },
    screen: {
      width: 1920,
      height: 1080,
      colorDepth: 24,
    },
    timezone: "Europe/Moscow",
  });

  assert.equal(fingerprint, "win32|ru-ru|europe/moscow|16|0|1920x1080x24");
  assert.equal(
    await deriveUiPairingToken(fingerprint),
    "a44d70b524ace1ee7920bec049a0e5e58fefbb0181542abb2725b9de83680623",
  );
});

test("token candidates preserve derived-first fallback order", () => {
  assert.deepEqual(
    createTokenCandidates("derived-token"),
    ["derived-token", "kompas-pages-local", "replace-this-token"],
  );
  assert.deepEqual(
    createTokenCandidates("kompas-pages-local"),
    ["kompas-pages-local", "replace-this-token"],
  );
});

test("runtime overlay loads when effective config is missing required profile", async () => {
  const contribution = {
    commands: {
      "demo.command": {
        Adapter: "system",
        Invoke: {
          Root: "type:System.IO.Path",
          Chain: [],
        },
      },
    },
    allowedTypes: ["System.IO.Path"],
    comAdapters: [],
  };
  const effectiveSettings = {
    AgentVersion: "1.0.0",
    ConfigVersion: "bootstrap",
    ConfigSchemaVersion: 1,
    UiUrl: "https://old.example/index.html",
    Profiles: [],
    ComAdapters: [],
    SystemAdapter: {
      AllowedTypeNames: [],
    },
    Security: {
      AllowedOrigins: ["https://old.example"],
    },
  };

  const calls = [];
  const bridge = {
    async fetchJson(pathname, init = {}) {
      calls.push({ pathname, init });
      if (pathname === "/config/effective") {
        return {
          response: { ok: true, status: 200 },
          payload: { payload: { settings: effectiveSettings } },
        };
      }
      if (pathname === "/config/load") {
        return {
          response: { ok: true, status: 200 },
          payload: { payload: { applied: true, version: { configVersion: "overlay-loaded" } } },
        };
      }
      throw new Error(`Unexpected path: ${pathname}`);
    },
  };

  const runtime = await ensureRuntimeOverlay({
    bridge,
    pageOrigin: "https://ui.example",
    pageUrl: "https://ui.example/index.html",
    sessionToken: "kompas-pages-local",
    contribution,
  });

  assert.equal(runtime.applied, true);
  assert.equal(runtime.runtimeVersion, "overlay-loaded");
  assert.equal(calls.length, 2);
  assert.equal(calls[1].pathname, "/config/load");

  const request = JSON.parse(calls[1].init.body);
  assert.equal(request.persist, false);
  assert.equal(request.settings.UiUrl, "https://ui.example/index.html");
  assert.deepEqual(
    request.settings.Security.AllowedOrigins,
    ["https://old.example", "https://ui.example"],
  );
});

test("runtime overlay skips /config/load when effective config already matches desired runtime", async () => {
  const contribution = {
    commands: {
      "demo.command": {
        Adapter: "system",
        Invoke: {
          Root: "type:System.IO.Path",
          Chain: [],
        },
      },
    },
    allowedTypes: ["System.IO.Path"],
    comAdapters: [],
  };
  const desiredSettings = await buildDesiredRuntimeSettings({
    effectiveSettings: {
      AgentVersion: "1.0.0",
      ConfigVersion: "bootstrap",
      ConfigSchemaVersion: 1,
      UiUrl: "https://ui.example/index.html",
      Profiles: [],
      ComAdapters: [],
      SystemAdapter: {
        AllowedTypeNames: [],
      },
      Security: {
        AllowedOrigins: ["https://ui.example"],
      },
    },
    pageOrigin: "https://ui.example",
    pageUrl: "https://ui.example/index.html",
    sessionToken: "kompas-pages-local",
    contribution,
  });

  const calls = [];
  const bridge = {
    async fetchJson(pathname) {
      calls.push(pathname);
      if (pathname === "/config/effective") {
        return {
          response: { ok: true, status: 200 },
          payload: { payload: { settings: desiredSettings } },
        };
      }
      throw new Error(`Unexpected path: ${pathname}`);
    },
  };

  const runtime = await ensureRuntimeOverlay({
    bridge,
    pageOrigin: "https://ui.example",
    pageUrl: "https://ui.example/index.html",
    sessionToken: "kompas-pages-local",
    contribution,
  });

  assert.equal(runtime.applied, false);
  assert.deepEqual(calls, ["/config/effective"]);
});

test("runtime overlay loads when config version or pairing token is stale even if runtime structure matches", async () => {
  const contribution = {
    commands: {
      "demo.command": {
        Adapter: "system",
        Invoke: {
          Root: "type:System.IO.Path",
          Chain: [],
        },
      },
    },
    allowedTypes: ["System.IO.Path"],
    comAdapters: [],
  };

  const desiredSettings = await buildDesiredRuntimeSettings({
    effectiveSettings: {
      AgentVersion: "1.0.0",
      ConfigVersion: "bootstrap",
      ConfigSchemaVersion: 1,
      UiUrl: "https://ui.example/index.html",
      Profiles: [],
      ComAdapters: [],
      SystemAdapter: {
        AllowedTypeNames: [],
      },
      Security: {
        AllowedOrigins: ["https://ui.example"],
        PairingToken: "stale-token",
      },
    },
    pageOrigin: "https://ui.example",
    pageUrl: "https://ui.example/index.html",
    sessionToken: "kompas-pages-local",
    contribution,
  });

  const staleSettings = {
    ...desiredSettings,
    ConfigVersion: "bootstrap",
    Security: {
      ...desiredSettings.Security,
      PairingToken: "stale-token",
    },
  };

  const calls = [];
  const bridge = {
    async fetchJson(pathname, init = {}) {
      calls.push({ pathname, init });
      if (pathname === "/config/effective") {
        return {
          response: { ok: true, status: 200 },
          payload: { payload: { settings: staleSettings } },
        };
      }
      if (pathname === "/config/load") {
        return {
          response: { ok: true, status: 200 },
          payload: { payload: { applied: true, version: { configVersion: desiredSettings.ConfigVersion } } },
        };
      }
      throw new Error(`Unexpected path: ${pathname}`);
    },
  };

  const runtime = await ensureRuntimeOverlay({
    bridge,
    pageOrigin: "https://ui.example",
    pageUrl: "https://ui.example/index.html",
    sessionToken: "kompas-pages-local",
    contribution,
  });

  assert.equal(runtime.applied, true);
  assert.deepEqual(runtime.assessment.reasons, ["config-version", "pairing-token"]);
  assert.equal(calls.length, 2);

  const request = JSON.parse(calls[1].init.body);
  assert.equal(request.settings.ConfigVersion, desiredSettings.ConfigVersion);
  assert.equal(request.settings.Security.PairingToken, "kompas-pages-local");
});

test("runtime overlay keeps existing kompas adapter fields while injecting cast surfaces", async () => {
  const contribution = {
    commands: {},
    allowedTypes: [],
    comAdapters: [KOMPAS_COM_ADAPTER],
  };

  const desiredSettings = await buildDesiredRuntimeSettings({
    effectiveSettings: {
      AgentVersion: "1.0.0",
      ConfigVersion: "bootstrap",
      ConfigSchemaVersion: 1,
      UiUrl: "https://ui.example/index.html",
      Profiles: [],
      ComAdapters: [
        {
          AdapterName: "kompas",
          DisplayName: "KOMPAS",
          InteropAssemblies: ["C:\\Program Files\\ASCON\\KOMPAS-3D v24\\Bin\\Interop.KompasAPI7.dll"],
          ReuseApplication: true,
        },
      ],
      SystemAdapter: {
        AllowedTypeNames: [],
      },
      Security: {
        AllowedOrigins: ["https://ui.example"],
      },
    },
    pageOrigin: "https://ui.example",
    pageUrl: "https://ui.example/index.html",
    sessionToken: "kompas-pages-local",
    contribution,
  });

  const kompasAdapter = desiredSettings.ComAdapters.find((adapter) => adapter.AdapterName === "kompas");
  assert(kompasAdapter);
  assert.deepEqual(
    kompasAdapter.InteropAssemblies,
    [
      "C:\\Program Files\\ASCON\\KOMPAS-3D v24\\Libs\\PolynomLib\\Bin\\Client\\Interop.KompasAPI7.dll",
      "C:\\Program Files\\ASCON\\KOMPAS-3D v24\\Bin\\Interop.KompasAPI7.dll",
    ],
  );
  assert.equal(kompasAdapter.ReuseApplication, true);
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ISymbols2DContainer"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ITable"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "IText"));
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ISymbols2DContainer")?.Iid,
    "f46b0086-17f2-4489-a5a7-0aa677610afd",
  );
});

test("legacy nested effective config is normalized and posted back as nested /config/load payload", async () => {
  const contribution = {
    commands: {
      "demo.command": {
        Adapter: "system",
        Invoke: {
          Root: "type:System.IO.Path",
          Chain: [],
        },
      },
    },
    allowedTypes: ["System.IO.Path"],
    comAdapters: [],
  };
  const legacySettings = {
    versions: {
      utilityVersion: "1.0.0",
      configVersion: "legacy-bootstrap",
      configSchemaVersion: 2,
    },
    runtime: {
      environmentName: "PagesE2ELegacyDevelopment",
      devMode: true,
      noBrowser: true,
    },
    server: {
      listenUrl: "http://127.0.0.1:38741",
    },
    ui: {
      url: "https://old.example/index.html",
      openMode: "Never",
      sessionWaitSeconds: 5,
    },
    catalog: {
      profiles: [],
    },
    adapters: {
      com: [],
      system: {
        allowedTypeNames: [],
      },
    },
    security: {
      allowedOrigins: ["https://old.example"],
      pairingToken: "***redacted***",
      loopbackOnly: true,
    },
    session: {
      heartbeatIntervalSeconds: 10,
      heartbeatTimeoutSeconds: 30,
      presenceTimeoutSeconds: 60,
      suppressAutoOpenOnPresenceSessions: true,
      sweepIntervalSeconds: 2,
    },
  };

  const normalized = normalizeRuntimeSettings(legacySettings);
  assert.equal(normalized.schemaVariant, "legacy-nested");
  assert.deepEqual(normalized.normalizedSettings.Security.AllowedOrigins, ["https://old.example"]);
  assert.deepEqual(normalized.normalizedSettings.SystemAdapter.AllowedTypeNames, []);

  const calls = [];
  const bridge = {
    async fetchJson(pathname, init = {}) {
      calls.push({ pathname, init });
      if (pathname === "/config/effective") {
        return {
          response: { ok: true, status: 200 },
          payload: { payload: { settings: legacySettings } },
        };
      }
      if (pathname === "/config/load") {
        return {
          response: { ok: true, status: 200 },
          payload: { payload: { applied: true, version: { configVersion: "legacy-overlay-loaded" } } },
        };
      }
      throw new Error(`Unexpected path: ${pathname}`);
    },
  };

  const runtime = await ensureRuntimeOverlay({
    bridge,
    pageOrigin: "https://ui.example",
    pageUrl: "https://ui.example/index.html",
    sessionToken: "replace-this-token",
    contribution,
  });

  assert.equal(runtime.applied, true);
  assert.equal(runtime.runtimeVersion, "legacy-overlay-loaded");
  assert.equal(calls.length, 2);

  const request = JSON.parse(calls[1].init.body);
  assert.equal(request.persist, false);
  assert.equal(request.settings.ConfigVersion, undefined);
  assert.equal(request.settings.UiUrl, undefined);
  assert.equal(request.settings.Versions.ConfigVersion.startsWith("kompas-pages-overlay-"), true);
  assert.equal(request.settings.Ui.Url, "https://ui.example/index.html");
  assert.deepEqual(
    request.settings.Security.AllowedOrigins,
    ["https://old.example", "https://ui.example"],
  );
  assert.deepEqual(
    request.settings.Adapters.System.AllowedTypeNames,
    ["System.IO.Path"],
  );
  assert.equal(request.settings.Catalog.Profiles[0].ProfileId, "kompas-pages-executor");

  const serialized = serializeRuntimeLoadSettings({
    schemaVariant: normalized.schemaVariant,
    rawSettings: normalized.rawSettings,
    desiredSettings: await buildDesiredRuntimeSettings({
      effectiveSettings: normalized.normalizedSettings,
      pageOrigin: "https://ui.example",
      pageUrl: "https://ui.example/index.html",
      sessionToken: "replace-this-token",
      contribution,
    }),
  });
  assert.equal(serialized.Versions.ConfigVersion.startsWith("kompas-pages-overlay-"), true);
  assert.equal(serialized.Ui.Url, "https://ui.example/index.html");
});

test("linked layout keeps table and cell dimensions in sync without a mode toggle", () => {
  assert.deepEqual(
    reconcileLinkedLayout({
      tableWidthMm: 260,
      tableHeightMm: 24,
      cellWidthMm: 1,
      cellHeightMm: 1,
    }, 8, 13, "table"),
    {
      tableWidthMm: 260,
      tableHeightMm: 24,
      cellWidthMm: 20,
      cellHeightMm: 3,
    },
  );

  assert.deepEqual(
    reconcileLinkedLayout({
      tableWidthMm: 1,
      tableHeightMm: 1,
      cellWidthMm: 15,
      cellHeightMm: 7.5,
    }, 8, 13, "cell"),
    {
      tableWidthMm: 195,
      tableHeightMm: 60,
      cellWidthMm: 15,
      cellHeightMm: 7.5,
    },
  );
});

test("export batching filters empty cells and follows output path rules", () => {
  const matrix = [
    ["A1", "", "A3"],
    ["", "", ""],
    ["C1", "C2", ""],
  ];
  const batches = createCellWriteBatches(matrix, 2);

  assert.equal(batches.length, 2);
  assert.deepEqual(
    batches.map((batch) => batch.map((command) => command.arguments)),
    [
      [
        { rowIndex: 0, columnIndex: 0, value: "A1" },
        { rowIndex: 0, columnIndex: 2, value: "A3" },
      ],
      [
        { rowIndex: 2, columnIndex: 0, value: "C1" },
        { rowIndex: 2, columnIndex: 1, value: "C2" },
      ],
    ],
  );

  assert.equal(
    buildAutoOutputPath({
      documentPath: "C:\\Drawings\\Sample.cdw",
      fileName: "table_M2.xlsx",
      tempPath: "C:\\Temp",
    }),
    "C:\\Drawings\\table_M2.tbl",
  );
  assert.equal(
    buildAutoOutputPath({
      documentPath: "",
      fileName: "table_M2.xlsx",
      tempPath: "C:\\Temp",
    }),
    "C:\\Temp\\kompas-pages\\table_M2.tbl",
  );
});
