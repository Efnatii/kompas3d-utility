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
  autoFitCellMatrixToLayout,
  buildWorkbookStyleContext,
  buildAutoOutputPath,
  buildInlineTempOutputPath,
  createCellTransferPayload,
  createCellWriteBatches,
  createFormattedCellWriteBatches,
  createXlsxToKompasTblModule,
  parseWorksheetCellMeta,
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
      "C:\\Program Files\\ASCON\\KOMPAS-3D v24\\Libs\\PolynomLib\\Bin\\Client\\Interop.Kompas6API5.dll",
      "C:\\Program Files\\ASCON\\KOMPAS-3D v24\\Bin\\Interop.KompasAPI7.dll",
    ],
  );
  assert.equal(kompasAdapter.ReuseApplication, true);
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "KompasObject"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ksDocument2D"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ksDynamicArray"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ksTextParam"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ksTextLineParam"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ksTextItemParam"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ksTextItemFont"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ISymbols2DContainer"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ITable"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ITableRange"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ICellFormat"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "IText"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ITextLine"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ITextItem"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "ITextFont"));
  assert(kompasAdapter.Surfaces.some((surface) => surface.Name === "IDocumentFrame"));
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "KompasObject")?.Iid,
    "e36bc97c-39d6-4402-9c25-c7008a217e02",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ksDocument2D")?.Iid,
    "af4e160d-5c89-4f21-b0f2-d53397bdaf78",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ksDynamicArray")?.Iid,
    "4d91cd9a-6e02-409d-9360-cf7fef60d31c",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ksTextParam")?.Iid,
    "7f7d6f96-97da-11d6-8732-00c0262cdd2c",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ksTextLineParam")?.Iid,
    "364521ba-94b5-11d6-8732-00c0262cdd2c",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ksTextItemParam")?.Iid,
    "364521b7-94b5-11d6-8732-00c0262cdd2c",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ksTextItemFont")?.Iid,
    "364521bd-94b5-11d6-8732-00c0262cdd2c",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ISymbols2DContainer")?.Iid,
    "f46b0086-17f2-4489-a5a7-0aa677610afd",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "IDocumentFrame")?.Iid,
    "4437faba-990f-45e2-b1a2-7754fb326b76",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ITableRange")?.Iid,
    "d78e47dc-172b-4824-a519-9bc2c0387b5c",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ICellFormat")?.Iid,
    "9f2f27e7-8fb2-4c6c-a54d-35db240060d8",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ITextLine")?.Iid,
    "aab72fe2-dea3-4fb6-b0dd-b926249ef67c",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ITextItem")?.Iid,
    "1de74afb-5026-4b85-861f-f0cfdbd443e6",
  );
  assert.equal(
    kompasAdapter.Surfaces.find((surface) => surface.Name === "ITextFont")?.Iid,
    "a6ad008d-58d1-48b5-bd29-e6795289fe4b",
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
  assert.match(
    buildInlineTempOutputPath({
      fileName: "table_M2.xlsx",
      tempPath: "C:\\Temp",
    }),
    /^C:\\Temp\\kompas-pages\\inline\\table_M2-inline-\d+\.tbl$/,
  );
});

test("rich XLSX payload keeps run-level formatting and multiline structure", () => {
  const stylesXml = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><name val="Calibri" /><color theme="1" /><sz val="11" /></font><font><name val="Arial" /><color rgb="00FF0000" /><sz val="10" /></font></fonts><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" /><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyAlignment="1" xfId="0"><alignment horizontal="center" wrapText="1" /></xf></cellXfs></styleSheet>`;
  const worksheetXml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><r><rPr><rFont val="Arial" /><b val="1" /><color rgb="00FF0000" /><sz val="14" /></rPr><t>Bold</t></r><r><t xml:space="preserve"> + </t></r><r><rPr><rFont val="Arial" /><i val="1" /><color rgb="000000FF" /><sz val="12" /><u val="single" /></rPr><t>Blue&#10;Italic</t></r></is></c></row></sheetData></worksheet>`;
  const workbook = {
    Themes: {
      themeElements: {
        clrScheme: [
          { rgb: "FFFFFF" },
          { rgb: "000000" },
        ],
      },
    },
    files: {
      "xl/styles.xml": {
        content: new TextEncoder().encode(stylesXml),
      },
    },
  };
  const styleContext = buildWorkbookStyleContext(workbook);
  const cellMeta = parseWorksheetCellMeta(worksheetXml).A1;
  const payload = createCellTransferPayload({
    address: "A1",
    rowIndex: 0,
    columnIndex: 0,
    text: "Bold + Blue\nItalic",
    styleIndex: cellMeta.styleIndex,
    richTextXml: cellMeta.inlineRichXml,
    styleContext,
    sharedStringEntry: null,
  });

  assert.equal(payload.alignCode, 1);
  assert.equal(payload.oneLine, false);
  assert.equal(payload.lines.length, 2);
  assert.deepEqual(
    payload.lines[0].items.map((item) => item.text),
    ["Bold", " + ", "Blue"],
  );
  assert.equal(payload.lines[0].items[0].fontName, "Arial");
  assert.equal(payload.lines[0].items[0].bold, true);
  assert.equal(payload.lines[0].items[0].italic, false);
  assert.equal(payload.lines[0].items[0].color, 0xFF0000);
  assert.equal(payload.lines[0].items[0].heightMm, 4.9389);
  assert.equal(payload.lines[1].items[0].text, "Italic");
  assert.equal(payload.lines[1].items[0].italic, true);
  assert.equal(payload.lines[1].items[0].underline, true);
  assert.equal(payload.lines[1].items[0].color, 0x0000FF);
});

test("plain styled multiline XLSX payload keeps cell font, wrap and left alignment", () => {
  const stylesXml = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><name val="Calibri" /><color rgb="00000000" /><sz val="11" /></font><font><name val="Courier New" /><color rgb="00666666" /><sz val="10" /></font></fonts><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" /><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1" applyAlignment="1" xfId="0"><alignment horizontal="left" wrapText="1" /></xf></cellXfs></styleSheet>`;
  const worksheetXml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" s="1" t="inlineStr"><is><t>Mono wrap&#10;Second line&#10;Third</t></is></c></row></sheetData></worksheet>`;
  const workbook = {
    Themes: { themeElements: { clrScheme: [] } },
    files: {
      "xl/styles.xml": {
        content: new TextEncoder().encode(stylesXml),
      },
    },
  };
  const styleContext = buildWorkbookStyleContext(workbook);
  const cellMeta = parseWorksheetCellMeta(worksheetXml).A1;
  const payload = createCellTransferPayload({
    address: "A1",
    rowIndex: 0,
    columnIndex: 0,
    text: "Mono wrap\nSecond line\nThird",
    styleIndex: cellMeta.styleIndex,
    richTextXml: cellMeta.inlineRichXml,
    styleContext,
    sharedStringEntry: null,
  });

  assert.equal(payload.alignCode, 0);
  assert.equal(payload.oneLine, false);
  assert.equal(payload.lines.length, 3);
  assert.deepEqual(
    payload.lines.map((line) => line.items[0].text),
    ["Mono wrap", "Second line", "Third"],
  );
  assert.equal(payload.lines[0].items[0].fontName, "Courier New");
  assert.equal(payload.lines[0].items[0].heightMm, 3.5278);
  assert.equal(payload.lines[0].items[0].color, 0x666666);
});

test("rich XLSX payload keeps multiple fonts on right-aligned wrapped text", () => {
  const stylesXml = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><name val="Calibri" /><color rgb="00000000" /><sz val="11" /></font></fonts><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" /><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1" xfId="0"><alignment horizontal="right" wrapText="1" /></xf></cellXfs></styleSheet>`;
  const worksheetXml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="B2" s="1" t="inlineStr"><is><r><rPr><rFont val="Times New Roman" /><sz val="16" /><color rgb="0000FF" /></rPr><t>Times</t></r><r><t xml:space="preserve"> / </t></r><r><rPr><rFont val="Courier New" /><b val="1" /><u val="single" /><sz val="11" /><color rgb="800000" /></rPr><t>Courier</t></r><r><rPr><rFont val="Arial" /><i val="1" /><sz val="9" /><color rgb="008000" /></rPr><t>&#10;Tail</t></r></is></c></row></sheetData></worksheet>`;
  const workbook = {
    Themes: { themeElements: { clrScheme: [] } },
    files: {
      "xl/styles.xml": {
        content: new TextEncoder().encode(stylesXml),
      },
    },
  };
  const styleContext = buildWorkbookStyleContext(workbook);
  const cellMeta = parseWorksheetCellMeta(worksheetXml).B2;
  const payload = createCellTransferPayload({
    address: "B2",
    rowIndex: 1,
    columnIndex: 1,
    text: "Times / Courier\nTail",
    styleIndex: cellMeta.styleIndex,
    richTextXml: cellMeta.inlineRichXml,
    styleContext,
    sharedStringEntry: null,
  });

  assert.equal(payload.alignCode, 2);
  assert.equal(payload.oneLine, false);
  assert.equal(payload.lines.length, 2);
  assert.deepEqual(
    payload.lines[0].items.map((item) => item.text),
    ["Times", " / ", "Courier"],
  );
  assert.equal(payload.lines[0].items[0].fontName, "Times New Roman");
  assert.equal(payload.lines[0].items[0].heightMm, 5.6444);
  assert.equal(payload.lines[0].items[0].color, 0x0000FF);
  assert.equal(payload.lines[0].items[2].fontName, "Courier New");
  assert.equal(payload.lines[0].items[2].bold, true);
  assert.equal(payload.lines[0].items[2].underline, true);
  assert.equal(payload.lines[0].items[2].color, 0x800000);
  assert.equal(payload.lines[1].items[0].fontName, "Arial");
  assert.equal(payload.lines[1].items[0].italic, true);
  assert.equal(payload.lines[1].items[0].color, 0x008000);
});

test("wrapText keeps oneLine disabled even for a single styled line", () => {
  const stylesXml = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><name val="Calibri" /><color rgb="00000000" /><sz val="11" /></font><font><name val="Tahoma" /><color rgb="00800080" /><sz val="11" /><u val="single" /></font></fonts><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" /><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1" applyAlignment="1" xfId="0"><alignment horizontal="left" wrapText="1" /></xf></cellXfs></styleSheet>`;
  const worksheetXml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="C2" s="1" t="inlineStr"><is><t>Wrap flag only</t></is></c></row></sheetData></worksheet>`;
  const workbook = {
    Themes: { themeElements: { clrScheme: [] } },
    files: {
      "xl/styles.xml": {
        content: new TextEncoder().encode(stylesXml),
      },
    },
  };
  const styleContext = buildWorkbookStyleContext(workbook);
  const cellMeta = parseWorksheetCellMeta(worksheetXml).C2;
  const payload = createCellTransferPayload({
    address: "C2",
    rowIndex: 1,
    columnIndex: 2,
    text: "Wrap flag only",
    styleIndex: cellMeta.styleIndex,
    richTextXml: cellMeta.inlineRichXml,
    styleContext,
    sharedStringEntry: null,
  });

  assert.equal(payload.alignCode, 0);
  assert.equal(payload.oneLine, false);
  assert.equal(payload.lines.length, 1);
  assert.equal(payload.lines[0].items[0].fontName, "Tahoma");
  assert.equal(payload.lines[0].items[0].underline, true);
  assert.equal(payload.lines[0].items[0].color, 0x800080);
});

test("indexed palette font colors resolve from styles.xml", () => {
  const stylesXml = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><colors><indexedColors><rgbColor rgb="000000"/><rgbColor rgb="FFFFFF"/><rgbColor rgb="FF0000"/><rgbColor rgb="00AA00"/></indexedColors></colors><fonts count="2"><font><name val="Calibri" /><color rgb="00000000" /><sz val="11" /></font><font><name val="Consolas" /><color indexed="2" /><sz val="10" /><b val="1" /></font></fonts><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" /><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="1" applyAlignment="1" xfId="0"><alignment horizontal="right" wrapText="0" /></xf></cellXfs></styleSheet>`;
  const worksheetXml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="E2" s="1" t="inlineStr"><is><t>Code 12345</t></is></c></row></sheetData></worksheet>`;
  const workbook = {
    Themes: { themeElements: { clrScheme: [] } },
    files: {
      "xl/styles.xml": {
        content: new TextEncoder().encode(stylesXml),
      },
    },
  };
  const styleContext = buildWorkbookStyleContext(workbook);
  const cellMeta = parseWorksheetCellMeta(worksheetXml).E2;
  const payload = createCellTransferPayload({
    address: "E2",
    rowIndex: 1,
    columnIndex: 4,
    text: "Code 12345",
    styleIndex: cellMeta.styleIndex,
    richTextXml: cellMeta.inlineRichXml,
    styleContext,
    sharedStringEntry: null,
  });

  assert.equal(payload.alignCode, 2);
  assert.equal(payload.oneLine, true);
  assert.equal(payload.lines[0].items[0].fontName, "Consolas");
  assert.equal(payload.lines[0].items[0].bold, true);
  assert.equal(payload.lines[0].items[0].color, 0xFF0000);
});

test("auto font color falls back to black and preserved spaces survive rich text parsing", () => {
  const stylesXml = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><name val="Calibri" /><color auto="1" /><sz val="11" /></font></fonts><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" /><xf numFmtId="0" fontId="0" fillId="0" borderId="0" applyAlignment="1" xfId="0"><alignment horizontal="left" wrapText="0" /></xf></cellXfs></styleSheet>`;
  const worksheetXml = `<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="B4" s="1" t="inlineStr"><is><r><rPr><rFont val="Arial" /><sz val="11" /><color auto="1" /></rPr><t xml:space="preserve">A  </t></r><r><rPr><rFont val="Consolas" /><sz val="11" /><b val="1" /><color rgb="00FF6600" /></rPr><t xml:space="preserve">B  </t></r><r><rPr><rFont val="Arial" /><sz val="11" /><color auto="1" /></rPr><t>C</t></r></is></c></row></sheetData></worksheet>`;
  const workbook = {
    Themes: { themeElements: { clrScheme: [] } },
    files: {
      "xl/styles.xml": {
        content: new TextEncoder().encode(stylesXml),
      },
    },
  };
  const styleContext = buildWorkbookStyleContext(workbook);
  const cellMeta = parseWorksheetCellMeta(worksheetXml).B4;
  const payload = createCellTransferPayload({
    address: "B4",
    rowIndex: 3,
    columnIndex: 1,
    text: "A  B  C",
    styleIndex: cellMeta.styleIndex,
    richTextXml: cellMeta.inlineRichXml,
    styleContext,
    sharedStringEntry: null,
  });

  assert.equal(payload.oneLine, true);
  assert.deepEqual(
    payload.lines[0].items.map((item) => item.text),
    ["A  ", "B  ", "C"],
  );
  assert.equal(payload.lines[0].items[0].color, 0x000000);
  assert.equal(payload.lines[0].items[1].fontName, "Consolas");
  assert.equal(payload.lines[0].items[1].bold, true);
  assert.equal(payload.lines[0].items[1].color, 0xFF6600);
});

test("auto-fit shrinks formatted content for a tight cell", () => {
  const sourceCell = {
    address: "A1",
    rowIndex: 0,
    columnIndex: 0,
    text: "Long title",
    horizontal: "center",
    alignCode: 1,
    wrapText: false,
    oneLine: true,
    hasContent: true,
    lines: [
      {
        items: [
          {
            text: "Long title",
            fontName: "Arial",
            heightMm: 6,
            bold: true,
            italic: false,
            underline: false,
            color: 0,
            widthFactor: 1,
          },
        ],
      },
    ],
  };

  const result = autoFitCellMatrixToLayout([[sourceCell]], {
    cellWidthMm: 12,
    cellHeightMm: 4,
  }, true);

  assert.equal(result.stats.enabled, true);
  assert.equal(result.stats.adjustedCellCount, 1);
  assert.ok(result.stats.minScale < 1);
  assert.ok(result.cellMatrix[0][0].lines[0].items[0].heightMm < sourceCell.lines[0].items[0].heightMm);
  assert.equal(sourceCell.lines[0].items[0].heightMm, 6);
});

test("auto-fit grows formatted content when the cell is spacious", () => {
  const sourceCell = {
    address: "B2",
    rowIndex: 1,
    columnIndex: 1,
    text: "OK",
    horizontal: "left",
    alignCode: 0,
    wrapText: false,
    oneLine: true,
    hasContent: true,
    lines: [
      {
        items: [
          {
            text: "OK",
            fontName: "Calibri",
            heightMm: 2,
            bold: false,
            italic: false,
            underline: false,
            color: 0,
            widthFactor: 1,
          },
        ],
      },
    ],
  };

  const result = autoFitCellMatrixToLayout([[sourceCell]], {
    cellWidthMm: 30,
    cellHeightMm: 12,
  }, true);

  assert.equal(result.stats.enabled, true);
  assert.equal(result.stats.adjustedCellCount, 1);
  assert.ok(result.stats.maxScale > 1);
  assert.ok(result.cellMatrix[0][0].lines[0].items[0].heightMm > sourceCell.lines[0].items[0].heightMm);
});

test("formatted cell batching emits native line and item commands", () => {
  const stylesXml = `<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="1"><font><name val="Calibri" /><color rgb="00000000" /><sz val="11" /></font></fonts><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" /></cellXfs></styleSheet>`;
  const workbook = {
    Themes: { themeElements: { clrScheme: [] } },
    files: {
      "xl/styles.xml": {
        content: new TextEncoder().encode(stylesXml),
      },
    },
  };
  const styleContext = buildWorkbookStyleContext(workbook);
  const payload = createCellTransferPayload({
    address: "A1",
    rowIndex: 0,
    columnIndex: 0,
    text: "Line 1\nLine 2",
    styleIndex: 0,
    richTextXml: "",
    styleContext,
    sharedStringEntry: null,
  });
  const batches = createFormattedCellWriteBatches([[payload]], 2);

  assert.equal(batches.length, 1);
  assert.equal(batches[0].cellCount, 1);
  assert.deepEqual(
    batches[0].commands.map((command) => command.commandId),
    [
      "xlsx-to-kompas-tbl.table-cell-clear-text",
      "xlsx-to-kompas-tbl.table-cell-set-one-line",
      "xlsx-to-kompas-tbl.table-cell-add-line",
      "xlsx-to-kompas-tbl.table-cell-add-item-before",
      "xlsx-to-kompas-tbl.table-cell-add-line",
      "xlsx-to-kompas-tbl.table-cell-add-item-before",
    ],
  );
  assert.equal(batches[0].commands[1].arguments.oneLine, false);
  assert.equal(batches[0].commands[3].arguments.lineIndex, 0);
  assert.equal(batches[0].commands[3].arguments.itemIndex, 0);
  assert.equal(batches[0].commands[3].arguments.value, "Line 1");
  assert.equal(batches[0].commands[5].arguments.lineIndex, 1);
  assert.equal(batches[0].commands[5].arguments.itemIndex, 0);
  assert.equal(batches[0].commands[5].arguments.value, "Line 2");
  assert.equal(batches[0].commands[3].arguments.fontName, "Calibri");
  assert.equal(batches[0].commands[3].arguments.heightMm, 3.8806);
});

test("xlsx runtime contribution uses active frame center commands for visible inline placement", () => {
  const dsl = {
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
      if (options.args) {
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

  const contribution = createXlsxToKompasTblModule().getRuntimeContribution({ dsl });
  const frameCenterX = contribution.commands["xlsx-to-kompas-tbl.active-frame-center-x"];
  const frameRefresh = contribution.commands["xlsx-to-kompas-tbl.active-frame-refresh"];
  const setOneLine = contribution.commands["xlsx-to-kompas-tbl.table-cell-set-one-line"];
  const addItemBefore = contribution.commands["xlsx-to-kompas-tbl.table-cell-add-item-before"];
  const setItem = contribution.commands["xlsx-to-kompas-tbl.table-cell-set-item"];
  const addItem = contribution.commands["xlsx-to-kompas-tbl.table-cell-add-item"];
  const getItemColor = contribution.commands["xlsx-to-kompas-tbl.table-cell-get-item-color"];
  const api5ActiveDocument2D = contribution.commands["xlsx-to-kompas-tbl.api5-active-document2d"];
  const api5CreateTextParam = contribution.commands["xlsx-to-kompas-tbl.api5-create-text-param"];
  const api5SetTableCellText = contribution.commands["xlsx-to-kompas-tbl.api5-document-set-table-cell-text"];

  assert.equal(frameCenterX.Invoke.ReturnPath, "stored:x");
  assert.equal(frameCenterX.Invoke.Chain.at(-1).Member, "GetZoomScale");
  assert.deepEqual(
    frameCenterX.Invoke.Chain.at(-1).Args,
    [
      { Literal: 0, Converter: "double", ByRef: true, CaptureAs: "x" },
      { Literal: 0, Converter: "double", ByRef: true, CaptureAs: "y" },
      { Literal: 0, Converter: "double", ByRef: true, CaptureAs: "scale" },
    ],
  );
  assert.equal(frameCenterX.Invoke.Chain.at(-2).Member, "IDocumentFrame");
  assert.equal(frameRefresh.Invoke.Chain.at(-1).Member, "RefreshWindow");
  assert.equal(setOneLine.Invoke.Chain.at(1).Member, "Range");
  assert.equal(setOneLine.Invoke.Chain.at(-2).Member, "ICellFormat");
  assert.equal(setOneLine.Invoke.Chain.at(-1).Member, "OneLine");
  assert.equal(setItem.Invoke.Chain.at(7).Member, "TextItem");
  assert.equal(setItem.Invoke.Chain.at(8).Member, "ITextItem");
  assert.equal(setItem.Invoke.Chain.at(10).Member, "ITextFont");
  assert.equal(addItemBefore.Invoke.Chain.at(7).Member, "AddBefore");
  assert.equal(addItemBefore.Invoke.Chain.at(8).Member, "ITextItem");
  assert.equal(addItem.Invoke.Chain.at(5).Member, "TextLine");
  assert.equal(addItem.Invoke.Chain.at(6).Member, "ITextLine");
  assert.equal(addItem.Invoke.Chain.at(8).Member, "ITextItem");
  assert.equal(addItem.Invoke.Chain.at(10).Member, "ITextFont");
  assert.equal(addItem.Invoke.Chain.at(11).Member, "FontName");
  assert.equal(addItem.Invoke.Chain.at(15).Member, "Color");
  assert.equal(addItem.Invoke.Chain.at(16).Member, "WidthFactor");
  assert.equal(getItemColor.Invoke.Chain.at(9).Member, "ITextFont");
  assert.equal(getItemColor.Invoke.Chain.at(10).Member, "Color");
  assert.deepEqual(api5ActiveDocument2D.DefaultArguments?.progIds, ["KOMPAS.Application.5"]);
  assert.equal(api5ActiveDocument2D.Invoke.Chain.length, 1);
  assert.equal(api5ActiveDocument2D.Invoke.Chain[0].Member, "ActiveDocument2D");
  assert.deepEqual(api5CreateTextParam.DefaultArguments?.progIds, ["KOMPAS.Application.5"]);
  assert.equal(api5CreateTextParam.Invoke.Chain.length, 1);
  assert.equal(api5CreateTextParam.Invoke.Chain[0].Member, "GetParamStruct");
  assert.equal(api5SetTableCellText.Invoke.Chain.length, 1);
  assert.equal(api5SetTableCellText.Invoke.Chain[0].Member, "ksSetTableColumnText");
});
