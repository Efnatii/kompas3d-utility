import { KOMPAS_COM_ADAPTER } from "../executor-shell.js";

const DEFAULT_LAYOUT = {
  tableWidthMm: 260,
  tableHeightMm: 24,
  cellWidthMm: 30,
  cellHeightMm: 8,
};

const EXPORT_BATCH_SIZE = 200;
const FALLBACK_BRIDGE_SCRIPT_NAME = "xlsx-to-kompas-tbl.bridge.ps1";
const FALLBACK_BRIDGE_ERROR_MARKERS = [
  "public member 'drawingtables' on type 'iview' not found",
  "drawingtables",
  "isymbols2dcontainer",
];

function stripExtension(fileName) {
  return String(fileName || "").replace(/\.[^.]+$/, "");
}

function normalizeWindowsPath(value) {
  return String(value || "")
    .replace(/\//g, "\\")
    .replace(/\\{2,}/g, "\\");
}

function joinWindowsPath(left, right) {
  const lhs = normalizeWindowsPath(left).replace(/[\\]+$/, "");
  const rhs = normalizeWindowsPath(right).replace(/^[\\]+/, "");
  if (!lhs) {
    return rhs;
  }
  if (!rhs) {
    return lhs;
  }
  return `${lhs}\\${rhs}`;
}

function dirname(filePath) {
  const value = normalizeWindowsPath(filePath);
  const index = value.lastIndexOf("\\");
  return index >= 0 ? value.slice(0, index) : "";
}

function sanitizeFileStem(fileName) {
  const stem = stripExtension(fileName || "table")
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "_")
    .trim();
  return stem || "table";
}

function ensureTblExtension(filePath) {
  const value = normalizeWindowsPath(filePath);
  return /\.tbl$/i.test(value) ? value : `${value}.tbl`;
}

function formatSize(rows, cols) {
  return `${rows} x ${cols}`;
}

function formatMm(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric)) {
    return "0";
  }
  return numeric.toFixed(4).replace(/\.?0+$/, "");
}

function roundLayoutValue(value) {
  const numeric = Number(value);
  if (!Number.isFinite(numeric) || numeric <= 0) {
    return 0;
  }
  return Number(numeric.toFixed(4));
}

function readNumeric(input, fallback = 0) {
  const numeric = Number.parseFloat(input.value);
  return Number.isFinite(numeric) ? numeric : fallback;
}

function formatCellValue(cell) {
  if (!cell) {
    return "";
  }
  if (typeof cell.w === "string") {
    return cell.w;
  }
  if (cell.v === null || cell.v === undefined) {
    return "";
  }
  if (typeof cell.v === "boolean") {
    return cell.v ? "True" : "False";
  }
  return String(cell.v);
}

function getMatrixDimensions(matrix) {
  const rows = Array.isArray(matrix) ? matrix.length : 0;
  let cols = 0;
  for (const row of matrix || []) {
    cols = Math.max(cols, Array.isArray(row) ? row.length : 0);
  }
  return { rows, cols };
}

function reconcileLinkedLayout(layout, rows, cols, source = "table") {
  const safeRows = Math.max(1, Number(rows) || 1);
  const safeCols = Math.max(1, Number(cols) || 1);
  const next = {
    tableWidthMm: roundLayoutValue(layout?.tableWidthMm) || DEFAULT_LAYOUT.tableWidthMm,
    tableHeightMm: roundLayoutValue(layout?.tableHeightMm) || DEFAULT_LAYOUT.tableHeightMm,
    cellWidthMm: roundLayoutValue(layout?.cellWidthMm) || DEFAULT_LAYOUT.cellWidthMm,
    cellHeightMm: roundLayoutValue(layout?.cellHeightMm) || DEFAULT_LAYOUT.cellHeightMm,
  };

  if (source === "cell") {
    next.tableWidthMm = roundLayoutValue(next.cellWidthMm * safeCols);
    next.tableHeightMm = roundLayoutValue(next.cellHeightMm * safeRows);
    return next;
  }

  next.cellWidthMm = roundLayoutValue(next.tableWidthMm / safeCols);
  next.cellHeightMm = roundLayoutValue(next.tableHeightMm / safeRows);
  return next;
}

function buildAutoOutputPath({ documentPath, fileName, tempPath }) {
  const targetDirectory = documentPath
    ? dirname(documentPath)
    : joinWindowsPath(tempPath || "", "kompas-pages");
  if (!targetDirectory) {
    return "";
  }
  return ensureTblExtension(joinWindowsPath(targetDirectory, sanitizeFileStem(fileName)));
}

function collectCellWrites(matrix) {
  const writes = [];
  const { rows, cols } = getMatrixDimensions(matrix);
  for (let rowIndex = 0; rowIndex < rows; rowIndex += 1) {
    const row = Array.isArray(matrix[rowIndex]) ? matrix[rowIndex] : [];
    for (let columnIndex = 0; columnIndex < cols; columnIndex += 1) {
      const value = String(row[columnIndex] ?? "");
      if (value === "") {
        continue;
      }
      writes.push({ rowIndex, columnIndex, value });
    }
  }
  return writes;
}

function chunkList(values, size) {
  const chunks = [];
  for (let index = 0; index < values.length; index += size) {
    chunks.push(values.slice(index, index + size));
  }
  return chunks;
}

function createCellWriteBatches(matrix, batchSize = EXPORT_BATCH_SIZE) {
  return chunkList(collectCellWrites(matrix), batchSize).map((batch) => batch.map((cell) => ({
    commandId: "xlsx-to-kompas-tbl.table-cell-set-text",
    arguments: {
      rowIndex: cell.rowIndex,
      columnIndex: cell.columnIndex,
      value: cell.value,
    },
    timeoutMilliseconds: 30000,
  })));
}

function readWorkbookMatrix(file, bytes) {
  const workbook = window.XLSX.read(bytes, {
    type: "array",
    cellNF: true,
    cellText: true,
    cellDates: false,
  });
  const firstSheetName = workbook.SheetNames[0];
  if (!firstSheetName) {
    throw new Error("XLSX workbook does not contain worksheets.");
  }

  const sheet = workbook.Sheets[firstSheetName];
  const range = window.XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const matrix = [];

  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row = [];
    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const address = window.XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      row.push(formatCellValue(sheet[address]));
    }
    matrix.push(row);
  }

  return {
    fileName: file.name,
    sheetName: firstSheetName,
    matrix,
  };
}

function buildFallbackBridgeScriptPath(tempPath) {
  return joinWindowsPath(joinWindowsPath(tempPath || "", "kompas-pages"), FALLBACK_BRIDGE_SCRIPT_NAME);
}

function buildFallbackBridgeArguments(scriptPath, action) {
  const quote = (value) => `"${String(value || "").replace(/"/g, '\\"')}"`;
  return `-NoLogo -NoProfile -ExecutionPolicy Bypass -File ${quote(scriptPath)} -Action ${quote(action)}`;
}

function shouldUseFallbackBridge(error) {
  const text = String(error?.message || error || "").toLowerCase();
  return FALLBACK_BRIDGE_ERROR_MARKERS.some((marker) => text.includes(marker));
}

function createXlsxToKompasTblModule() {
  return {
    id: "xlsx-to-kompas-tbl",
    title: "XLSX to KOMPAS TBL",
    subtitle: "XLSX парсится в браузере, а export и insert выполняются через runtime overlay WebBridge.Utility.",
    tabLabel: "XLSX/TBL",
    tabDetail: "active",

    getRuntimeContribution(context) {
      return {
        commands: {
          "xlsx-to-kompas-tbl.application.info": context.dsl.command(
            "kompas",
            "application",
            [],
            {
              defaultArguments: {
                attachOnly: true,
                createIfMissing: false,
                visible: false,
              },
            },
          ),
          "xlsx-to-kompas-tbl.active-document": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("get", "ActiveDocument"),
            ],
            {
              defaultArguments: {
                attachOnly: true,
                createIfMissing: false,
                visible: false,
              },
            },
          ),
          "xlsx-to-kompas-tbl.active-view": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("get", "ViewsAndLayersManager"),
              context.dsl.step("get", "Views"),
              context.dsl.step("get", "ActiveView"),
            ],
          ),
          "xlsx-to-kompas-tbl.view-table-count": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("get", "DrawingTables"),
              context.dsl.step("get", "Count"),
            ],
          ),
          "xlsx-to-kompas-tbl.create-table": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("get", "DrawingTables"),
              context.dsl.step("call", "Add", {
                args: [
                  context.dsl.arg("rows", "int"),
                  context.dsl.arg("cols", "int"),
                  context.dsl.arg("cellHeightMm", "double"),
                  context.dsl.arg("cellWidthMm", "double"),
                  context.dsl.literal(0, "int"),
                ],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.table-cell-set-text": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("index", "Cell", {
                args: [
                  context.dsl.arg("rowIndex", "int"),
                  context.dsl.arg("columnIndex", "int"),
                ],
              }),
              context.dsl.step("get", "Text"),
              context.dsl.step("set", "Str", {
                valueArgument: "value",
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.table-save": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("call", "Save", {
                args: [context.dsl.arg("path", "path")],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.insert-table": context.dsl.command(
            "kompas",
            "handle",
            [
              context.dsl.step("get", "DrawingTables"),
              context.dsl.step("call", "Load", {
                args: [context.dsl.arg("path", "path")],
              }),
            ],
          ),
          "xlsx-to-kompas-tbl.open-document": context.dsl.command(
            "kompas",
            "application",
            [
              context.dsl.step("get", "Documents"),
              context.dsl.step("call", "Open", {
                args: [
                  context.dsl.arg("path", "path"),
                  context.dsl.literal(false, "bool"),
                  context.dsl.literal(true, "bool"),
                ],
              }),
            ],
            {
              defaultArguments: {
                attachOnly: false,
                createIfMissing: true,
                visible: true,
              },
            },
          ),
        },
        allowedTypes: [],
        comAdapters: [KOMPAS_COM_ADAPTER],
      };
    },

    mount(container, context) {
      const savedLayout = context.storage.get("layout", DEFAULT_LAYOUT) || DEFAULT_LAYOUT;
      const initialLayoutDriver = savedLayout.layoutDriver === "cell" ? "cell" : "table";
      const state = {
        fileName: "",
        sheetName: "",
        matrix: [],
        status: null,
        tempPath: "",
        tempPathLoaded: false,
        lastExport: null,
        lastBytes: null,
        autoFollowOutput: true,
        layoutDriver: initialLayoutDriver,
        layout: reconcileLinkedLayout(savedLayout, 1, 1, initialLayoutDriver),
        pollTimer: null,
        pollBusy: false,
        active: false,
        dragDepth: 0,
        fallbackScriptTextPromise: null,
        fallbackScriptPath: "",
        fallbackScriptReady: false,
      };

      container.innerHTML = `
        <div class="module-grid">
          <div class="stack">
            <section class="status-box">
              <div class="status-row"><span>File</span><strong id="xlsx-file-name">not loaded</strong></div>
              <div class="status-row"><span>Sheet</span><strong id="xlsx-sheet-name">-</strong></div>
              <div class="status-row"><span>Matrix</span><strong id="xlsx-matrix-size">0 x 0</strong></div>
              <div class="status-row"><span>Document</span><strong id="xlsx-document-name">no active doc</strong></div>
              <div class="status-row"><span>View</span><strong id="xlsx-view-name">-</strong></div>
            </section>

            <section class="panel panel--inner">
              <div class="panel__head">
                <div>
                  <h2>Source</h2>
                  <p class="panel__subtitle">Автостатус KOMPAS обновляется каждые 2 секунды при активной вкладке.</p>
                </div>
              </div>

              <div class="dropzone" id="xlsx-dropzone">
                <div class="dropzone__copy">
                  <strong>Drop XLSX</strong>
                  <span>Перетащите .xlsx сюда или откройте файл вручную.</span>
                </div>
                <label class="button button--ghost" for="xlsx-file-input">Open XLSX</label>
              </div>

              <input id="xlsx-file-input" type="file" accept=".xlsx,.xlsm,.xlsb,.xls" hidden>

              <div class="field-grid">
                <label class="field">
                  <span>table width, mm</span>
                  <input id="xlsx-table-width" type="number" step="0.1" min="0" value="${state.layout.tableWidthMm}">
                </label>
                <label class="field">
                  <span>table height, mm</span>
                  <input id="xlsx-table-height" type="number" step="0.1" min="0" value="${state.layout.tableHeightMm}">
                </label>
                <label class="field">
                  <span>cell width, mm</span>
                  <input id="xlsx-cell-width" type="number" step="0.1" min="0" value="${state.layout.cellWidthMm}">
                </label>
                <label class="field">
                  <span>cell height, mm</span>
                  <input id="xlsx-cell-height" type="number" step="0.1" min="0" value="${state.layout.cellHeightMm}">
                </label>
              </div>

              <label class="field">
                <span>output path</span>
                <div class="field-inline">
                  <input id="xlsx-output-path" type="text" spellcheck="false" placeholder="%TEMP%\\kompas-pages\\table.tbl">
                  <button class="button button--ghost" type="button" id="xlsx-follow-button">Follow</button>
                </div>
                <small class="field__hint" id="xlsx-output-mode">auto-follow: waiting for XLSX</small>
              </label>

              <div class="summary-box" id="xlsx-layout-summary">Размеры будут рассчитаны после загрузки файла.</div>
              <div class="result-box" id="xlsx-result-box">Ожидание данных.</div>

              <div class="action-row">
                <button class="button" type="button" id="xlsx-export-button">Export</button>
                <button class="button button--ghost" type="button" id="xlsx-insert-button" disabled>Insert</button>
                <button class="button button--ghost" type="button" id="xlsx-download-button" disabled>Download</button>
                <button class="button button--ghost" type="button" id="xlsx-reset-button">Reset</button>
              </div>
            </section>
          </div>

          <section class="panel panel--inner">
            <div class="panel__head">
              <div>
                <h2>Preview</h2>
                <p class="panel__subtitle" id="xlsx-preview-meta">UsedRange первого листа.</p>
              </div>
            </div>
            <div class="preview-wrap">
              <table class="preview-table" id="xlsx-preview-table">
                <tbody>
                  <tr><td class="preview-table__empty">Загрузите XLSX.</td></tr>
                </tbody>
              </table>
            </div>
          </section>
        </div>
      `;

      const refs = {
        dropzone: container.querySelector("#xlsx-dropzone"),
        fileInput: container.querySelector("#xlsx-file-input"),
        fileName: container.querySelector("#xlsx-file-name"),
        sheetName: container.querySelector("#xlsx-sheet-name"),
        matrixSize: container.querySelector("#xlsx-matrix-size"),
        documentName: container.querySelector("#xlsx-document-name"),
        viewName: container.querySelector("#xlsx-view-name"),
        tableWidth: container.querySelector("#xlsx-table-width"),
        tableHeight: container.querySelector("#xlsx-table-height"),
        cellWidth: container.querySelector("#xlsx-cell-width"),
        cellHeight: container.querySelector("#xlsx-cell-height"),
        outputPath: container.querySelector("#xlsx-output-path"),
        outputMode: container.querySelector("#xlsx-output-mode"),
        followButton: container.querySelector("#xlsx-follow-button"),
        layoutSummary: container.querySelector("#xlsx-layout-summary"),
        resultBox: container.querySelector("#xlsx-result-box"),
        exportButton: container.querySelector("#xlsx-export-button"),
        insertButton: container.querySelector("#xlsx-insert-button"),
        downloadButton: container.querySelector("#xlsx-download-button"),
        resetButton: container.querySelector("#xlsx-reset-button"),
        previewMeta: container.querySelector("#xlsx-preview-meta"),
        previewTable: container.querySelector("#xlsx-preview-table"),
      };

      function persistLayout() {
        context.storage.set("layout", {
          ...state.layout,
          layoutDriver: state.layoutDriver,
        });
      }

      function currentDimensions() {
        return getMatrixDimensions(state.matrix);
      }

      function syncLayoutInputs() {
        refs.tableWidth.value = String(state.layout.tableWidthMm || 0);
        refs.tableHeight.value = String(state.layout.tableHeightMm || 0);
        refs.cellWidth.value = String(state.layout.cellWidthMm || 0);
        refs.cellHeight.value = String(state.layout.cellHeightMm || 0);
      }

      function renderPreview() {
        const body = document.createElement("tbody");
        if (!state.matrix.length) {
          const row = document.createElement("tr");
          const cell = document.createElement("td");
          cell.className = "preview-table__empty";
          cell.textContent = "Загрузите XLSX.";
          row.append(cell);
          body.append(row);
          refs.previewTable.replaceChildren(body);
          return;
        }

        for (const rowValues of state.matrix) {
          const row = document.createElement("tr");
          for (const value of rowValues) {
            const cell = document.createElement("td");
            cell.textContent = value;
            row.append(cell);
          }
          body.append(row);
        }
        refs.previewTable.replaceChildren(body);
      }

      function suggestOutputPath() {
        if (!state.autoFollowOutput || !state.fileName) {
          return;
        }
        const suggested = buildAutoOutputPath({
          documentPath: state.status?.documentPath || "",
          fileName: state.fileName,
          tempPath: state.tempPath,
        });
        if (suggested) {
          refs.outputPath.value = suggested;
        }
      }

      function renderSummary() {
        const { rows, cols } = currentDimensions();
        if (!rows || !cols) {
          refs.layoutSummary.textContent = "Размеры будут рассчитаны после загрузки файла.";
          return;
        }
        const writeCount = collectCellWrites(state.matrix).length;
        refs.layoutSummary.textContent = [
          `cell=${formatMm(state.layout.cellWidthMm)} x ${formatMm(state.layout.cellHeightMm)} mm`,
          `table=${formatMm(state.layout.tableWidthMm)} x ${formatMm(state.layout.tableHeightMm)} mm`,
          `writes=${writeCount}/${rows * cols}`,
        ].join(" | ");
      }

      function renderStatus() {
        refs.fileName.textContent = state.fileName || "not loaded";
        refs.sheetName.textContent = state.sheetName || "-";
        refs.matrixSize.textContent = formatSize(currentDimensions().rows, currentDimensions().cols);
        refs.documentName.textContent = state.status?.documentPath || state.status?.documentName || "no active doc";
        refs.viewName.textContent = state.status?.viewName || "-";
        refs.previewMeta.textContent = state.sheetName
          ? `${state.sheetName} | ${formatSize(currentDimensions().rows, currentDimensions().cols)}`
          : "UsedRange первого листа.";
      }

      function renderOutputMode() {
        refs.followButton.classList.toggle("is-active", state.autoFollowOutput);
        if (state.autoFollowOutput) {
          refs.outputMode.textContent = state.fileName
            ? "auto-follow: путь следует за активным документом KOMPAS или %TEMP%"
            : "auto-follow: waiting for XLSX";
          return;
        }
        refs.outputMode.textContent = "manual path: follow отключён до Reset или кнопки Follow";
      }

      function updateModuleIndicator() {
        if (!context.getBridgeState().runtimeReady) {
          context.setModuleBadge("doc idle", false);
          context.setModuleMeta("Bridge подключён не полностью или runtime ещё не загружен.");
          return;
        }
        if (state.status?.connected && state.status?.viewHandleId) {
          context.setModuleBadge("doc ok", true);
          context.setModuleMeta(`${state.status.documentPath || state.status.documentName} | view=${state.status.viewName || "-"}`);
          return;
        }
        if (state.status?.connected) {
          context.setModuleBadge("doc none", false);
          context.setModuleMeta("KOMPAS запущен, но активный 2D документ или вид недоступен.");
          return;
        }
        context.setModuleBadge("kompas off", false);
        context.setModuleMeta(state.status?.errorMessage || "KOMPAS не найден.");
      }

      function updateActionState() {
        const bridgeState = context.getBridgeState();
        const hasMatrix = currentDimensions().rows > 0 && currentDimensions().cols > 0;
        const hasExportPath = Boolean((state.lastExport?.outputPath || refs.outputPath.value || "").trim());
        const hasBytes = Boolean(state.lastBytes && state.lastBytes.length > 0);

        refs.exportButton.disabled = !bridgeState.runtimeReady || !hasMatrix;
        refs.insertButton.disabled = !bridgeState.runtimeReady || !hasExportPath;
        refs.downloadButton.disabled = !hasBytes;
      }

      function renderAll() {
        syncLayoutInputs();
        suggestOutputPath();
        renderSummary();
        renderStatus();
        renderOutputMode();
        renderPreview();
        updateModuleIndicator();
        updateActionState();
      }

      async function ensureTempPathLoaded(force = false) {
        if (!context.getBridgeState().runtimeReady) {
          return "";
        }
        if (state.tempPathLoaded && !force) {
          return state.tempPath;
        }
        try {
          state.tempPath = normalizeWindowsPath(await context.getTempPath());
          state.tempPathLoaded = true;
        } catch {
          state.tempPath = "";
        }
        return state.tempPath;
      }

      async function refreshStatus(options = {}) {
        const quiet = options.quiet === true;
        if (!context.getBridgeState().runtimeReady) {
          state.status = null;
          renderAll();
          return state.status;
        }

        await ensureTempPathLoaded();

        try {
          const applicationExecution = await context.executeCommand("xlsx-to-kompas-tbl.application.info", {}, 15000);
          const application = applicationExecution.result || {};
          const documentExecution = await context.executeCommand("xlsx-to-kompas-tbl.active-document", {}, 15000);
          const documentResult = documentExecution.result || null;
          let viewResult = null;
          if (documentResult?.handleId) {
            const viewExecution = await context.executeCommand(
              "xlsx-to-kompas-tbl.active-view",
              { handleId: documentResult.handleId },
              15000,
            );
            viewResult = viewExecution.result || null;
          }

          state.status = {
            connected: Boolean(application.connected ?? true),
            applicationProgId: String(application.progId || ""),
            documentName: String(documentResult?.name || ""),
            documentPath: String(documentResult?.path || ""),
            documentHandleId: String(documentResult?.handleId || ""),
            viewName: String(viewResult?.name || ""),
            viewHandleId: String(viewResult?.handleId || ""),
            hasActiveDocument: Boolean(documentResult?.handleId),
            errorMessage: "",
          };
        } catch (error) {
          state.status = {
            connected: false,
            applicationProgId: "",
            documentName: "",
            documentPath: "",
            documentHandleId: "",
            viewName: "",
            viewHandleId: "",
            hasActiveDocument: false,
            errorMessage: String(error.message || error),
          };
          if (!quiet) {
            context.logger.error("status", state.status.errorMessage);
          }
        }

        renderAll();
        return state.status;
      }

      function syncPolling() {
        const shouldPoll = state.active
          && context.getBridgeState().runtimeReady
          && document.visibilityState === "visible";
        if (!shouldPoll) {
          if (state.pollTimer !== null) {
            window.clearInterval(state.pollTimer);
            state.pollTimer = null;
          }
          return;
        }
        if (state.pollTimer !== null) {
          return;
        }

        state.pollTimer = window.setInterval(() => {
          if (state.pollBusy) {
            return;
          }
          state.pollBusy = true;
          refreshStatus({ quiet: true })
            .catch(() => {})
            .finally(() => {
              state.pollBusy = false;
            });
        }, 2000);
      }

      function applyLayoutChange(driver, patch) {
        const { rows, cols } = currentDimensions();
        state.layoutDriver = driver;
        state.layout = reconcileLinkedLayout(
          { ...state.layout, ...patch },
          rows || 1,
          cols || 1,
          driver,
        );
        persistLayout();
        renderAll();
      }

      async function resolveOutputPath() {
        await ensureTempPathLoaded();
        const manualValue = normalizeWindowsPath(refs.outputPath.value.trim());
        const effective = manualValue
          ? ensureTblExtension(manualValue)
          : buildAutoOutputPath({
            documentPath: state.status?.documentPath || "",
            fileName: state.fileName || "table.xlsx",
            tempPath: state.tempPath,
          });
        if (!effective) {
          throw new Error("Не удалось определить output path.");
        }
        refs.outputPath.value = effective;
        return effective;
      }

      async function ensureActiveViewStatus() {
        const status = await refreshStatus({ quiet: true });
        if (!status?.viewHandleId) {
          throw new Error("Активный 2D документ KOMPAS или его view не найден.");
        }
        return status;
      }

      async function loadFallbackBridgeScriptText() {
        if (!state.fallbackScriptTextPromise) {
          const scriptUrl = new URL(`./${FALLBACK_BRIDGE_SCRIPT_NAME}`, import.meta.url);
          state.fallbackScriptTextPromise = fetch(scriptUrl).then(async (response) => {
            if (!response.ok) {
              throw new Error(`Не удалось загрузить fallback bridge asset: ${response.status}`);
            }
            return response.text();
          });
        }
        return state.fallbackScriptTextPromise;
      }

      async function ensureFallbackBridgeScriptPath() {
        const tempPath = await ensureTempPathLoaded();
        if (!tempPath) {
          throw new Error("TEMP path недоступен для fallback bridge.");
        }
        const scriptPath = buildFallbackBridgeScriptPath(tempPath);
        if (state.fallbackScriptReady && state.fallbackScriptPath === scriptPath) {
          return scriptPath;
        }

        const scriptDirectory = dirname(scriptPath);
        const scriptText = await loadFallbackBridgeScriptText();
        if (!scriptDirectory) {
          throw new Error("Не удалось определить директорию fallback bridge script.");
        }

        await context.ensureDirectory(scriptDirectory);
        await context.writeFileText(scriptPath, scriptText);
        state.fallbackScriptPath = scriptPath;
        state.fallbackScriptReady = true;
        return scriptPath;
      }

      function parseFallbackBridgePayload(commandResult) {
        const stdout = String(commandResult?.stdout || "");
        const lines = stdout
          .split(/\r?\n/g)
          .map((line) => line.trim())
          .filter(Boolean);
        for (let index = lines.length - 1; index >= 0; index -= 1) {
          try {
            return JSON.parse(lines[index]);
          } catch {
            // Skip non-JSON lines emitted by PowerShell.
          }
        }
        return null;
      }

      async function invokeFallbackBridge(action, payload, timeoutMilliseconds = 180000) {
        const scriptPath = await ensureFallbackBridgeScriptPath();
        const result = await context.runCommand(
          "powershell.exe",
          buildFallbackBridgeArguments(scriptPath, action),
          dirname(scriptPath),
          timeoutMilliseconds,
          JSON.stringify(payload || {}),
          null,
        );

        if (result.timedOut) {
          throw new Error(`Fallback bridge timed out after ${timeoutMilliseconds} ms.`);
        }

        const bridgePayload = parseFallbackBridgePayload(result);
        if (bridgePayload?.success === false) {
          throw new Error(String(bridgePayload.errorMessage || bridgePayload.errorCode || "Fallback bridge failed."));
        }
        if (Number(result.exitCode || 0) !== 0 && !bridgePayload?.success) {
          throw new Error(String(result.stderr || result.stdout || `Fallback bridge exited with code ${result.exitCode}.`).trim());
        }
        if (!bridgePayload) {
          throw new Error(String(result.stderr || "Fallback bridge did not return JSON payload.").trim());
        }

        return bridgePayload;
      }

      async function exportTableViaRuntime(status, outputPath, rows, cols) {
        const sharedContextId = `xlsx-export-${Date.now()}-${Math.random().toString(16).slice(2)}`;
        const createExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.create-table",
          {
            handleId: status.viewHandleId,
            rows,
            cols,
            cellWidthMm: state.layout.cellWidthMm,
            cellHeightMm: state.layout.cellHeightMm,
            __sharedContextId: sharedContextId,
          },
          60000,
        );
        if (!createExecution.result?.handleId) {
          throw new Error("KOMPAS did not return a drawing table handle.");
        }

        const batches = createCellWriteBatches(state.matrix, EXPORT_BATCH_SIZE);
        const totalWrites = collectCellWrites(state.matrix).length;
        let completedWrites = 0;
        for (const batch of batches) {
          await context.executeBatchCommand(batch, {
            sharedContextId,
            reportVerbosity: "Compact",
            stopOnError: true,
          });
          completedWrites += batch.length;
          refs.resultBox.textContent = `Экспорт: ${completedWrites}/${totalWrites} ячеек`;
        }

        const saveExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.table-save",
          {
            path: outputPath,
            __sharedContextId: sharedContextId,
          },
          60000,
        );
        if (saveExecution.result === false) {
          throw new Error("KOMPAS returned Save=false.");
        }

        return {
          outputPath,
          rows,
          cols,
        };
      }

      async function exportTableViaFallback(outputPath, rows, cols) {
        refs.resultBox.textContent = "Экспорт выполняется через fallback bridge...";
        const result = await invokeFallbackBridge("export", {
          sourceName: state.fileName || "table.xlsx",
          outputPath,
          cellWidthMm: state.layout.cellWidthMm,
          cellHeightMm: state.layout.cellHeightMm,
          matrix: state.matrix,
        });
        return {
          outputPath: normalizeWindowsPath(result.outputPath || outputPath),
          rows: Number(result.rows) || rows,
          cols: Number(result.cols) || cols,
        };
      }

      async function exportTable() {
        const { rows, cols } = currentDimensions();
        if (!rows || !cols) {
          throw new Error("XLSX matrix is empty.");
        }

        const status = await ensureActiveViewStatus();
        const outputPath = await resolveOutputPath();
        const outputDirectory = dirname(outputPath);
        if (!outputDirectory) {
          throw new Error("Output directory is empty.");
        }

        refs.resultBox.textContent = "Экспорт выполняется...";
        refs.exportButton.disabled = true;

        try {
          await context.ensureDirectory(outputDirectory);
          if (await context.fileExists(outputPath)) {
            await context.deleteFile(outputPath);
          }
          let exportResult;
          try {
            exportResult = await exportTableViaRuntime(status, outputPath, rows, cols);
          } catch (error) {
            if (!shouldUseFallbackBridge(error)) {
              throw error;
            }
            context.logger.info("export", "runtime create-table unsupported, switching to fallback bridge");
            exportResult = await exportTableViaFallback(outputPath, rows, cols);
          }

          state.lastBytes = await context.readFileBytes(exportResult.outputPath);
          state.lastExport = {
            outputPath: exportResult.outputPath,
            fileSize: state.lastBytes.length,
            rows: exportResult.rows,
            cols: exportResult.cols,
          };
          refs.resultBox.textContent = `OK | ${exportResult.outputPath} | ${state.lastBytes.length} bytes`;
          context.logger.info("export", exportResult.outputPath);
          await refreshStatus({ quiet: true });
        } finally {
          refs.exportButton.disabled = false;
          updateActionState();
        }
      }

      async function insertTableViaRuntime(status, tblPath) {
        const beforeExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.view-table-count",
          { handleId: status.viewHandleId },
          15000,
        );
        await context.executeCommand(
          "xlsx-to-kompas-tbl.insert-table",
          {
            handleId: status.viewHandleId,
            path: tblPath,
          },
          60000,
        );
        const afterExecution = await context.executeCommand(
          "xlsx-to-kompas-tbl.view-table-count",
          { handleId: status.viewHandleId },
          15000,
        );
        return {
          tableCountBefore: Number(beforeExecution.result) || 0,
          tableCountAfter: Number(afterExecution.result) || 0,
        };
      }

      async function insertTableViaFallback(tblPath) {
        refs.resultBox.textContent = "Вставка выполняется через fallback bridge...";
        const result = await invokeFallbackBridge("insert", { tblPath }, 120000);
        return {
          tableCountBefore: Number(result.tableCountBefore) || 0,
          tableCountAfter: Number(result.tableCountAfter) || 0,
        };
      }

      async function insertTable() {
        const status = await ensureActiveViewStatus();
        const tblPath = await resolveOutputPath();
        if (!await context.fileExists(tblPath)) {
          throw new Error(`Table file was not found: ${tblPath}`);
        }

        refs.resultBox.textContent = "Вставка выполняется...";
        let result;
        try {
          result = await insertTableViaRuntime(status, tblPath);
        } catch (error) {
          if (!shouldUseFallbackBridge(error)) {
            throw error;
          }
          context.logger.info("insert", "runtime DrawingTables path unsupported, switching to fallback bridge");
          result = await insertTableViaFallback(tblPath);
        }

        refs.resultBox.textContent = `Inserted | ${result.tableCountBefore} -> ${result.tableCountAfter}`;
        context.logger.info("insert", tblPath);
        await refreshStatus({ quiet: true });
      }

      async function downloadTable() {
        if ((!state.lastBytes || state.lastBytes.length === 0) && state.lastExport?.outputPath) {
          state.lastBytes = await context.readFileBytes(state.lastExport.outputPath);
        }
        if (!state.lastBytes || state.lastBytes.length === 0) {
          throw new Error("Нет экспортированного файла для загрузки.");
        }
        const fileName = state.lastExport?.outputPath
          ? state.lastExport.outputPath.split(/[\\/]/).pop()
          : `${sanitizeFileStem(state.fileName || "table")}.tbl`;
        context.downloadBytes(state.lastBytes, fileName || "table.tbl");
        context.logger.info("download", fileName || "table.tbl");
      }

      function resetState() {
        state.fileName = "";
        state.sheetName = "";
        state.matrix = [];
        state.lastExport = null;
        state.lastBytes = null;
        state.autoFollowOutput = true;
        refs.fileInput.value = "";
        refs.outputPath.value = "";
        refs.resultBox.textContent = "Ожидание данных.";
        renderAll();
      }

      async function handleWorkbookFile(file) {
        if (!file) {
          return;
        }
        const parsed = readWorkbookMatrix(file, await file.arrayBuffer());
        state.fileName = parsed.fileName;
        state.sheetName = parsed.sheetName;
        state.matrix = parsed.matrix;
        state.lastExport = null;
        state.lastBytes = null;
        state.autoFollowOutput = true;
        const { rows, cols } = currentDimensions();
        state.layout = reconcileLinkedLayout(state.layout, rows || 1, cols || 1, state.layoutDriver);
        persistLayout();
        refs.resultBox.textContent = "Матрица загружена. Можно запускать export.";
        context.logger.info("xlsx-parsed", `${state.fileName} ${formatSize(rows, cols)}`);
        await ensureTempPathLoaded();
        renderAll();
      }

      function bindLayoutInput(input, property, driver) {
        input.addEventListener("input", () => {
          applyLayoutChange(driver, {
            [property]: readNumeric(input, DEFAULT_LAYOUT[property]),
          });
        });
      }

      refs.fileInput.addEventListener("change", (event) => {
        handleWorkbookFile(event.target.files?.[0]).catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("xlsx-parse", refs.resultBox.textContent);
        });
      });

      refs.dropzone.addEventListener("dragenter", (event) => {
        event.preventDefault();
        state.dragDepth += 1;
        refs.dropzone.classList.add("is-dragging");
      });
      refs.dropzone.addEventListener("dragover", (event) => {
        event.preventDefault();
      });
      refs.dropzone.addEventListener("dragleave", (event) => {
        event.preventDefault();
        state.dragDepth = Math.max(0, state.dragDepth - 1);
        if (state.dragDepth === 0) {
          refs.dropzone.classList.remove("is-dragging");
        }
      });
      refs.dropzone.addEventListener("drop", (event) => {
        event.preventDefault();
        state.dragDepth = 0;
        refs.dropzone.classList.remove("is-dragging");
        const file = event.dataTransfer?.files?.[0];
        handleWorkbookFile(file).catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("drop", refs.resultBox.textContent);
        });
      });

      refs.outputPath.addEventListener("input", () => {
        state.autoFollowOutput = false;
        renderOutputMode();
        updateActionState();
      });

      refs.followButton.addEventListener("click", async () => {
        state.autoFollowOutput = true;
        await ensureTempPathLoaded();
        renderAll();
      });

      refs.exportButton.addEventListener("click", () => {
        exportTable().catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("export", refs.resultBox.textContent);
          updateActionState();
        });
      });

      refs.insertButton.addEventListener("click", () => {
        insertTable().catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("insert", refs.resultBox.textContent);
        });
      });

      refs.downloadButton.addEventListener("click", () => {
        downloadTable().catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("download", refs.resultBox.textContent);
        });
      });

      refs.resetButton.addEventListener("click", resetState);

      bindLayoutInput(refs.tableWidth, "tableWidthMm", "table");
      bindLayoutInput(refs.tableHeight, "tableHeightMm", "table");
      bindLayoutInput(refs.cellWidth, "cellWidthMm", "cell");
      bindLayoutInput(refs.cellHeight, "cellHeightMm", "cell");

      const onRuntimeLoaded = () => {
        ensureTempPathLoaded(true)
          .then(() => refreshStatus({ quiet: true }))
          .catch(() => {})
          .finally(() => {
            syncPolling();
          });
      };
      const onBridgeDisconnected = () => {
        state.status = null;
        syncPolling();
        renderAll();
      };
      const onVisibilityChange = () => {
        syncPolling();
      };

      context.events.addEventListener("runtime-loaded", onRuntimeLoaded);
      context.events.addEventListener("bridge-disconnected", onBridgeDisconnected);
      document.addEventListener("visibilitychange", onVisibilityChange);

      renderAll();

      return {
        activate() {
          state.active = true;
          syncPolling();
          if (context.getBridgeState().runtimeReady) {
            return refreshStatus({ quiet: true });
          }
          return Promise.resolve();
        },
        deactivate() {
          state.active = false;
          syncPolling();
          return Promise.resolve();
        },
        dispose() {
          if (state.pollTimer !== null) {
            window.clearInterval(state.pollTimer);
            state.pollTimer = null;
          }
          context.events.removeEventListener("runtime-loaded", onRuntimeLoaded);
          context.events.removeEventListener("bridge-disconnected", onBridgeDisconnected);
          document.removeEventListener("visibilitychange", onVisibilityChange);
        },
      };
    },
  };
}

export {
  DEFAULT_LAYOUT,
  EXPORT_BATCH_SIZE,
  buildAutoOutputPath,
  buildFallbackBridgeArguments,
  buildFallbackBridgeScriptPath,
  createCellWriteBatches,
  createXlsxToKompasTblModule,
  readWorkbookMatrix,
  reconcileLinkedLayout,
  shouldUseFallbackBridge,
};
