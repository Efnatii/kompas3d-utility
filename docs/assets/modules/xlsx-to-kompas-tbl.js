const DEFAULT_LAYOUT = {
  mode: "table",
  tableWidthMm: 260,
  tableHeightMm: 24,
  cellWidthMm: 30,
  cellHeightMm: 8,
};

function stripExtension(fileName) {
  return String(fileName || "").replace(/\.[^.]+$/, "");
}

function dirname(filePath) {
  const value = String(filePath || "");
  const index = Math.max(value.lastIndexOf("\\"), value.lastIndexOf("/"));
  return index >= 0 ? value.slice(0, index) : "";
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

function parseProcessPayload(execution) {
  const process = execution?.result;
  if (!process?.hasExited) {
    throw new Error("bridge command did not exit");
  }

  const stdout = String(process.stdout || "").trim();
  let payload = {};
  if (stdout) {
    payload = JSON.parse(stdout);
  }

  if (process.exitCode !== 0) {
    throw new Error(payload.errorMessage || process.stderr || `bridge command failed with ${process.exitCode}`);
  }

  return { process, payload };
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

function createXlsxToKompasTblModule() {
  return {
    id: "xlsx-to-kompas-tbl",
    title: "XLSX to KOMPAS TBL",
    subtitle: "Первый лист XLSX парсится в браузере, команды export/insert выполняются через WebBridge.Utility.",
    tabLabel: "XLSX/TBL",
    tabDetail: "active",

    getRuntimeContribution(context) {
      const scriptPath = context.resolveWorkspacePath("scripts", "webbridge_xlsx_to_kompas_tbl.ps1");
      const workingDirectory = context.getWorkspaceRoot();
      return {
        commands: {
          "xlsx-to-kompas-tbl.status": context.dsl.createProcessCommand({
            scriptPath,
            action: "status",
            workingDirectory,
            timeoutMilliseconds: 30000,
          }),
          "xlsx-to-kompas-tbl.export": context.dsl.createProcessCommand({
            scriptPath,
            action: "export",
            workingDirectory,
            timeoutMilliseconds: 180000,
          }),
          "xlsx-to-kompas-tbl.insert": context.dsl.createProcessCommand({
            scriptPath,
            action: "insert",
            workingDirectory,
            timeoutMilliseconds: 60000,
          }),
          "xlsx-to-kompas-tbl.open-document": context.dsl.createProcessCommand({
            scriptPath,
            action: "open-document",
            workingDirectory,
            timeoutMilliseconds: 60000,
          }),
        },
        allowedTypes: [],
        allowedProcesses: ["powershell.exe"],
      };
    },

    mount(container, context) {
      const savedLayout = context.storage.get("layout", DEFAULT_LAYOUT) || DEFAULT_LAYOUT;
      const state = {
        fileName: "",
        sheetName: "",
        matrix: [],
        status: null,
        outputTouched: false,
        lastExport: null,
        lastBytes: null,
        mode: savedLayout.mode || DEFAULT_LAYOUT.mode,
        tableWidthMm: Number(savedLayout.tableWidthMm || DEFAULT_LAYOUT.tableWidthMm),
        tableHeightMm: Number(savedLayout.tableHeightMm || DEFAULT_LAYOUT.tableHeightMm),
        cellWidthMm: Number(savedLayout.cellWidthMm || DEFAULT_LAYOUT.cellWidthMm),
        cellHeightMm: Number(savedLayout.cellHeightMm || DEFAULT_LAYOUT.cellHeightMm),
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
                <h2>Source</h2>
                <div class="action-row">
                  <label class="button button--ghost" for="xlsx-file-input">Open XLSX</label>
                  <button class="button button--ghost" type="button" id="xlsx-refresh-status">Status</button>
                </div>
              </div>
              <input id="xlsx-file-input" type="file" accept=".xlsx,.xlsm,.xlsb,.xls" hidden>

              <div class="mode-switch" role="radiogroup" aria-label="Layout mode">
                <button class="mode-button" type="button" data-mode="table">table</button>
                <button class="mode-button" type="button" data-mode="cell">cell</button>
              </div>

              <div class="field-grid">
                <label class="field">
                  <span>table width, mm</span>
                  <input id="xlsx-table-width" type="number" step="0.1" min="0" value="${state.tableWidthMm}">
                </label>
                <label class="field">
                  <span>table height, mm</span>
                  <input id="xlsx-table-height" type="number" step="0.1" min="0" value="${state.tableHeightMm}">
                </label>
                <label class="field">
                  <span>cell width, mm</span>
                  <input id="xlsx-cell-width" type="number" step="0.1" min="0" value="${state.cellWidthMm}">
                </label>
                <label class="field">
                  <span>cell height, mm</span>
                  <input id="xlsx-cell-height" type="number" step="0.1" min="0" value="${state.cellHeightMm}">
                </label>
              </div>

              <label class="field">
                <span>output path</span>
                <input id="xlsx-output-path" type="text" spellcheck="false" placeholder="Blank = %TEMP%\\kompas-pages\\*.tbl">
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
        fileInput: container.querySelector("#xlsx-file-input"),
        fileName: container.querySelector("#xlsx-file-name"),
        sheetName: container.querySelector("#xlsx-sheet-name"),
        matrixSize: container.querySelector("#xlsx-matrix-size"),
        documentName: container.querySelector("#xlsx-document-name"),
        viewName: container.querySelector("#xlsx-view-name"),
        modeButtons: Array.from(container.querySelectorAll(".mode-button")),
        tableWidth: container.querySelector("#xlsx-table-width"),
        tableHeight: container.querySelector("#xlsx-table-height"),
        cellWidth: container.querySelector("#xlsx-cell-width"),
        cellHeight: container.querySelector("#xlsx-cell-height"),
        outputPath: container.querySelector("#xlsx-output-path"),
        layoutSummary: container.querySelector("#xlsx-layout-summary"),
        resultBox: container.querySelector("#xlsx-result-box"),
        refreshStatus: container.querySelector("#xlsx-refresh-status"),
        exportButton: container.querySelector("#xlsx-export-button"),
        insertButton: container.querySelector("#xlsx-insert-button"),
        downloadButton: container.querySelector("#xlsx-download-button"),
        resetButton: container.querySelector("#xlsx-reset-button"),
        previewMeta: container.querySelector("#xlsx-preview-meta"),
        previewTable: container.querySelector("#xlsx-preview-table"),
      };

      function persistLayout() {
        context.storage.set("layout", {
          mode: state.mode,
          tableWidthMm: readNumeric(refs.tableWidth, DEFAULT_LAYOUT.tableWidthMm),
          tableHeightMm: readNumeric(refs.tableHeight, DEFAULT_LAYOUT.tableHeightMm),
          cellWidthMm: readNumeric(refs.cellWidth, DEFAULT_LAYOUT.cellWidthMm),
          cellHeightMm: readNumeric(refs.cellHeight, DEFAULT_LAYOUT.cellHeightMm),
        });
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

      function computeLayout() {
        const rows = state.matrix.length;
        const cols = state.matrix[0]?.length || 0;
        if (!rows || !cols) {
          return null;
        }

        const tableWidthMm = readNumeric(refs.tableWidth, DEFAULT_LAYOUT.tableWidthMm);
        const tableHeightMm = readNumeric(refs.tableHeight, DEFAULT_LAYOUT.tableHeightMm);
        const cellWidthMm = readNumeric(refs.cellWidth, DEFAULT_LAYOUT.cellWidthMm);
        const cellHeightMm = readNumeric(refs.cellHeight, DEFAULT_LAYOUT.cellHeightMm);

        let effectiveCellWidth = cellWidthMm;
        let effectiveCellHeight = cellHeightMm;
        if (state.mode === "table") {
          effectiveCellWidth = tableWidthMm / cols;
          effectiveCellHeight = tableHeightMm / rows;
        }

        return {
          rows,
          cols,
          cellWidthMm: Number(effectiveCellWidth.toFixed(4)),
          cellHeightMm: Number(effectiveCellHeight.toFixed(4)),
          tableWidthMm: Number((effectiveCellWidth * cols).toFixed(4)),
          tableHeightMm: Number((effectiveCellHeight * rows).toFixed(4)),
        };
      }

      function syncModeButtons() {
        for (const button of refs.modeButtons) {
          button.classList.toggle("is-active", button.dataset.mode === state.mode);
        }
      }

      function suggestOutputPath() {
        if (state.outputTouched || !state.fileName) {
          return;
        }
        const documentPath = state.status?.documentPath || "";
        if (!documentPath) {
          refs.outputPath.value = "";
          return;
        }
        const targetDirectory = dirname(documentPath);
        if (!targetDirectory) {
          refs.outputPath.value = "";
          return;
        }
        refs.outputPath.value = `${targetDirectory}\\${stripExtension(state.fileName)}.tbl`;
      }

      function renderSummary() {
        const layout = computeLayout();
        if (!layout) {
          refs.layoutSummary.textContent = "Размеры будут рассчитаны после загрузки файла.";
          return;
        }
        refs.layoutSummary.textContent = [
          `mode=${state.mode}`,
          `cell=${formatMm(layout.cellWidthMm)} x ${formatMm(layout.cellHeightMm)} mm`,
          `table=${formatMm(layout.tableWidthMm)} x ${formatMm(layout.tableHeightMm)} mm`,
        ].join(" | ");
      }

      function renderStatus() {
        refs.fileName.textContent = state.fileName || "not loaded";
        refs.sheetName.textContent = state.sheetName || "-";
        refs.matrixSize.textContent = formatSize(state.matrix.length, state.matrix[0]?.length || 0);
        refs.documentName.textContent = state.status?.documentPath || state.status?.documentName || "no active doc";
        refs.viewName.textContent = state.status?.viewName || "-";
        refs.previewMeta.textContent = state.sheetName
          ? `${state.sheetName} | ${formatSize(state.matrix.length, state.matrix[0]?.length || 0)}`
          : "UsedRange первого листа.";
      }

      function updateActionState() {
        const bridgeState = context.getBridgeState();
        const hasMatrix = state.matrix.length > 0;
        const hasExport = Boolean(state.lastExport?.outputPath);
        const hasBytes = Boolean(state.lastBytes && state.lastBytes.length > 0);

        refs.exportButton.disabled = !bridgeState.runtimeReady || !hasMatrix;
        refs.insertButton.disabled = !bridgeState.runtimeReady || !hasExport;
        refs.downloadButton.disabled = !hasBytes;
      }

      function renderAll() {
        syncModeButtons();
        suggestOutputPath();
        renderSummary();
        renderStatus();
        renderPreview();
        updateActionState();
      }

      async function refreshStatus() {
        if (!context.getBridgeState().runtimeReady) {
          state.status = null;
          context.setModuleBadge("doc idle", false);
          context.setModuleMeta("Bridge подключён не полностью или runtime ещё не загружен.");
          renderAll();
          return;
        }

        const execution = await context.executeCommand("xlsx-to-kompas-tbl.status", {}, 30000);
        const { payload } = parseProcessPayload(execution);
        state.status = payload;

        if (payload.connected && payload.hasActiveDocument) {
          context.setModuleBadge("doc ok", true);
          context.setModuleMeta(`${payload.documentPath || payload.documentName} | view=${payload.viewName || "-"}`);
        } else if (payload.connected) {
          context.setModuleBadge("doc none", false);
          context.setModuleMeta("KOMPAS запущен, но активный 2D документ не открыт.");
        } else {
          context.setModuleBadge("kompas off", false);
          context.setModuleMeta(payload.errorMessage || "KOMPAS не найден.");
        }

        context.logger.info("status", context.getBridgeState().runtimeReady ? "runtime ready" : "runtime idle");
        renderAll();
      }

      async function exportTable() {
        if (!state.matrix.length) {
          throw new Error("XLSX matrix is empty.");
        }

        const layout = computeLayout();
        if (!layout || layout.cellWidthMm <= 0 || layout.cellHeightMm <= 0) {
          throw new Error("Некорректные размеры таблицы.");
        }

        refs.resultBox.textContent = "Экспорт выполняется...";
        refs.exportButton.disabled = true;

        try {
          const request = {
            sourceName: state.fileName || "table.xlsx",
            outputPath: refs.outputPath.value.trim(),
            cellWidthMm: layout.cellWidthMm,
            cellHeightMm: layout.cellHeightMm,
            matrix: state.matrix,
          };

          const execution = await context.executeCommand(
            "xlsx-to-kompas-tbl.export",
            { stdin: JSON.stringify(request) },
            180000,
          );
          const { payload } = parseProcessPayload(execution);
          if (payload.success !== true) {
            throw new Error(payload.errorMessage || "Экспорт завершился с ошибкой.");
          }

          state.lastExport = payload;
          state.outputTouched = true;
          refs.outputPath.value = payload.outputPath || refs.outputPath.value;
          state.lastBytes = await context.readFileBytes(payload.outputPath);
          refs.resultBox.textContent = `OK | ${payload.outputPath} | ${payload.fileSize} bytes`;
          context.logger.info("export", payload.outputPath);
          await refreshStatus();
        } finally {
          refs.exportButton.disabled = false;
          updateActionState();
        }
      }

      async function insertTable() {
        const tblPath = state.lastExport?.outputPath || refs.outputPath.value.trim();
        if (!tblPath) {
          throw new Error("Нет пути к .tbl для вставки.");
        }

        refs.resultBox.textContent = "Вставка выполняется...";
        const execution = await context.executeCommand(
          "xlsx-to-kompas-tbl.insert",
          { stdin: JSON.stringify({ tblPath }) },
          60000,
        );
        const { payload } = parseProcessPayload(execution);
        if (payload.success !== true) {
          throw new Error(payload.errorMessage || "Вставка завершилась с ошибкой.");
        }

        refs.resultBox.textContent = `Inserted | ${payload.tableCountBefore} -> ${payload.tableCountAfter}`;
        context.logger.info("insert", tblPath);
        await refreshStatus();
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
          : `${stripExtension(state.fileName || "table")}.tbl`;
        context.downloadBytes(state.lastBytes, fileName || "table.tbl");
        context.logger.info("download", fileName || "table.tbl");
      }

      function resetState() {
        state.fileName = "";
        state.sheetName = "";
        state.matrix = [];
        state.lastExport = null;
        state.lastBytes = null;
        state.outputTouched = false;
        refs.fileInput.value = "";
        refs.outputPath.value = "";
        refs.resultBox.textContent = "Ожидание данных.";
        renderAll();
      }

      async function handleFileSelect(event) {
        const file = event.target.files?.[0];
        if (!file) {
          return;
        }
        const parsed = readWorkbookMatrix(file, await file.arrayBuffer());
        state.fileName = parsed.fileName;
        state.sheetName = parsed.sheetName;
        state.matrix = parsed.matrix;
        state.lastExport = null;
        state.lastBytes = null;
        state.outputTouched = false;
        refs.resultBox.textContent = "Матрица загружена. Можно запускать export.";
        context.logger.info("xlsx-parsed", `${state.fileName} ${formatSize(state.matrix.length, state.matrix[0]?.length || 0)}`);
        renderAll();
      }

      function bindLayoutInput(input, property) {
        input.addEventListener("input", () => {
          state[property] = readNumeric(input, DEFAULT_LAYOUT[property]);
          persistLayout();
          renderAll();
        });
      }

      refs.fileInput.addEventListener("change", (event) => {
        handleFileSelect(event).catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("xlsx-parse", refs.resultBox.textContent);
        });
      });

      refs.outputPath.addEventListener("input", () => {
        state.outputTouched = true;
      });

      refs.refreshStatus.addEventListener("click", () => {
        refreshStatus().catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("status", refs.resultBox.textContent);
        });
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

      for (const button of refs.modeButtons) {
        button.addEventListener("click", () => {
          state.mode = button.dataset.mode;
          persistLayout();
          renderAll();
        });
      }

      bindLayoutInput(refs.tableWidth, "tableWidthMm");
      bindLayoutInput(refs.tableHeight, "tableHeightMm");
      bindLayoutInput(refs.cellWidth, "cellWidthMm");
      bindLayoutInput(refs.cellHeight, "cellHeightMm");

      const onRuntimeLoaded = () => {
        refreshStatus().catch((error) => {
          refs.resultBox.textContent = String(error.message || error);
          context.logger.error("runtime-loaded", refs.resultBox.textContent);
        });
      };
      const onBridgeDisconnected = () => {
        state.status = null;
        context.setModuleBadge("doc idle", false);
        context.setModuleMeta("Bridge отключён.");
        renderAll();
      };

      context.events.addEventListener("runtime-loaded", onRuntimeLoaded);
      context.events.addEventListener("bridge-disconnected", onBridgeDisconnected);

      renderAll();

      return {
        activate() {
          if (context.getBridgeState().runtimeReady) {
            return refreshStatus();
          }
          return Promise.resolve();
        },
        dispose() {
          context.events.removeEventListener("runtime-loaded", onRuntimeLoaded);
          context.events.removeEventListener("bridge-disconnected", onBridgeDisconnected);
        },
      };
    },
  };
}

export { createXlsxToKompasTblModule };
