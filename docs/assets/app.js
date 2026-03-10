const DEFAULTS = {
  utilityUrl: "http://127.0.0.1:38741",
  pairingToken: "kompas-pages-local",
  profileId: "kompas-pages",
  clientName: "kompas-pages-ui",
  uiVersion: "1.0.0",
  defaultLayoutMode: "table",
};

const ui = {
  bridgeBadge: document.getElementById("bridge-badge"),
  kompasBadge: document.getElementById("kompas-badge"),
  bridgeMeta: document.getElementById("bridge-meta"),
  kompasMeta: document.getElementById("kompas-meta"),
  utilityUrl: document.getElementById("utility-url"),
  pairingToken: document.getElementById("pairing-token"),
  connectButton: document.getElementById("connect-button"),
  disconnectButton: document.getElementById("disconnect-button"),
  refreshStatusButton: document.getElementById("refresh-status-button"),
  xlsxFile: document.getElementById("xlsx-file"),
  fileName: document.getElementById("file-name"),
  sheetName: document.getElementById("sheet-name"),
  matrixSize: document.getElementById("matrix-size"),
  tableWidthMm: document.getElementById("table-width-mm"),
  tableHeightMm: document.getElementById("table-height-mm"),
  cellWidthMm: document.getElementById("cell-width-mm"),
  cellHeightMm: document.getElementById("cell-height-mm"),
  outputPath: document.getElementById("output-path"),
  layoutSummary: document.getElementById("layout-summary"),
  exportButton: document.getElementById("export-button"),
  downloadButton: document.getElementById("download-button"),
  resetButton: document.getElementById("reset-button"),
  resultBox: document.getElementById("result-box"),
  previewMeta: document.getElementById("preview-meta"),
  previewTable: document.getElementById("preview-table"),
  logOutput: document.getElementById("log-output"),
  clearLogButton: document.getElementById("clear-log-button"),
  tabs: Array.from(document.querySelectorAll(".rail__tab")),
  panels: Array.from(document.querySelectorAll(".tab")),
  modeButtons: Array.from(document.querySelectorAll(".mode-switch__button")),
};

const state = {
  layoutMode: DEFAULTS.defaultLayoutMode,
  workbookName: "",
  sheetName: "",
  matrix: [],
  bridge: null,
  exportResult: null,
  downloadBytes: null,
};

function setBridgeState(label, online = false) {
  ui.bridgeBadge.textContent = label;
  ui.bridgeBadge.className = online ? "badge" : "badge badge--dim";
}

function setKompasBadge(label, active = false) {
  ui.kompasBadge.textContent = label;
  ui.kompasBadge.className = active ? "badge" : "badge badge--dim";
}

function logLine(message, detail = "") {
  const stamp = new Date().toLocaleTimeString("ru-RU", { hour12: false });
  const suffix = detail ? ` ${detail}` : "";
  ui.logOutput.textContent += `\n[${stamp}] ${message}${suffix}`;
  ui.logOutput.scrollTop = ui.logOutput.scrollHeight;
}

function replaceLog(message) {
  ui.logOutput.textContent = message;
}

function readNumber(input) {
  const numeric = Number.parseFloat(input.value);
  return Number.isFinite(numeric) ? numeric : 0;
}

function querySettings() {
  const params = new URLSearchParams(window.location.search);
  return {
    utilityUrl: params.get("utilityUrl"),
    pairingToken: params.get("pairingToken"),
    autoConnect: params.get("autoConnect") === "1",
  };
}

function persistBridgeSettings() {
  window.localStorage.setItem(
    "kompas-pages.bridge",
    JSON.stringify({
      utilityUrl: ui.utilityUrl.value.trim(),
      pairingToken: ui.pairingToken.value.trim(),
    }),
  );
}

function restoreBridgeSettings() {
  const query = querySettings();
  const raw = window.localStorage.getItem("kompas-pages.bridge");
  const saved = raw ? JSON.parse(raw) : {};
  ui.utilityUrl.value = query.utilityUrl || saved.utilityUrl || DEFAULTS.utilityUrl;
  ui.pairingToken.value = query.pairingToken || saved.pairingToken || DEFAULTS.pairingToken;
  return query.autoConnect;
}

function switchTab(tabId) {
  for (const tab of ui.tabs) {
    tab.classList.toggle("is-active", tab.dataset.tab === tabId);
  }
  for (const panel of ui.panels) {
    panel.classList.toggle("tab--active", panel.dataset.panel === tabId);
  }
}

function switchMode(mode) {
  state.layoutMode = mode;
  for (const button of ui.modeButtons) {
    button.classList.toggle("is-active", button.dataset.mode === mode);
  }
  updateLayoutSummary();
}

function formatCellText(cell) {
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

function renderPreview() {
  const body = document.createElement("tbody");
  if (!state.matrix.length) {
    const row = document.createElement("tr");
    const cell = document.createElement("td");
    cell.className = "preview-table__empty";
    cell.textContent = "Загрузите XLSX.";
    row.append(cell);
    body.append(row);
    ui.previewTable.replaceChildren(body);
    return;
  }

  for (const rowData of state.matrix) {
    const row = document.createElement("tr");
    for (const value of rowData) {
      const cell = document.createElement("td");
      cell.textContent = value;
      row.append(cell);
    }
    body.append(row);
  }
  ui.previewTable.replaceChildren(body);
}

function updateLayoutSummary() {
  const rows = state.matrix.length;
  const cols = state.matrix[0]?.length ?? 0;
  if (!rows || !cols) {
    ui.layoutSummary.textContent = "Размеры будут рассчитаны после загрузки файла.";
    return null;
  }

  const tableWidthMm = readNumber(ui.tableWidthMm);
  const tableHeightMm = readNumber(ui.tableHeightMm);
  const cellWidthMm = readNumber(ui.cellWidthMm);
  const cellHeightMm = readNumber(ui.cellHeightMm);

  let effectiveCellWidth = cellWidthMm;
  let effectiveCellHeight = cellHeightMm;
  if (state.layoutMode === "table") {
    effectiveCellWidth = tableWidthMm / cols;
    effectiveCellHeight = tableHeightMm / rows;
  }

  const payload = {
    rows,
    cols,
    cellWidthMm: Number(effectiveCellWidth.toFixed(4)),
    cellHeightMm: Number(effectiveCellHeight.toFixed(4)),
    tableWidthMm: Number((effectiveCellWidth * cols).toFixed(4)),
    tableHeightMm: Number((effectiveCellHeight * rows).toFixed(4)),
  };

  ui.layoutSummary.textContent = [
    `mode=${state.layoutMode}`,
    `cell=${payload.cellWidthMm} x ${payload.cellHeightMm} mm`,
    `table=${payload.tableWidthMm} x ${payload.tableHeightMm} mm`,
  ].join(" · ");
  return payload;
}

function parseWorkbook(file, bytes) {
  const workbook = XLSX.read(bytes, {
    type: "array",
    cellNF: true,
    cellText: true,
    cellDates: false,
  });
  const firstSheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[firstSheetName];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const matrix = [];
  for (let rowIndex = range.s.r; rowIndex <= range.e.r; rowIndex += 1) {
    const row = [];
    for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
      const address = XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex });
      row.push(formatCellText(sheet[address]));
    }
    matrix.push(row);
  }

  state.workbookName = file.name;
  state.sheetName = firstSheetName;
  state.matrix = matrix;
  state.exportResult = null;
  state.downloadBytes = null;
  ui.downloadButton.disabled = true;
  ui.fileName.textContent = file.name;
  ui.sheetName.textContent = firstSheetName;
  ui.matrixSize.textContent = `${matrix.length} × ${matrix[0]?.length ?? 0}`;
  ui.previewMeta.textContent = `${firstSheetName} · ${matrix.length} × ${matrix[0]?.length ?? 0}`;
  ui.resultBox.textContent = "Матрица готова. Можно запускать экспорт.";
  renderPreview();
  updateLayoutSummary();
  logLine("xlsx parsed", `${file.name} ${matrix.length}x${matrix[0]?.length ?? 0}`);
}

function executionResult(execution) {
  return execution?.result ?? null;
}

function parseBridgeStdout(execution) {
  const processResult = executionResult(execution);
  if (!processResult?.hasExited) {
    throw new Error("bridge command did not exit");
  }
  const stdout = String(processResult.stdout || "").trim();
  if (!stdout) {
    if (processResult.exitCode === 0) {
      return {};
    }
    throw new Error(processResult.stderr || "bridge command returned empty stdout");
  }
  return JSON.parse(stdout);
}

class WebBridgeClient {
  constructor({ baseUrl, pairingToken }) {
    this.baseUrl = baseUrl.replace(/\/+$/, "");
    this.pairingToken = pairingToken;
    this.sessionId = null;
    this.socket = null;
    this.heartbeatTimer = null;
    this.heartbeatIntervalMs = 10000;
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
    return { response, payload };
  }

  async connect() {
    const { response, payload } = await this.fetchJson("/session/register", {
      method: "POST",
      body: JSON.stringify({
        clientName: DEFAULTS.clientName,
        uiVersion: DEFAULTS.uiVersion,
        desiredSessionId: this.sessionId,
      }),
    });
    if (!response.ok) {
      throw new Error(`session/register ${response.status}`);
    }
    const registered = payload?.payload ?? payload;
    this.sessionId = registered.sessionId;
    this.heartbeatIntervalMs = Math.max(1000, Number(registered.heartbeatIntervalSeconds || 10) * 1000);
    const wsUrl = `${registered.wsUrl}&token=${encodeURIComponent(this.pairingToken)}`;
    await this.#openSocket(wsUrl);
    return registered;
  }

  async disconnect(reason = "manual-close") {
    if (this.sessionId) {
      try {
        await this.fetchJson("/session/closing", {
          method: "POST",
          body: JSON.stringify({ sessionId: this.sessionId, reason }),
        });
      } catch {
        // best effort
      }
    }
    this.stopHeartbeat();
    if (this.socket) {
      try {
        this.socket.close(1000, reason);
      } catch {
        // ignore
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
      throw new Error("WebSocket closed");
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

  async execute(commandId, args = {}, timeoutMilliseconds = 60000) {
    const { response, payload } = await this.fetchJson("/commands/execute", {
      method: "POST",
      body: JSON.stringify({
        profileId: DEFAULTS.profileId,
        commandId,
        arguments: args,
        timeoutMilliseconds,
      }),
    });
    const execution = payload?.payload ?? payload;
    if (!response.ok || !execution?.success) {
      const rendered = execution?.error ? JSON.stringify(execution.error) : JSON.stringify(payload);
      throw new Error(`${commandId} failed: ${rendered}`);
    }
    return execution;
  }

  async #openSocket(wsUrl) {
    this.helloReceived = false;
    this.socket = new WebSocket(wsUrl);

    this.socket.addEventListener("message", (event) => {
      try {
        const envelope = JSON.parse(event.data);
        if (envelope.type === "hello") {
          this.helloReceived = true;
        }
      } catch {
        // ignore
      }
    });

    await new Promise((resolve, reject) => {
      const timeoutId = window.setTimeout(() => reject(new Error("ws hello timeout")), 20000);
      this.socket.addEventListener("open", () => {
        try {
          this.send("hello", {
            clientName: DEFAULTS.clientName,
            sessionId: this.sessionId,
          });
        } catch (error) {
          window.clearTimeout(timeoutId);
          reject(error);
        }
      });

      const poll = window.setInterval(() => {
        if (this.helloReceived) {
          window.clearInterval(poll);
          window.clearTimeout(timeoutId);
          this.stopHeartbeat();
          this.heartbeatTimer = window.setInterval(() => {
            try {
              this.send("heartbeat", { sessionId: this.sessionId });
            } catch {
              this.stopHeartbeat();
            }
          }, this.heartbeatIntervalMs);
          resolve();
        }
        if (this.socket?.readyState === WebSocket.CLOSED) {
          window.clearInterval(poll);
          window.clearTimeout(timeoutId);
          reject(new Error("ws closed before hello"));
        }
      }, 100);
    });
  }
}

async function connectBridge() {
  persistBridgeSettings();
  if (state.bridge) {
    await state.bridge.disconnect("reconnect");
  }
  state.bridge = new WebBridgeClient({
    baseUrl: ui.utilityUrl.value.trim(),
    pairingToken: ui.pairingToken.value.trim(),
  });
  setBridgeState("connect...");
  const registered = await state.bridge.connect();
  setBridgeState("online", true);
  ui.bridgeMeta.textContent = `session=${registered.sessionId} · heartbeat=${registered.heartbeatIntervalSeconds}s`;
  logLine("bridge connected", registered.sessionId);
  await refreshKompasStatus();
}

async function disconnectBridge() {
  if (!state.bridge) {
    return;
  }
  await state.bridge.disconnect();
  state.bridge = null;
  setBridgeState("offline");
  setKompasBadge("doc ?");
  ui.bridgeMeta.textContent = "Сессия не зарегистрирована.";
  ui.kompasMeta.textContent = "Статус KOMPAS не запрошен.";
  logLine("bridge disconnected");
}

async function refreshKompasStatus() {
  if (!state.bridge) {
    ui.kompasMeta.textContent = "Bridge не подключён.";
    return;
  }
  try {
    const execution = await state.bridge.execute("kompas.pages.status", {}, 30000);
    const status = parseBridgeStdout(execution);
    if (status.connected && status.hasActiveDocument) {
      setKompasBadge("doc ok", true);
      ui.kompasMeta.textContent = `${status.documentPath || status.documentName} · view=${status.viewName || "-"}`;
    } else if (status.connected) {
      setKompasBadge("doc none");
      ui.kompasMeta.textContent = "KOMPAS запущен, но активный 2D документ не открыт.";
    } else {
      setKompasBadge("kompas off");
      ui.kompasMeta.textContent = status.errorMessage || "KOMPAS не найден.";
    }
    logLine("kompas status", ui.kompasMeta.textContent);
  } catch (error) {
    setKompasBadge("status err");
    ui.kompasMeta.textContent = String(error.message || error);
    logLine("kompas status failed", ui.kompasMeta.textContent);
  }
}

async function loadDownloadBytes(path) {
  if (!state.bridge || !path) {
    return;
  }
  const existsExecution = await state.bridge.execute("system.file.exists", { path }, 15000);
  if (!executionResult(existsExecution)) {
    return;
  }
  const bytesExecution = await state.bridge.execute("system.file.read-bytes", { path }, 30000);
  const bytes = executionResult(bytesExecution);
  state.downloadBytes = new Uint8Array(Array.isArray(bytes) ? bytes : []);
  ui.downloadButton.disabled = state.downloadBytes.length === 0;
}

async function exportTable() {
  if (!state.bridge) {
    throw new Error("Bridge не подключён.");
  }
  if (!state.matrix.length) {
    throw new Error("XLSX не загружен.");
  }
  const layout = updateLayoutSummary();
  if (!layout || layout.cellWidthMm <= 0 || layout.cellHeightMm <= 0) {
    throw new Error("Некорректные размеры таблицы.");
  }

  const request = {
    sourceName: state.workbookName || "table.xlsx",
    outputPath: ui.outputPath.value.trim(),
    cellWidthMm: layout.cellWidthMm,
    cellHeightMm: layout.cellHeightMm,
    rows: layout.rows,
    cols: layout.cols,
    matrix: state.matrix,
  };

  ui.exportButton.disabled = true;
  ui.resultBox.textContent = "Экспорт выполняется...";
  logLine("export start", `${request.rows}x${request.cols}`);

  try {
    const execution = await state.bridge.execute("kompas.pages.export", {
      stdin: JSON.stringify(request),
    }, 180000);
    const processResult = executionResult(execution);
    const payload = parseBridgeStdout(execution);
    if (processResult.exitCode !== 0 || payload.success !== true) {
      const message = payload.errorMessage || processResult.stderr || "Экспорт завершился с ошибкой.";
      throw new Error(message);
    }

    state.exportResult = payload;
    await loadDownloadBytes(payload.outputPath);
    ui.resultBox.textContent = `OK · ${payload.outputPath} · ${payload.fileSize} bytes`;
    ui.outputPath.value = payload.outputPath;
    logLine("export done", `${payload.outputPath} ${payload.fileSize} bytes`);
    await refreshKompasStatus();
  } finally {
    ui.exportButton.disabled = false;
  }
}

function downloadTable() {
  if (!state.downloadBytes || !state.downloadBytes.length || !state.exportResult?.outputPath) {
    return;
  }
  const fileName = state.exportResult.outputPath.split(/[\\/]/).pop() || "table.tbl";
  const blob = new Blob([state.downloadBytes], { type: "application/octet-stream" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  link.click();
  URL.revokeObjectURL(url);
  logLine("download", fileName);
}

function resetState() {
  state.workbookName = "";
  state.sheetName = "";
  state.matrix = [];
  state.exportResult = null;
  state.downloadBytes = null;
  ui.xlsxFile.value = "";
  ui.fileName.textContent = "не выбран";
  ui.sheetName.textContent = "-";
  ui.matrixSize.textContent = "0 × 0";
  ui.previewMeta.textContent = "UsedRange первого листа";
  ui.resultBox.textContent = "Экспорт ещё не выполнялся.";
  ui.outputPath.value = "";
  ui.downloadButton.disabled = true;
  renderPreview();
  updateLayoutSummary();
  logLine("state reset");
}

function bindEvents() {
  ui.tabs.forEach((tab) => {
    tab.addEventListener("click", () => switchTab(tab.dataset.tab));
  });
  ui.modeButtons.forEach((button) => {
    button.addEventListener("click", () => switchMode(button.dataset.mode));
  });
  ui.xlsxFile.addEventListener("change", async (event) => {
    const file = event.target.files?.[0];
    if (!file) {
      return;
    }
    const bytes = await file.arrayBuffer();
    parseWorkbook(file, bytes);
  });
  [ui.tableWidthMm, ui.tableHeightMm, ui.cellWidthMm, ui.cellHeightMm].forEach((input) => {
    input.addEventListener("input", updateLayoutSummary);
  });
  ui.connectButton.addEventListener("click", async () => {
    try {
      await connectBridge();
    } catch (error) {
      setBridgeState("error");
      ui.bridgeMeta.textContent = String(error.message || error);
      logLine("bridge connect failed", ui.bridgeMeta.textContent);
    }
  });
  ui.disconnectButton.addEventListener("click", () => {
    disconnectBridge().catch(() => {});
  });
  ui.refreshStatusButton.addEventListener("click", () => {
    refreshKompasStatus().catch(() => {});
  });
  ui.exportButton.addEventListener("click", async () => {
    try {
      await exportTable();
    } catch (error) {
      ui.resultBox.textContent = String(error.message || error);
      logLine("export failed", ui.resultBox.textContent);
    }
  });
  ui.downloadButton.addEventListener("click", downloadTable);
  ui.resetButton.addEventListener("click", resetState);
  ui.clearLogButton.addEventListener("click", () => replaceLog("ready"));
  window.addEventListener("beforeunload", () => {
    if (state.bridge) {
      state.bridge.disconnect("page-close").catch(() => {});
    }
  });
}

function init() {
  bindEvents();
  switchTab("xlsx");
  switchMode(DEFAULTS.defaultLayoutMode);
  const autoConnect = restoreBridgeSettings();
  replaceLog("ready");
  setBridgeState("offline");
  setKompasBadge("doc ?");
  updateLayoutSummary();

  if (autoConnect) {
    connectBridge().catch((error) => {
      ui.bridgeMeta.textContent = String(error.message || error);
      logLine("autoconnect failed", ui.bridgeMeta.textContent);
    });
  }
}

init();
