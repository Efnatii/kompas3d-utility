import test from "node:test";
import assert from "node:assert/strict";
import path from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const webRoot = path.resolve(__dirname, "..");
const runE2ePath = path.join(webRoot, "run_e2e.mjs");

function runLiveE2e(args = [], env = {}) {
  return new Promise((resolve, reject) => {
    const child = spawn(
      process.execPath,
      [runE2ePath, ...args],
      {
        cwd: webRoot,
        env: {
          ...process.env,
          ...env,
        },
        windowsHide: true,
        stdio: ["ignore", "pipe", "pipe"],
      },
    );

    let stdout = "";
    let stderr = "";
    child.stdout.on("data", (chunk) => {
      stdout += String(chunk);
    });
    child.stderr.on("data", (chunk) => {
      stderr += String(chunk);
    });
    child.on("error", reject);
    child.on("exit", (code) => {
      if (code !== 0) {
        reject(new Error(`run_e2e exited with code ${code}\nSTDOUT:\n${stdout}\nSTDERR:\n${stderr}`));
        return;
      }
      try {
        resolve(JSON.parse(stdout));
      } catch (error) {
        reject(new Error(`Failed to parse run_e2e stdout as JSON: ${error}\nSTDOUT:\n${stdout}\nSTDERR:\n${stderr}`));
      }
    });
  });
}

test("live KOMPAS workflow, rich formatting, auto-fit, and command surface scenarios pass on real bridge", { timeout: 60 * 60 * 1000 }, async (t) => {
  if (process.env.KOMPAS_LIVE_E2E !== "1") {
    t.skip("Set KOMPAS_LIVE_E2E=1 to run live WebBridge/KOMPAS end-to-end proof.");
    return;
  }

  const report = await runLiveE2e([
    "--scenario",
    "all",
    "--utility-config-mode",
    "development-flat",
  ]);

  assert.equal(report.success, true);
  assert.equal(report.scenario, "all");
  assert.ok(Array.isArray(report.scenarios));
  assert.equal(report.openedDocumentHandleState?.active, true);
  assert.equal(report.openedDocumentHandleState?.visible, true);

  const workflow = report.scenarios.find((scenario) => scenario.name === "workflow");
  const richProof = report.scenarios.find((scenario) => scenario.name === "rich-proof");
  const autoFitProof = report.scenarios.find((scenario) => scenario.name === "autofit-proof");
  const commandProof = report.scenarios.find((scenario) => scenario.name === "command-proof");

  assert.ok(workflow, "workflow scenario report is missing");
  assert.ok(richProof, "rich-proof scenario report is missing");
  assert.ok(autoFitProof, "autofit-proof scenario report is missing");
  assert.ok(commandProof, "command-proof scenario report is missing");
  assert.match(String(workflow.inlineResultText || ""), /^Inline \| no \.tbl \| handle=.+ ref=\d+ \| anchor=/);
  assert.ok(String(workflow.inlineHandleId || "").length > 0);
  assert.ok(Number(workflow.inlineTableReference) > 0);
  assert.ok(Number(workflow.outputTableSize) > 0);
  assert.ok(Number(workflow.downloadedSize) > 0);
  assert.ok(Number(richProof.richOutputSize) > 0);
  assert.ok(String(richProof.inlineHandleId || "").length > 0);
  assert.ok(Number(richProof.inlineTableReference) > 0);
  assert.equal(richProof.styleCaseCount, 25);
  assert.equal(richProof.expectedAddresses?.length, 25);
  assert.equal(richProof.expectedAddresses?.[0], "A1");
  assert.equal(richProof.expectedAddresses?.at(-1), "E5");
  assert.equal(richProof.styleCases?.find((styleCase) => styleCase.address === "C2")?.label, "tahoma-wrap-flag");
  assert.equal(richProof.normalizedExportedSnapshots?.A1?.lines?.[0]?.items?.[0]?.text, "Bold");
  assert.equal(richProof.normalizedExportedSnapshots?.B1?.lines?.[0]?.items?.[0]?.text, "Plain styled");
  assert.equal(richProof.normalizedExportedSnapshots?.C1?.lines?.[0]?.items?.[0]?.fontName, "Courier New");
  assert.equal(richProof.normalizedExportedSnapshots?.D1?.lines?.[0]?.items?.[0]?.fontName, "Times New Roman");
  assert.equal(richProof.normalizedExportedSnapshots?.E1?.lines?.[0]?.items?.[0]?.color, 0x000000);
  assert.equal(richProof.normalizedExportedSnapshots?.C2?.oneLine, false);
  assert.equal(richProof.normalizedExportedSnapshots?.C2?.lineCount, 1);
  assert.equal(richProof.normalizedExportedSnapshots?.E2?.lines?.[0]?.items?.[0]?.color, 0xFF0000);
  assert.equal(richProof.normalizedInlineSnapshots?.B4?.text, "A  B  C");
  assert.equal(richProof.normalizedInlineSnapshots?.A5?.lines?.[1]?.items?.[0]?.text, "OK");
  assert.equal(richProof.normalizedInlineSnapshots?.D5?.lines?.[2]?.items?.[0]?.text, "Bot");
  assert.equal(richProof.normalizedInlineSnapshots?.E5?.lines?.[0]?.items?.[0]?.text, "Final case");
  assert.equal(autoFitProof.baselineCellSize?.autoFitEnabled, false);
  assert.equal(autoFitProof.shrinkCellSize?.autoFitEnabled, true);
  assert.equal(autoFitProof.growCellSize?.autoFitEnabled, true);
  assert.match(String(autoFitProof.baselineSummaryText || ""), /text=excel 1:1/);
  assert.match(String(autoFitProof.shrinkSummaryText || ""), /text=auto-fit/);
  assert.match(String(autoFitProof.growSummaryText || ""), /text=auto-fit/);
  assert.ok(Number(autoFitProof.shrinkAdjustedCellCount) > 0);
  assert.ok(Number(autoFitProof.growAdjustedCellCount) > 0);
  assert.ok(Array.isArray(autoFitProof.trackedAddresses));
  assert.equal(autoFitProof.trackedAddresses.length, 25);
  assert.ok(Array.isArray(autoFitProof.provenShrinkAddresses));
  assert.ok(Array.isArray(autoFitProof.provenGrowAddresses));
  assert.ok(autoFitProof.provenShrinkAddresses.includes("A1"));
  assert.ok(autoFitProof.provenGrowAddresses.includes("E1"));
  assert.ok(autoFitProof.provenShrinkAddresses.length >= 5);
  assert.ok(autoFitProof.provenGrowAddresses.length >= 5);
  assert.ok(Number(autoFitProof.baselineHeightMetrics?.A1?.maxItemHeightMm) > Number(autoFitProof.shrinkHeightMetrics?.A1?.maxItemHeightMm));
  assert.ok(Number(autoFitProof.growHeightMetrics?.E1?.maxItemHeightMm) > Number(autoFitProof.baselineHeightMetrics?.E1?.maxItemHeightMm));
  assert.equal(autoFitProof.normalizedBaselineSnapshots?.A1?.text, autoFitProof.normalizedShrinkSnapshots?.A1?.text);
  assert.equal(autoFitProof.normalizedBaselineSnapshots?.A1?.text, autoFitProof.normalizedGrowSnapshots?.A1?.text);
  assert.equal(autoFitProof.normalizedBaselineSnapshots?.E1?.text, autoFitProof.normalizedGrowSnapshots?.E1?.text);
  assert.equal(commandProof.iterationCount, 10);
  assert.deepEqual(commandProof.missingCommands, []);
  assert.ok(Array.isArray(commandProof.sampleIterations));
  assert.equal(commandProof.sampleIterations.length, 10);
  assert.equal(commandProof.sampleIterations?.[0]?.tableCountBefore, 0);
  assert.equal(commandProof.sampleIterations?.[0]?.tableCountAfter, 0);
  assert.equal(commandProof.sampleIterations?.[0]?.api5CellAddress, "A1");
  assert.equal(commandProof.sampleIterations?.[0]?.api5CellText, "api5-1");
  assert.ok(Number(commandProof.trackedCommandCounts?.["xlsx-to-kompas-tbl.create-table"]) >= 10);
  assert.ok(Number(commandProof.trackedCommandCounts?.["xlsx-to-kompas-tbl.table-cell-set-item"]) >= 10);
  assert.ok(Number(commandProof.trackedCommandCounts?.["xlsx-to-kompas-tbl.api5-document-set-table-cell-text"]) >= 10);
  assert.ok(Number(commandProof.trackedCommandCounts?.["xlsx-to-kompas-tbl.api5-document-clear-table-cell-text"]) >= 10);
  assert.ok(Number(commandProof.trackedCommandCounts?.["xlsx-to-kompas-tbl.open-document"]) >= 10);
});
