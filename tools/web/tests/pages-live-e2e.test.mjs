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

test("live KOMPAS workflow and rich formatting scenarios pass on real bridge", { timeout: 15 * 60 * 1000 }, async (t) => {
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

  const workflow = report.scenarios.find((scenario) => scenario.name === "workflow");
  const richProof = report.scenarios.find((scenario) => scenario.name === "rich-proof");

  assert.ok(workflow, "workflow scenario report is missing");
  assert.ok(richProof, "rich-proof scenario report is missing");
  assert.match(String(workflow.inlineResultText || ""), /^Inline \| .+ \| anchor=/);
  assert.ok(Number(workflow.outputTableSize) > 0);
  assert.ok(Number(workflow.downloadedSize) > 0);
  assert.ok(Number(richProof.richOutputSize) > 0);
  assert.ok(Number(richProof.inlineArtifactSize) > 0);
  assert.equal(richProof.normalizedExportedA1?.lines?.[0]?.items?.[0]?.text, "Bold");
  assert.equal(richProof.normalizedExportedB1?.lines?.[0]?.items?.[0]?.text, "Plain styled");
});
