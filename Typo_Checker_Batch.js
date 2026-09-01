//
// Typo_Checker_Batch.js
//
// Compatibility launcher for the subscription-only, locally batched typo
// workflow. The former remote Batch API implementation was retired because it
// required a paid API key and pinned an older model. This launcher delegates to
// Typo_Checker.js in explicit standalone-input mode, which resolves the current
// official frontier before every query, uses xhigh reasoning, performs the
// mandatory independent recheck, and writes a complete JSON job record.
//
// Input:  one headerless XLSX column (column A)
// Output: three headerless columns: original text, corrected text, review note
//

import fs from "node:fs";
import path from "node:path";
import { spawnSync } from "node:child_process";

const SCRIPT_DIRECTORY = path.dirname(path.resolve(process.argv[1] || "."));
const CHECKER_PATH = path.join(SCRIPT_DIRECTORY, "Typo_Checker.js");
const INPUT_XLSX = path.resolve(process.env.INPUT_XLSX || "input_typos.xlsx");
const OUTPUT_XLSX = path.resolve(process.env.OUTPUT_XLSX || "output_typos.xlsx");
const JOB_RECORD_JSON = path.resolve(process.env.JOB_RECORD_JSON || `${OUTPUT_XLSX}.job.json`);
const SELF_TEST = /^(1|true|yes)$/i.test(process.env.TYPO_CHECKER_BATCH_SELF_TEST || "false");

function normalizedPath(filePath) {
  return path.resolve(filePath).replace(/[\\/]+$/, "").toLowerCase();
}

function assertDistinctPaths(artifacts) {
  const seen = new Map();
  for (const artifact of artifacts) {
    const resolved = normalizedPath(artifact.path);
    const existing = seen.get(resolved);
    if (existing) throw new Error(`${artifact.label} must not overwrite ${existing}: ${artifact.path}`);
    seen.set(resolved, artifact.label);
  }
}

function subscriptionEnvironment(source = process.env) {
  const child = { ...source };
  for (const name of Object.keys(child)) {
    if (
      (/^(?:OPENAI_|AZURE_OPENAI_)/i.test(name) && !/^OPENAI_MODEL$/i.test(name)) ||
      /^(?:CODEX_API_KEY|CODEX_ACCESS_TOKEN)$/i.test(name)
    ) {
      delete child[name];
    }
  }
  delete child.OUTPUT_JSON;
  delete child.TYPO_CHECKER_BATCH_SELF_TEST;
  return child;
}

function runLauncherSelfTest() {
  const cleaned = subscriptionEnvironment({
    OPENAI_MODEL: "gpt-test-frontier",
    OPENAI_API_KEY: "secret",
    OPENAI_FUTURE_CREDENTIAL: "secret",
    AZURE_OPENAI_ENDPOINT: "https://example.invalid",
    CODEX_ACCESS_TOKEN: "secret",
    SAFE_VALUE: "retained",
  });
  if (
    cleaned.OPENAI_MODEL !== "gpt-test-frontier" ||
    cleaned.OPENAI_API_KEY ||
    cleaned.OPENAI_FUTURE_CREDENTIAL ||
    cleaned.AZURE_OPENAI_ENDPOINT ||
    cleaned.CODEX_ACCESS_TOKEN ||
    cleaned.SAFE_VALUE !== "retained"
  ) {
    throw new Error("Subscription environment isolation self-test failed.");
  }
}

function runChecker(environment) {
  const run = spawnSync(process.execPath, [CHECKER_PATH], {
    env: environment,
    cwd: process.cwd(),
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
    windowsHide: true,
  });
  if (run.stdout) process.stdout.write(run.stdout);
  if (run.stderr) process.stderr.write(run.stderr);
  if (run.error || run.status !== 0) {
    throw new Error(`Primary typo checker failed with exit code ${run.status}: ${run.error?.message || "see diagnostics above"}`);
  }
}

function main() {
  if (process.argv.length > 2) {
    throw new Error("This script does not support positional arguments. Configure it with environment variables.");
  }
  if (!fs.existsSync(CHECKER_PATH) || !fs.statSync(CHECKER_PATH).isFile()) {
    throw new Error(`Primary typo checker not found: ${CHECKER_PATH}`);
  }

  const environment = subscriptionEnvironment();
  if (SELF_TEST) {
    runLauncherSelfTest();
    environment.TYPO_CHECKER_SELF_TEST = "true";
    runChecker(environment);
    return;
  }

  assertDistinctPaths([
    { label: "INPUT_XLSX", path: INPUT_XLSX },
    { label: "OUTPUT_XLSX", path: OUTPUT_XLSX },
    { label: "JOB_RECORD_JSON", path: JOB_RECORD_JSON },
  ]);
  environment.INPUT_XLSX = INPUT_XLSX;
  environment.OUTPUT_XLSX = OUTPUT_XLSX;
  environment.JOB_RECORD_JSON = JOB_RECORD_JSON;
  environment.ALLOW_STANDALONE_INPUT = "true";
  runChecker(environment);
}

try {
  main();
} catch (error) {
  console.error("Fatal error:", error.message);
  process.exitCode = 1;
}
