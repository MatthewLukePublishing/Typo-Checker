import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const tests = [
  ["Typo_Checker.js", "TYPO_CHECKER_SELF_TEST"],
  ["Typo_Checker_Batch.js", "TYPO_CHECKER_BATCH_SELF_TEST"],
  ["Compare_5_Columns.js", "COMPARE_5_COLUMNS_SELF_TEST"],
  ["5_Columns Rechecker.js", "FIVE_COLUMNS_RECHECKER_SELF_TEST"],
  ["Style_Guide_Checker.js", "STYLE_GUIDE_CHECKER_SELF_TEST"],
];

for (const [scriptName, environmentName] of tests) {
  process.stdout.write(`\n== ${scriptName} ==\n`);
  const run = spawnSync(process.execPath, [path.join(scriptDirectory, scriptName)], {
    cwd: scriptDirectory,
    env: { ...process.env, [environmentName]: "true" },
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
    windowsHide: true,
  });

  if (run.stdout) process.stdout.write(run.stdout);
  if (run.stderr) process.stderr.write(run.stderr);
  if (run.error || run.status !== 0) {
    console.error(`${scriptName} failed with exit code ${run.status}: ${run.error?.message || "see diagnostics above"}`);
    process.exit(run.status || 1);
  }
}

if (process.platform === "win32") {
  process.stdout.write("\n== Run_Latest_Typo_Checker.ps1 ==\n");
  const powershellPath = path.join(
    process.env.SystemRoot || "C:\\Windows",
    "System32",
    "WindowsPowerShell",
    "v1.0",
    "powershell.exe",
  );
  if (!fs.existsSync(powershellPath)) {
    console.error(`Windows PowerShell was not found: ${powershellPath}`);
    process.exit(1);
  }
  const run = spawnSync(powershellPath, [
    "-NoLogo",
    "-NoProfile",
    "-NonInteractive",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    path.join(scriptDirectory, "Run_Latest_Typo_Checker.ps1"),
    "-SelfTest",
  ], {
    cwd: scriptDirectory,
    encoding: "utf8",
    maxBuffer: 4 * 1024 * 1024,
    windowsHide: true,
  });
  if (run.stdout) process.stdout.write(run.stdout);
  if (run.stderr) process.stderr.write(run.stderr);
  if (run.error || run.status !== 0) {
    console.error(`Run_Latest_Typo_Checker.ps1 failed with exit code ${run.status}: ${run.error?.message || "see diagnostics above"}`);
    process.exit(run.status || 1);
  }
}

console.log("\nAll offline self-tests passed.");
