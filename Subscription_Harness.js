import crypto from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";

const MODEL_POLICY = "official_latest_frontier";
const REASONING_EFFORT = "xhigh";
const FORCED_LOGIN_METHOD = "chatgpt";
const OFFICIAL_LATEST_MODEL_URL = "https://developers.openai.com/api/docs/guides/latest-model.md";
const OFFICIAL_BASE_URL = "https://developers.openai.com";
const OFFICIAL_MODEL_HOST = new URL(OFFICIAL_BASE_URL).hostname;
const OFFICIAL_CLI_TAIL = path.normalize(path.join("@openai", "codex", "bin", "codex.js")).toLowerCase();
const CODEX_QUERY_TIMEOUT_MS = parsePositiveInteger(
  process.env.CODEX_QUERY_TIMEOUT_MS || 1_800_000,
  "CODEX_QUERY_TIMEOUT_MS",
  { maximum: 7_200_000 },
);
const CODEX_LOGIN_TIMEOUT_MS = 30_000;

function normalizedPath(filePath) {
  return path.resolve(String(filePath || "")).replace(/[\\/]+$/, "").toLowerCase();
}

export function assertDistinctPaths(artifacts) {
  const seen = new Map();
  for (const artifact of artifacts) {
    const resolved = normalizedPath(artifact.path);
    if (!resolved) throw new Error(`${artifact.label} path is blank.`);
    const existing = seen.get(resolved);
    if (existing) throw new Error(`${artifact.label} must not overwrite ${existing}: ${artifact.path}`);
    seen.set(resolved, artifact.label);
  }
}

export function parsePositiveInteger(value, label, options = {}) {
  const parsed = Number(value);
  const maximum = options.maximum ?? Number.MAX_SAFE_INTEGER;
  if (!Number.isSafeInteger(parsed) || parsed < 1 || parsed > maximum) {
    throw new Error(`${label} must be an integer from 1 through ${maximum}; received '${value}'.`);
  }
  return parsed;
}

export function safeWorksheetName(value, fallback = "Results") {
  const clean = String(value || "")
    .replace(/[\u0000-\u001F\\/:*?\[\]]/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .replace(/^'+|'+$/g, "");
  const safe = clean || fallback;
  return safe.slice(0, 31).replace(/^'+|'+$/g, "") || "Results";
}

export function resultWorksheetName(sourceName, suffix) {
  const cleanSuffix = String(suffix || "")
    .replace(/[\u0000-\u001F\\/:*?\[\]]/g, "")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 30);
  const source = safeWorksheetName(sourceName);
  if (!cleanSuffix) return source;
  return `${source.slice(0, 31 - cleanSuffix.length)}${cleanSuffix}`;
}

function writeJsonAtomic(filePath, value) {
  const resolved = path.resolve(filePath);
  fs.mkdirSync(path.dirname(resolved), { recursive: true });
  const suffix = crypto.randomBytes(6).toString("hex");
  const temporary = `${resolved}.${process.pid}.${suffix}.tmp`;
  try {
    fs.writeFileSync(temporary, `${JSON.stringify(value, null, 2)}\n`, "utf8");
    fs.renameSync(temporary, resolved);
  } finally {
    if (fs.existsSync(temporary)) fs.unlinkSync(temporary);
  }
}

export function parseLatestModelInfo(markdown) {
  const lines = String(markdown || "").split(/\r?\n/);
  const start = lines.findIndex((line) => /^latestModelInfo:\s*$/.test(line));
  if (start >= 0) {
    const info = {};
    for (let index = start + 1; index < lines.length; index += 1) {
      if (!lines[index].trim()) continue;
      const match = lines[index].match(/^ {2}([A-Za-z][A-Za-z0-9_-]*):\s*(.+?)\s*$/);
      if (!match) break;
      info[match[1]] = match[2].replace(/^["']|["']$/g, "");
    }
    return info;
  }
  const comment = String(markdown || "").match(/<!--\s*latestModelInfo\s*\n([\s\S]*?)\n\s*-->/m);
  if (!comment) return undefined;
  const info = {};
  for (const line of comment[1].split(/\r?\n/)) {
    const match = line.match(/^\s*([A-Za-z][A-Za-z0-9_-]*):\s*(.+?)\s*$/);
    if (match) info[match[1]] = match[2].replace(/^["']|["']$/g, "");
  }
  return info;
}

function assertOfficialModelResponseUrl(value) {
  const resolved = new URL(String(value || ""));
  if (resolved.protocol !== "https:" || resolved.hostname !== OFFICIAL_MODEL_HOST) {
    throw new Error(`official latest-model request redirected outside ${OFFICIAL_MODEL_HOST}`);
  }
  return resolved.toString();
}

async function fetchOfficialLatestModelMarkdown() {
  let lastError;
  for (let attempt = 1; attempt <= 3; attempt += 1) {
    const controller = new AbortController();
    const timeout = setTimeout(() => controller.abort(), 15_000);
    try {
      const response = await fetch(OFFICIAL_LATEST_MODEL_URL, {
        headers: { accept: "text/markdown,text/plain,*/*" },
        signal: controller.signal,
      });
      const sourceUrl = assertOfficialModelResponseUrl(response.url);
      if (response.ok) return { markdown: await response.text(), sourceUrl };
      lastError = new Error(`official latest-model request returned HTTP ${response.status}`);
      if (response.status < 500 && response.status !== 429) break;
    } catch (error) {
      lastError = error;
    } finally {
      clearTimeout(timeout);
    }
    if (attempt < 3) await new Promise((resolve) => setTimeout(resolve, 250 * attempt));
  }
  throw lastError || new Error("official latest-model request failed");
}

export async function resolveLatestSubscriptionModel() {
  const { markdown, sourceUrl } = await fetchOfficialLatestModelMarkdown();
  const info = parseLatestModelInfo(markdown);
  const model = String(info?.model || "").trim();
  const migrationGuide = String(info?.migrationGuide || "").trim();
  const promptingGuide = String(info?.promptingGuide || "").trim();
  if (!/^gpt-[a-z0-9.-]+$/i.test(model) || !migrationGuide || !promptingGuide) {
    throw new Error("official latestModelInfo is missing a valid model or guide reference");
  }
  return {
    schemaVersion: 1,
    policy: MODEL_POLICY,
    model,
    resolvedAt: new Date().toISOString(),
    sourceUrl,
    migrationGuideUrl: new URL(migrationGuide, OFFICIAL_BASE_URL).toString(),
    promptingGuideUrl: new URL(promptingGuide, OFFICIAL_BASE_URL).toString(),
  };
}

function codexEnvironment(source = process.env) {
  const childEnvironment = { ...source };
  for (const name of Object.keys(childEnvironment)) {
    if (
      /^(?:OPENAI_|AZURE_OPENAI_)/i.test(name) ||
      /^(?:CODEX_API_KEY|CODEX_ACCESS_TOKEN|AZURE_OPENAI_API_KEY)$/i.test(name)
    ) {
      delete childEnvironment[name];
    }
  }
  return childEnvironment;
}

function resolveCodexRuntime() {
  const nodePath = process.env.CODEX_NODE_EXE || process.execPath;
  const candidates = [
    process.env.CODEX_CLI_JS,
    process.env.APPDATA
      ? path.join(process.env.APPDATA, "npm", "node_modules", "@openai", "codex", "bin", "codex.js")
      : "",
  ].filter(Boolean);
  const cliPath = candidates.find((candidate) => fs.existsSync(candidate));
  if (!fs.existsSync(nodePath)) throw new Error(`Codex Node runtime was not found: ${nodePath}`);
  if (!cliPath) throw new Error("The official standalone Codex CLI is not installed.");
  if (!path.normalize(cliPath).toLowerCase().endsWith(OFFICIAL_CLI_TAIL)) {
    throw new Error(`Refusing non-official Codex CLI path: ${cliPath}`);
  }
  return { nodePath, cliPath };
}

function runCodex(runtime, args, options = {}) {
  return spawnSync(runtime.nodePath, [runtime.cliPath, ...args], {
    cwd: options.cwd,
    env: codexEnvironment(),
    input: options.input,
    encoding: "utf8",
    maxBuffer: 64 * 1024 * 1024,
    timeout: options.timeoutMs || CODEX_QUERY_TIMEOUT_MS,
    windowsHide: true,
  });
}

function assertChatGptLogin(runtime) {
  const status = runCodex(runtime, [
    "-c",
    `forced_login_method="${FORCED_LOGIN_METHOD}"`,
    "login",
    "status",
  ], { timeoutMs: CODEX_LOGIN_TIMEOUT_MS });
  const output = `${status.stdout || ""}\n${status.stderr || ""}`;
  if (status.error || status.status !== 0 || !/Logged in using ChatGPT/i.test(output)) {
    throw new Error("This workflow requires Codex to be logged in using ChatGPT. API-key authentication is prohibited.");
  }
}

function assertCurrentModel(resolution, configuredModel, expectedModel) {
  if (!resolution || resolution.policy !== MODEL_POLICY || !resolution.model) {
    throw new Error("Official latest-model resolution returned an invalid or blank result.");
  }
  if (configuredModel && configuredModel !== resolution.model) {
    throw new Error(
      `Configured model '${configuredModel}' is outdated; the current official frontier model is '${resolution.model}'.`,
    );
  }
  if (expectedModel && expectedModel !== resolution.model) {
    throw new Error(
      `This job recorded model '${expectedModel}', but the current official frontier model is '${resolution.model}'. Start a new job; no stale-model fallback is permitted.`,
    );
  }
}

function safeQueryId(value) {
  const queryId = String(value || "").trim();
  if (!/^[A-Za-z0-9][A-Za-z0-9_-]{0,99}$/.test(queryId)) {
    throw new Error(`Invalid query ID: ${queryId || "(blank)"}`);
  }
  return queryId;
}

function assertJobCanComplete(job, queryFailure) {
  if (queryFailure) {
    throw new Error(`Cannot complete a job after a query failure: ${queryFailure.message}`);
  }
  const incomplete = job?.queries?.find((query) => query.status !== "completed");
  if (incomplete) {
    throw new Error(`Cannot complete a job while query '${incomplete.queryId}' is '${incomplete.status}'.`);
  }
}

class SubscriptionHarness {
  constructor(options) {
    this.jobType = String(options.jobType || "subscription_workflow");
    this.jobRecordPath = path.resolve(options.jobRecordPath);
    this.configuredModel = String(options.configuredModel || "").trim();
    this.context = options.context || {};
    this.protectedPaths = options.protectedPaths || [];
    assertDistinctPaths([
      ...this.protectedPaths,
      { label: "The subscription job record", path: this.jobRecordPath },
    ]);
    this.runtime = null;
    this.tempRoot = null;
    this.job = null;
    this.initialization = null;
    this.queryQueue = Promise.resolve();
    this.queryFailure = null;
  }

  initialize() {
    if (!this.initialization) this.initialization = this.#initialize();
    return this.initialization;
  }

  async #initialize() {
    this.runtime = resolveCodexRuntime();
    assertChatGptLogin(this.runtime);
    const resolution = await resolveLatestSubscriptionModel();
    assertCurrentModel(resolution, this.configuredModel);
    this.tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), "typos-subscription-"));
    const now = new Date().toISOString();
    this.job = {
      schemaVersion: 1,
      jobType: this.jobType,
      status: "running",
      provider: "CodexSubscription",
      executionTransport: "CodexCliHarness",
      authentication: "ChatGPT",
      authenticationPolicy: "chatgpt_subscription_only",
      model: resolution.model,
      modelPolicy: MODEL_POLICY,
      modelResolution: resolution,
      configuredModel: this.configuredModel || null,
      reasoningEffort: REASONING_EFFORT,
      queryTimeoutMs: CODEX_QUERY_TIMEOUT_MS,
      context: this.context,
      queries: [],
      createdAt: now,
      updatedAt: now,
    };
    writeJsonAtomic(this.jobRecordPath, this.job);
    return this.job;
  }

  query(options) {
    const operation = this.queryQueue.then(() => {
      if (this.queryFailure) {
        throw new Error(`A prior subscription query failed: ${this.queryFailure.message}`);
      }
      return this.#query(options);
    });
    const tracked = operation.catch((error) => {
      if (!this.queryFailure) this.queryFailure = error;
      throw error;
    });
    this.queryQueue = tracked.catch(() => undefined);
    return tracked;
  }

  #markQueryFailed(queryRecord, error) {
    queryRecord.status = "failed";
    queryRecord.error = String(error?.message || error).slice(0, 12_000);
    queryRecord.completedAt = new Date().toISOString();
    this.job.updatedAt = queryRecord.completedAt;
    writeJsonAtomic(this.jobRecordPath, this.job);
  }

  async #query(options) {
    await this.initialize();
    const queryId = safeQueryId(options.queryId);
    const prompt = String(options.prompt || "");
    if (!prompt.trim()) throw new Error(`Prompt is blank for ${queryId}.`);
    if (!options.schema || typeof options.schema !== "object") {
      throw new Error(`Output schema is missing for ${queryId}.`);
    }
    if (this.job.queries.some((query) => query.queryId === queryId)) {
      throw new Error(`Duplicate query ID: ${queryId}`);
    }

    assertChatGptLogin(this.runtime);
    let resolution;
    try {
      resolution = await resolveLatestSubscriptionModel();
    } catch (error) {
      throw new Error(`Official frontier-model resolution failed before ${queryId}: ${error.message}`);
    }
    assertCurrentModel(resolution, this.configuredModel, this.job.model);

    const schemaPath = path.join(this.tempRoot, `${queryId}.schema.json`);
    const responsePath = path.join(this.tempRoot, `${queryId}.response.json`);
    writeJsonAtomic(schemaPath, options.schema);
    const queryRecord = {
      queryId,
      model: resolution.model,
      modelPolicy: resolution.policy,
      modelResolvedAt: resolution.resolvedAt,
      modelResolutionSourceUrl: resolution.sourceUrl,
      reasoningEffort: REASONING_EFFORT,
      metadata: options.metadata || {},
      status: "sending",
    };
    this.job.queries.push(queryRecord);
    this.job.updatedAt = new Date().toISOString();
    writeJsonAtomic(this.jobRecordPath, this.job);

    let response;
    try {
      const run = runCodex(this.runtime, [
        "exec",
        "-",
        "--model",
        queryRecord.model,
        "-c",
        `forced_login_method="${FORCED_LOGIN_METHOD}"`,
        "-c",
        `model_reasoning_effort="${REASONING_EFFORT}"`,
        "--strict-config",
        "--sandbox",
        "read-only",
        "--skip-git-repo-check",
        "--ephemeral",
        "--ignore-user-config",
        "--ignore-rules",
        "--output-schema",
        schemaPath,
        "--output-last-message",
        responsePath,
        "--cd",
        this.tempRoot,
        "--color",
        "never",
      ], { cwd: this.tempRoot, input: prompt });

      if (run.error || run.status !== 0) {
        const diagnostics = `${run.error?.message || ""}\n${run.stdout || ""}\n${run.stderr || ""}`.slice(-12_000);
        throw new Error(`Codex subscription query failed for ${queryId} with exit code ${run.status}.\n${diagnostics}`);
      }
      if (!fs.existsSync(responsePath)) throw new Error(`Codex did not write a response for ${queryId}.`);
      try {
        response = JSON.parse(fs.readFileSync(responsePath, "utf8").replace(/^\uFEFF/, ""));
      } catch (error) {
        throw new Error(`Could not parse the Codex response for ${queryId}: ${error.message}`);
      }
    } catch (error) {
      this.#markQueryFailed(queryRecord, error);
      throw error;
    }

    queryRecord.status = "completed";
    queryRecord.completedAt = new Date().toISOString();
    this.job.updatedAt = queryRecord.completedAt;
    writeJsonAtomic(this.jobRecordPath, this.job);
    return { response, resolution };
  }

  async complete(summary = {}) {
    await this.initialize();
    await this.queryQueue;
    assertJobCanComplete(this.job, this.queryFailure);
    this.job.status = "completed";
    this.job.summary = summary;
    this.job.completedAt = new Date().toISOString();
    this.job.updatedAt = this.job.completedAt;
    writeJsonAtomic(this.jobRecordPath, this.job);
  }

  fail(error) {
    if (!this.job) return;
    this.job.status = "failed";
    this.job.error = String(error?.message || error).slice(0, 12_000);
    this.job.failedAt = new Date().toISOString();
    this.job.updatedAt = this.job.failedAt;
    writeJsonAtomic(this.jobRecordPath, this.job);
  }

  dispose() {
    if (!this.tempRoot) return;
    const base = path.resolve(os.tmpdir()) + path.sep;
    const target = path.resolve(this.tempRoot);
    if (!target.startsWith(base) || !path.basename(target).startsWith("typos-subscription-")) {
      throw new Error(`Refusing to remove unexpected temporary path: ${target}`);
    }
    fs.rmSync(target, { recursive: true, force: true });
    this.tempRoot = null;
  }
}

export function createSubscriptionHarness(options) {
  return new SubscriptionHarness(options);
}

export function runSubscriptionHarnessSelfTest() {
  const info = parseLatestModelInfo([
    "latestModelInfo:",
    "  model: gpt-test-frontier",
    "  migrationGuide: /migration",
    "  promptingGuide: /prompting",
  ].join("\n"));
  if (info?.model !== "gpt-test-frontier" || info?.migrationGuide !== "/migration") {
    throw new Error("Latest-model metadata parser self-test failed.");
  }
  const cleaned = codexEnvironment({
    OPENAI_API_KEY: "secret",
    OPENAI_BASE_URL: "https://example.invalid",
    CODEX_ACCESS_TOKEN: "secret",
    AZURE_OPENAI_ENDPOINT: "https://example.invalid",
    SAFE_VALUE: "retained",
  });
  if (cleaned.OPENAI_API_KEY || cleaned.OPENAI_BASE_URL || cleaned.CODEX_ACCESS_TOKEN || cleaned.AZURE_OPENAI_ENDPOINT || cleaned.SAFE_VALUE !== "retained") {
    throw new Error("Subscription child-environment self-test failed.");
  }
  let collisionRejected = false;
  try {
    assertDistinctPaths([
      { label: "input", path: "C:/temp/same.xlsx" },
      { label: "output", path: "c:/TEMP/same.xlsx" },
    ]);
  } catch (error) {
    collisionRejected = /must not overwrite input/i.test(String(error?.message || error));
  }
  if (!collisionRejected) throw new Error("Path-collision self-test failed.");
  if (assertOfficialModelResponseUrl(OFFICIAL_LATEST_MODEL_URL) !== OFFICIAL_LATEST_MODEL_URL) {
    throw new Error("Official-model URL validation self-test failed.");
  }
  let offDomainRedirectRejected = false;
  try {
    assertOfficialModelResponseUrl("https://example.invalid/latest-model.md");
  } catch (error) {
    offDomainRedirectRejected = /redirected outside/i.test(String(error?.message || error));
  }
  if (!offDomainRedirectRejected) throw new Error("Off-domain model redirect self-test failed.");
  if (parsePositiveInteger("6", "ROW_CONCURRENCY", { maximum: 64 }) !== 6) {
    throw new Error("Positive-integer parser self-test failed.");
  }
  let invalidIntegerRejected = false;
  try {
    parsePositiveInteger("0", "ROW_CONCURRENCY", { maximum: 64 });
  } catch (error) {
    invalidIntegerRejected = /must be an integer/i.test(String(error?.message || error));
  }
  if (!invalidIntegerRejected) throw new Error("Invalid-integer rejection self-test failed.");
  const safeSheet = safeWorksheetName("A very long / invalid worksheet name that exceeds Excel limits");
  if (safeSheet.length > 31 || /[\\/:*?\[\]]/.test(safeSheet)) {
    throw new Error("Worksheet-name sanitizer self-test failed.");
  }
  if (!resultWorksheetName("1234567890123456789012345678901", "_diffs").endsWith("_diffs")) {
    throw new Error("Result worksheet suffix self-test failed.");
  }
  let failedQueryCompletionRejected = false;
  try {
    assertJobCanComplete(
      { queries: [{ queryId: "row_000001", status: "failed" }] },
      new Error("simulated query failure"),
    );
  } catch (error) {
    failedQueryCompletionRejected = /cannot complete a job after a query failure/i.test(
      String(error?.message || error),
    );
  }
  if (!failedQueryCompletionRejected) throw new Error("Failed-query completion self-test failed.");
  return {
    status: "ok",
    officialModelPolicy: MODEL_POLICY,
    reasoningEffort: REASONING_EFFORT,
    apiCredentialsStripped: true,
    pathCollisionRejected: true,
    offDomainModelRedirectRejected: true,
    invalidIntegerRejected: true,
    worksheetNameSanitized: true,
    worksheetSuffixPreserved: true,
    failedQueryCompletionRejected: true,
    boundedCodexExecution: true,
  };
}
