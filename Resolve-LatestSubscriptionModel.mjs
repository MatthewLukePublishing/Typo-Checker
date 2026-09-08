import crypto from "node:crypto";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { pathToFileURL } from "node:url";

export const REQUIRED_REASONING_EFFORT = "xhigh";
export const MODEL_POLICY = "official_latest_frontier";
export const MODEL_POLICY_ERROR = "LATEST_MODEL_POLICY_FAILURE";
export const OFFICIAL_MODEL_SOURCE = "https://developers.openai.com/api/docs/guides/latest-model.md";
// ChatGPT subscription discovery, as used by the official Codex client.
// No API endpoint override, bundled catalog, or local model cache is accepted.
export const MODEL_CATALOG_SOURCE = "https://chatgpt.com/backend-api/codex/models";

function policyError(message) {
  const error = new Error(message);
  error.code = MODEL_POLICY_ERROR;
  return error;
}

function codexEnvironment() {
  const childEnv = { ...process.env };
  for (const name of Object.keys(childEnv)) {
    if (
      /^(?:OPENAI_|AZURE_OPENAI_)/i.test(name) ||
      /^(?:CODEX_API_KEY|CODEX_ACCESS_TOKEN)$/i.test(name)
    ) {
      delete childEnv[name];
    }
  }
  return childEnv;
}

function defaultCodexCliPath() {
  const configured = String(process.env.CODEX_CLI_JS || "").trim();
  if (configured) return path.resolve(configured);
  const appData = String(process.env.APPDATA || "").trim();
  return appData
    ? path.join(appData, "npm", "node_modules", "@openai", "codex", "bin", "codex.js")
    : "";
}

function normalizeRuntime(options) {
  if (options.runtime?.command) {
    return {
      command: path.resolve(String(options.runtime.command)),
      prefixArgs: Array.isArray(options.runtime.prefixArgs)
        ? options.runtime.prefixArgs.map((value) => String(value))
        : [],
      cwd: options.runtime.prefixArgs?.length
        ? path.dirname(path.resolve(String(options.runtime.prefixArgs[0])))
        : path.dirname(path.resolve(String(options.runtime.command))),
      timeoutMs: Number(options.timeoutMs || 30_000),
    };
  }
  const legacyRuntime = options.runtime || options;
  const nodePath = path.resolve(String(legacyRuntime.nodePath || process.env.CODEX_NODE_EXE || process.execPath));
  const cliPath = String(legacyRuntime.cliPath || defaultCodexCliPath());
  return {
    command: nodePath,
    prefixArgs: cliPath ? [path.resolve(cliPath)] : [],
    cwd: cliPath ? path.dirname(path.resolve(cliPath)) : path.dirname(nodePath),
    timeoutMs: Number(options.timeoutMs || 30_000),
  };
}

export function parseOfficialFrontier(markdown) {
  const normalized = String(markdown).replace(/^\uFEFF/, "").replace(/\r\n/g, "\n");
  const frontMatter = normalized.match(/^---\n([\s\S]*?)\n---(?:\n|$)/);
  const blocks = frontMatter?.[1].match(/^latestModelInfo: *\n(?:[ \t]+[^\n]*(?:\n|$))+/gm) || [];
  const models = blocks.length === 1
    ? [...blocks[0].matchAll(/^ {2}model: *(gpt-[a-z0-9.-]+) *$/gm)].map((match) => match[1])
    : [];
  if (models.length !== 1) {
    throw policyError("Live OpenAI guidance does not identify one unambiguous frontier model. No fallback is permitted.");
  }
  return models[0];
}

export function supportsEffort(model, effort) {
  return Array.isArray(model?.supported_reasoning_levels) &&
    model.supported_reasoning_levels.some((entry) => entry?.effort === effort);
}

export function selectLatestCatalogModel(catalog, officialModel) {
  if (!/^gpt-[a-z0-9.-]+$/.test(String(officialModel || ""))) {
    throw policyError("A freshly queried official frontier model is required; catalog priority cannot select it.");
  }
  const matches = (Array.isArray(catalog?.models) ? catalog.models : [])
    .filter((model) => model?.slug === officialModel);
  if (matches.length !== 1 || matches[0].visibility !== "list" || matches[0].upgrade) {
    throw policyError(`The live Codex service does not confirm access to the official frontier '${officialModel}'. No older model is permitted.`);
  }
  return matches[0];
}

export function assertModelSupportsEffort(model, requiredEffort = REQUIRED_REASONING_EFFORT) {
  if (!supportsEffort(model, requiredEffort)) {
    throw policyError(`The official frontier '${model?.slug || "(blank)"}' does not support '${requiredEffort}'. Refusing to fall back to an older model.`);
  }
  return model;
}

function runCli(runtime, args) {
  if (!runtime.command || !fs.existsSync(runtime.command)) {
    throw policyError("The official Codex client is required for ChatGPT sign-in and model queries.");
  }
  if (runtime.prefixArgs.length && !fs.existsSync(runtime.prefixArgs[0])) {
    throw policyError("The official Codex client is required for ChatGPT sign-in and model queries.");
  }
  const run = spawnSync(runtime.command, [...runtime.prefixArgs, ...args], {
    cwd: runtime.cwd,
    env: codexEnvironment(),
    encoding: "utf8",
    maxBuffer: 1024 * 1024,
    timeout: runtime.timeoutMs,
    windowsHide: true,
  });
  if (run.error || run.status !== 0) {
    // CLI diagnostics can contain authentication details. Never forward them.
    throw policyError("Codex authentication/client inspection failed. Sign in with ChatGPT; no fallback is permitted.");
  }
  return String(run.stdout || "") + "\n" + String(run.stderr || "");
}

function readManagedChatGptAuth() {
  const codexRoot = process.env.CODEX_HOME || path.join(os.homedir(), ".codex");
  try {
    return JSON.parse(fs.readFileSync(path.join(codexRoot, "auth.json"), "utf8"));
  } catch {
    throw policyError("Live model discovery requires the Codex-managed ChatGPT session in auth.json. A keyring-only session cannot be read by this resolver; no API-key fallback is permitted.");
  }
}

async function fetchLiveText(url, { fetchImpl, timeoutMs, headers = {}, maxBytes }) {
  try {
    const response = await fetchImpl(url, {
      cache: "no-store",
      redirect: "error",
      signal: AbortSignal.timeout(timeoutMs),
      headers: { "Cache-Control": "no-cache, no-store", Pragma: "no-cache", ...headers },
    });
    if (response.status !== 200) {
      throw policyError(`Live model discovery returned HTTP ${response.status} from ${new URL(url).hostname}. No cached fallback is permitted.`);
    }
    const declaredBytes = Number(response.headers.get("content-length"));
    if (declaredBytes > maxBytes) throw policyError("Live model metadata exceeds the permitted response size.");
    let bytes = 0;
    const chunks = [];
    for await (const chunk of response.body) {
      bytes += chunk.length;
      if (bytes > maxBytes) throw policyError("Live model metadata exceeds the permitted response size.");
      chunks.push(Buffer.from(chunk));
    }
    return Buffer.concat(chunks).toString("utf8");
  } catch (error) {
    if (error.code === MODEL_POLICY_ERROR) throw error;
    // Never log request headers, response bodies, or transport diagnostics.
    throw policyError(`Fresh model discovery failed for ${new URL(url).hostname}. No cached fallback is permitted.`);
  }
}

export async function resolveLatestSubscriptionModel(options = {}, dependencies = {}) {
  const timeoutMs = Number(options.timeoutMs || 30_000);
  if (!Number.isSafeInteger(timeoutMs) || timeoutMs < 1000 || timeoutMs > 60_000) {
    throw policyError("Model-discovery timeout must be an integer from 1000 to 60000 milliseconds.");
  }
  if (options.reasoningEffort !== undefined && options.reasoningEffort !== REQUIRED_REASONING_EFFORT) {
    throw policyError("Every primary query, review, retry, and recheck requires xhigh reasoning.");
  }
  const runtime = normalizeRuntime({ ...options, timeoutMs });
  const fetchImpl = dependencies.fetchImpl || globalThis.fetch;
  // Query the official authority on every invocation; keep no memoized result.
  const guidance = await fetchLiveText(OFFICIAL_MODEL_SOURCE, {
    fetchImpl,
    timeoutMs,
    headers: { Accept: "text/markdown" },
    maxBytes: 1024 * 1024,
  });
  const officialModel = parseOfficialFrontier(guidance);
  if (options.expectedModel !== undefined && options.expectedModel !== officialModel) {
    throw policyError(`This job recorded '${options.expectedModel}', but the live official frontier is '${officialModel}'. Start a new job; no stale-model query is permitted.`);
  }

  const inspectCli = dependencies.runCli || runCli;
  if (!/Logged in using ChatGPT/i.test(inspectCli(runtime, ["login", "status"]))) {
    throw policyError("Codex must be signed in with ChatGPT. API-key authentication is not permitted.");
  }
  const clientVersion = inspectCli(runtime, ["--version"])
    .match(/\bcodex-cli (\d+\.\d+\.\d+)(?:[\s.-]|$)/)?.[1];
  if (!clientVersion) throw policyError("Could not verify the Codex client version for live subscription discovery.");
  const auth = (dependencies.readAuth || readManagedChatGptAuth)();
  if (
    auth?.auth_mode !== "chatgpt" ||
    auth.OPENAI_API_KEY ||
    typeof auth.tokens?.access_token !== "string" ||
    !auth.tokens.access_token ||
    typeof auth.tokens?.account_id !== "string" ||
    !auth.tokens.account_id
  ) {
    throw policyError("A Codex-managed ChatGPT session is required. API keys and alternate authentication are not permitted.");
  }
  const catalogUrl = new URL(MODEL_CATALOG_SOURCE);
  catalogUrl.searchParams.set("client_version", clientVersion);
  const catalogText = await fetchLiveText(catalogUrl.href, {
    fetchImpl,
    timeoutMs,
    maxBytes: 8 * 1024 * 1024,
    headers: {
      Accept: "application/json",
      Authorization: `Bearer ${auth.tokens.access_token}`,
      "ChatGPT-Account-Id": auth.tokens.account_id,
    },
  });
  let catalog;
  try {
    catalog = JSON.parse(catalogText);
  } catch {
    throw policyError("The live Codex service returned invalid model metadata. No cached fallback is permitted.");
  }
  const selected = selectLatestCatalogModel(catalog, officialModel);
  assertModelSupportsEffort(selected);
  if (selected.minimal_client_version !== undefined) {
    // The live service can serialize this as a dotted version, while older
    // catalog schemas use a three-integer tuple. Validate both strictly.
    const rawMinimum = selected.minimal_client_version;
    const minimum = typeof rawMinimum === "string" && /^\d+\.\d+\.\d+$/.test(rawMinimum)
      ? rawMinimum.split(".").map(Number)
      : rawMinimum;
    if (
      !Array.isArray(minimum) ||
      minimum.length !== 3 ||
      minimum.some((part) => !Number.isSafeInteger(part) || part < 0)
    ) {
      throw policyError("The live frontier has an invalid minimum client version.");
    }
    const version = clientVersion.split(".").map(Number);
    const difference = version
      .map((part, index) => part - minimum[index])
      .find((part) => part !== 0) || 0;
    if (difference < 0) {
      throw policyError("The live frontier requires a newer Codex client. No older model is permitted.");
    }
  }
  return {
    schemaVersion: 3,
    policy: MODEL_POLICY,
    model: officialModel,
    reasoningEffort: REQUIRED_REASONING_EFFORT,
    resolvedAt: new Date().toISOString(),
    sourceUrl: OFFICIAL_MODEL_SOURCE,
    catalogSourceUrl: catalogUrl.href,
    clientVersion,
    authentication: "chatgpt",
    discovery: "live-no-cache",
    officialMetadataSha256: crypto.createHash("sha256").update(guidance).digest("hex"),
    catalogSha256: crypto.createHash("sha256").update(catalogText).digest("hex"),
  };
}

const invokedUrl = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (import.meta.url === invokedUrl) {
  resolveLatestSubscriptionModel()
    .then((result) => process.stdout.write(`${JSON.stringify(result, null, 2)}\n`))
    .catch((error) => {
      process.stderr.write(`LATEST_MODEL_RESOLUTION_FAILED|${error.message}\n`);
      process.exitCode = 1;
    });
}
