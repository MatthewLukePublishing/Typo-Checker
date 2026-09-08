"use strict";

const assert = require("node:assert/strict");
const path = require("node:path");
const test = require("node:test");
const { pathToFileURL } = require("node:url");

const ROOT = path.resolve(__dirname);
const resolverModule = import(pathToFileURL(path.join(ROOT, "Resolve-LatestSubscriptionModel.mjs")).href);

function guidance(model = "gpt-frontier") {
  return `---\nlatestModelInfo:\n  model: ${model}\n  migrationGuide: /guide\n  promptingGuide: /prompt\n---\n# Official guidance\n`;
}

function model(slug = "gpt-frontier", effort = "xhigh") {
  return { slug, visibility: "list", supported_reasoning_levels: [{ effort }] };
}

function fixture() {
  const state = {
    guidance: guidance(),
    catalog: { models: [model()] },
    calls: [],
    cliCalls: [],
    auth: {
      auth_mode: "chatgpt",
      tokens: { access_token: "fixture-credential", account_id: "fixture-account" },
    },
    login: "Logged in using ChatGPT",
    catalogStatus: 200,
  };
  const deps = {
    runCli: (_runtime, args) => {
      state.cliCalls.push(args);
      return args[0] === "login" ? state.login : "codex-cli 0.152.1";
    },
    readAuth: () => state.auth,
    fetchImpl: async (url, options) => {
      state.calls.push({ url, options });
      if (new URL(url).hostname === "developers.openai.com") {
        if (state.guidanceError) throw new Error("fixture transport detail");
        return new Response(state.guidance, { status: state.guidanceStatus || 200 });
      }
      if (state.catalogError) throw new Error("fixture-credential must never appear in diagnostics");
      return new Response(JSON.stringify(state.catalog), { status: state.catalogStatus });
    },
  };
  return { state, deps };
}

test("every resolution queries both live authorities without cache and records nonsecret evidence", async () => {
  const resolver = await resolverModule;
  const { state, deps } = fixture();
  for (let query = 0; query < 2; query++) {
    const result = await resolver.resolveLatestSubscriptionModel({ expectedModel: "gpt-frontier" }, deps);
    assert.equal(result.model, "gpt-frontier");
    assert.equal(result.reasoningEffort, "xhigh");
    assert.equal(result.discovery, "live-no-cache");
    assert.match(result.officialMetadataSha256, /^[0-9a-f]{64}$/);
    assert.match(result.catalogSha256, /^[0-9a-f]{64}$/);
    assert.doesNotMatch(JSON.stringify(result), /fixture-credential|fixture-account/);
  }
  assert.equal(state.calls.length, 4, "the second check must not reuse the first resolution");
  for (const { url, options } of state.calls) {
    assert.equal(options.cache, "no-store");
    assert.equal(options.redirect, "error");
    assert.match(options.headers["Cache-Control"], /no-cache/);
    assert.ok(options.signal);
    if (url === resolver.OFFICIAL_MODEL_SOURCE) {
      assert.equal(options.headers.Authorization, undefined);
    } else {
      assert.equal(new URL(url).origin, "https://chatgpt.com");
      assert.equal(new URL(url).searchParams.get("client_version"), "0.152.1");
      assert.equal(options.headers.Authorization, "Bearer fixture-credential");
    }
  }
  assert.deepEqual(state.cliCalls, [["login", "status"], ["--version"], ["login", "status"], ["--version"]]);
});

test("missing, hidden, duplicate, or insufficient-effort frontier never selects an older model", async () => {
  const resolver = await resolverModule;
  for (const current of [
    [],
    [{ ...model(), visibility: "hidden" }],
    [model(), model()],
    [{ ...model(), upgrade: { model: "gpt-next" } }],
    [model("gpt-frontier", "high")],
    [{ ...model(), minimal_client_version: [99, 0, 0] }],
  ]) {
    const { state, deps } = fixture();
    state.catalog.models = [{ ...model("gpt-older"), priority: 0 }, ...current];
    await assert.rejects(
      resolver.resolveLatestSubscriptionModel({}, deps),
      { code: resolver.MODEL_POLICY_ERROR },
    );
  }
});

test("changed frontier or weaker reasoning prevents the next query", async () => {
  const resolver = await resolverModule;
  const { state, deps } = fixture();
  await resolver.resolveLatestSubscriptionModel({}, deps);
  state.guidance = guidance("gpt-next");
  state.catalog.models = [model("gpt-next")];
  await assert.rejects(
    resolver.resolveLatestSubscriptionModel({ expectedModel: "gpt-frontier" }, deps),
    /Start a new job/,
  );
  assert.equal(state.calls.length, 3, "a stale job must stop before querying the catalog");
  await assert.rejects(
    resolver.resolveLatestSubscriptionModel({ reasoningEffort: "high" }, deps),
    /requires xhigh/,
  );
  assert.equal(state.calls.length, 3);
});

test("minimum client versions accept strict dotted strings and integer tuples", async () => {
  const resolver = await resolverModule;
  for (const minimum of ["0.152.0", "0.152.1", [0, 152, 1], "0.151.99"]) {
    const { state, deps } = fixture();
    state.catalog.models[0].minimal_client_version = minimum;
    assert.equal((await resolver.resolveLatestSubscriptionModel({}, deps)).model, "gpt-frontier");
  }
  for (const minimum of ["0.153.0", "0.152.2", [0, 153, 0]]) {
    const { state, deps } = fixture();
    state.catalog.models[0].minimal_client_version = minimum;
    await assert.rejects(
      resolver.resolveLatestSubscriptionModel({}, deps),
      /requires a newer Codex client/,
    );
  }
  for (const minimum of [
    null,
    "",
    "0.152",
    "0.152.1-beta",
    "0.152.1junk",
    " 0.152.1",
    "9007199254740992.0.0",
    [],
    [0, 152],
    [0, "152", 1],
    [0, -1, 1],
  ]) {
    const { state, deps } = fixture();
    state.catalog.models[0].minimal_client_version = minimum;
    await assert.rejects(
      resolver.resolveLatestSubscriptionModel({}, deps),
      /invalid minimum client version/,
    );
  }
});

test("live discovery failures cannot reuse a previous resolution or leak credentials", async () => {
  const resolver = await resolverModule;
  for (const failure of ["guidanceError", "catalogError", "catalogStatus", "guidanceStatus"]) {
    const { state, deps } = fixture();
    await resolver.resolveLatestSubscriptionModel({}, deps);
    state[failure] = failure.endsWith("Status") ? 503 : true;
    await assert.rejects(resolver.resolveLatestSubscriptionModel({}, deps), (error) => {
      assert.equal(error.code, resolver.MODEL_POLICY_ERROR);
      assert.doesNotMatch(error.message, /fixture-credential|fixture transport detail/);
      return true;
    });
  }
});

test("official metadata must identify exactly one frontier model", async () => {
  const resolver = await resolverModule;
  assert.equal(resolver.parseOfficialFrontier(guidance().replace(/\n/g, "\r\n")), "gpt-frontier");
  for (const malformed of [
    "# Latest model: gpt-frontier",
    "---\nother: gpt-frontier\n---\n",
    guidance().replace("  model: gpt-frontier", "  model: gpt-frontier\n  model: gpt-other"),
    guidance().replace("  migrationGuide: /guide", "latestModelInfo:\n  model: gpt-other"),
    guidance().replace("gpt-frontier", "https://example.test/model"),
  ]) {
    assert.throws(() => resolver.parseOfficialFrontier(malformed), /unambiguous frontier/);
  }
});

test("API-key, missing, and alternate authentication stop before catalog discovery", async () => {
  const resolver = await resolverModule;
  for (const auth of [
    null,
    { auth_mode: "apikey" },
    { auth_mode: "chatgpt", tokens: {} },
    {
      auth_mode: "chatgpt",
      OPENAI_API_KEY: "fixture-key",
      tokens: { access_token: "fixture-credential", account_id: "fixture-account" },
    },
  ]) {
    const { state, deps } = fixture();
    state.auth = auth;
    await assert.rejects(resolver.resolveLatestSubscriptionModel({}, deps), /ChatGPT session is required/);
    assert.equal(state.calls.length, 1);
  }
  const { state, deps } = fixture();
  state.login = "Logged in using an API key";
  await assert.rejects(resolver.resolveLatestSubscriptionModel({}, deps), /signed in with ChatGPT/);
  assert.equal(state.calls.length, 1);
});
