const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const CODE = fs.readFileSync(path.join(__dirname, "..", "Code.gs"), "utf8");

function latestBody(overrides = {}) {
  return JSON.stringify({
    status: "success",
    data: {
      code: "WTI_USD",
      price: 81.78,
      currency: "USD",
      unit: "barrel",
      source: "market_reporting",
      created_at: "2026-07-19T14:10:50.373Z",
      ...overrides,
    },
  });
}

function historyBody(records) {
  return JSON.stringify({ status: "success", data: { prices: records } });
}

function createHarness() {
  const properties = new Map();
  const cache = new Map();
  const responses = [];
  const requests = [];

  const userProperties = {
    getProperty: (key) => properties.get(key) || null,
    setProperty: (key, value) => properties.set(key, value),
    deleteProperty: (key) => properties.delete(key),
  };
  const userCache = {
    get: (key) => cache.get(key) || null,
    put: (key, value) => cache.set(key, value),
    remove: (key) => cache.delete(key),
  };
  const context = {
    CacheService: { getUserCache: () => userCache },
    Date,
    Error,
    HtmlService: {},
    JSON,
    Math,
    Number,
    PropertiesService: { getUserProperties: () => userProperties },
    SpreadsheetApp: {},
    String,
    UrlFetchApp: {
      fetch(url, options) {
        requests.push({ url, options });
        const next = responses.shift();
        if (next instanceof Error) throw next;
        if (!next) throw new Error("No fixture response queued");
        return {
          getResponseCode: () => next.status,
          getContentText: () => next.body,
        };
      },
    },
    console,
    decodeURIComponent,
    encodeURIComponent,
    isFinite,
    parseFloat,
    parseInt,
    setTimeout,
  };
  vm.createContext(context);
  new vm.Script(CODE, { filename: "Code.gs" }).runInContext(context);

  return {
    cache,
    context,
    properties,
    queue(status, body) {
      responses.push({ status, body });
    },
    queueError(error) {
      responses.push(error);
    },
    requests,
  };
}

function configure(harness) {
  harness.context.saveApiKey("test-key-not-a-secret");
}

test("credential lifecycle never returns the stored key", () => {
  const harness = createHarness();
  assert.throws(() => harness.context.saveApiKey("  "), /API key is required/);
  assert.deepEqual(
    JSON.parse(JSON.stringify(harness.context.getApiKeyStatus())),
    { configured: false },
  );

  const saved = harness.context.saveApiKey("test-key-not-a-secret");
  assert.equal(saved.success, true);
  assert.equal(JSON.stringify(saved).includes("test-key-not-a-secret"), false);
  assert.deepEqual(
    JSON.parse(JSON.stringify(harness.context.getApiKeyStatus())),
    { configured: true },
  );

  harness.context.deleteApiKey();
  assert.equal(harness.context.getApiKeyStatus().configured, false);
});

test("OILPRICE rejects a missing key with a recovery action", () => {
  const harness = createHarness();
  assert.throws(
    () => harness.context.OILPRICE("WTI_USD"),
    /configure an API key/i,
  );
});

for (const [status, pattern] of [
  [401, /invalid or revoked API key/i],
  [403, /pricing/i],
  [429, /rate or quota limit/i],
]) {
  test(`OILPRICE maps HTTP ${status} to actionable recovery`, () => {
    const harness = createHarness();
    configure(harness);
    harness.queue(status, JSON.stringify({ error: "fixture" }));
    assert.throws(() => harness.context.OILPRICE("WTI_USD"), pattern);
  });
}

test("OILPRICE maps Apps Script fetch failures to timeout recovery", () => {
  const harness = createHarness();
  configure(harness);
  harness.queueError(new Error("Socket timeout"));
  assert.throws(() => harness.context.OILPRICE("WTI_USD"), /timed out/i);
});

test("OILPRICE rejects malformed JSON and an empty successful response", () => {
  const malformed = createHarness();
  configure(malformed);
  malformed.queue(200, "not-json");
  assert.throws(() => malformed.context.OILPRICE("WTI_USD"), /malformed JSON/i);

  const empty = createHarness();
  configure(empty);
  empty.queue(200, JSON.stringify({ status: "success", data: { prices: [] } }));
  assert.throws(() => empty.context.OILPRICE("WTI_USD"), /no usable price/i);
});

test("OILPRICE rejects successful records with missing source fields", () => {
  for (const [field, pattern] of [
    ["price", /finite price/i],
    ["currency", /currency/i],
    ["unit", /unit/i],
    ["created_at", /timestamp/i],
  ]) {
    const harness = createHarness();
    configure(harness);
    const record = JSON.parse(latestBody()).data;
    delete record[field];
    harness.queue(200, JSON.stringify({ status: "success", data: record }));
    assert.throws(() => harness.context.OILPRICE("WTI_USD"), pattern);
  }
});

test("OILPRICE supports the production flat record and caches its source data", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(200, latestBody());
  assert.equal(harness.context.OILPRICE("wti_usd"), 81.78);
  assert.equal(harness.context.OILPRICE("WTI_USD"), 81.78);
  assert.equal(harness.requests.length, 1);
  assert.match(harness.requests[0].url, /by_code=WTI_USD$/);
  assert.equal(
    harness.requests[0].options.headers.Authorization,
    "Token test-key-not-a-secret",
  );
});

test("a stale cache envelope is discarded before formula refresh", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(200, latestBody({ price: 80 }));
  assert.equal(harness.context.OILPRICE("WTI_USD"), 80);

  const cacheKey = [...harness.cache.keys()][0];
  const envelope = JSON.parse(harness.cache.get(cacheKey));
  envelope.cachedAt = Date.now() - 301000;
  harness.cache.set(cacheKey, JSON.stringify(envelope));
  harness.queue(200, latestBody({ price: 82 }));

  assert.equal(harness.context.OILPRICE("WTI_USD"), 82);
  assert.equal(harness.requests.length, 2);
});

test("historical formula retains API timestamps and rejects missing timestamps", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(
    200,
    historyBody([
      {
        code: "WTI_USD",
        price: 80,
        currency: "USD",
        unit: "barrel",
        source: "market_reporting",
        created_at: "2026-07-18T14:10:50.373Z",
      },
    ]),
  );
  assert.deepEqual(JSON.parse(JSON.stringify(harness.context.OILPRICE_HISTORY("WTI_USD", 1))), [
    ["2026-07-18T14:10:50.373Z", 80],
  ]);

  const invalid = createHarness();
  configure(invalid);
  invalid.queue(
    200,
    historyBody([
      { code: "WTI_USD", price: 80, currency: "USD", unit: "barrel" },
    ]),
  );
  assert.throws(
    () => invalid.context.OILPRICE_HISTORY("WTI_USD", 1),
    /timestamp/i,
  );
});

test("batch refresh rejects the documented Apps Script request limit", () => {
  const harness = createHarness();
  configure(harness);
  const codes = Array.from({ length: 26 }, (_, index) => `CODE_${index}`);
  assert.throws(
    () => harness.context.fetchLatestPrices(codes),
    /at most 25 commodity codes/i,
  );
  assert.equal(harness.requests.length, 0);
});

test("batch refresh writes only validated API metadata", () => {
  const harness = createHarness();
  configure(harness);
  let written = null;
  harness.context.writeToDataSheet = (records) => {
    written = records;
  };
  harness.queue(
    200,
    historyBody([
      {
        code: "WTI_USD",
        price: 81.78,
        currency: "USD",
        unit: "barrel",
        source: "market_reporting",
        created_at: "2026-07-19T14:10:50.373Z",
      },
    ]),
  );

  const result = harness.context.fetchLatestPrices(["WTI_USD"]);
  assert.equal(result.count, 1);
  assert.deepEqual(JSON.parse(JSON.stringify(written)), [
    {
      code: "WTI_USD",
      price: 81.78,
      currency: "USD",
      unit: "barrel",
      source: "market_reporting",
      timestamp: "2026-07-19T14:10:50.373Z",
    },
  ]);
});

test("reference conversion uses the validated latest record", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(200, latestBody({ price: 58 }));
  assert.equal(harness.context.OILPRICE_CONVERT("WTI_USD"), 10);
  assert.equal(harness.requests.length, 1);
});

test("exchange-rate conversion has no fabricated fallback", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(
    200,
    historyBody([
      {
        code: "GBP_USD",
        price: 1.3,
        currency: "USD",
        unit: "GBP",
        created_at: "2026-07-19T14:10:50.373Z",
      },
    ]),
  );
  assert.throws(() => harness.context.fetchExchangeRates(), /both GBP_USD and EUR_USD/i);
});

test("testConnection validates source data instead of accepting any HTTP 200", () => {
  const invalid = createHarness();
  configure(invalid);
  invalid.queue(200, JSON.stringify({ status: "success", data: {} }));
  assert.equal(invalid.context.testConnection().success, false);

  const valid = createHarness();
  configure(valid);
  valid.queue(200, latestBody());
  assert.equal(valid.context.testConnection().success, true);
});

test("user info does not invent a tier or request limit", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(200, JSON.stringify({ data: {} }));
  const info = harness.context.getUserInfo();
  assert.equal(info.tier, "unknown");
  assert.equal(info.limit, null);
  assert.equal(info.used, null);
});
