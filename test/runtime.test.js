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

function bunkerBody(records) {
  return JSON.stringify({ status: "success", data: { prices: records } });
}

function bunkerRecord(overrides = {}) {
  return {
    port: "SINGAPORE",
    fuel_type: "VLSFO",
    price: 612.5,
    currency: "USD",
    unit: "metric_ton",
    region: "Asia",
    source: "fixture_provider",
    as_of: "2026-07-25T09:30:00.000Z",
    ...overrides,
  };
}

function createHarness(shared = {}) {
  const documentPropertyValues = shared.documentPropertyValues || new Map();
  const userPropertyValues = shared.userPropertyValues || new Map();
  const documentCache = shared.documentCache || new Map();
  let userCache = shared.userCache || new Map();
  const cachePuts = [];
  let lockAvailable = true;
  let lockHeld = false;
  const responses = [];
  const requests = [];
  let activeSpreadsheetId = "sheet-a";

  const propertyStore = (values) => ({
    getProperty: (key) => values.get(key) || null,
    setProperty: (key, value) => values.set(key, value),
    deleteProperty: (key) => values.delete(key),
  });
  const documentProperties = propertyStore(documentPropertyValues);
  const userProperties = propertyStore(userPropertyValues);
  const cacheStore = (scope, values) => ({
    get: (key) => values().get(key) || null,
    put: (key, value, ttlSeconds) => {
      values().set(key, value);
      cachePuts.push({ scope, key, ttlSeconds });
    },
    remove: (key) => values().delete(key),
  });
  const documentCacheStore = cacheStore("document", () => documentCache);
  const userCacheStore = cacheStore("user", () => userCache);
  const context = {
    CacheService: {
      getDocumentCache: () => documentCacheStore,
      getUserCache: () => userCacheStore,
    },
    Date,
    Error,
    HtmlService: {},
    JSON,
    LockService: {
      getDocumentLock: () => ({
        tryLock: () => {
          if (!lockAvailable || lockHeld) return false;
          lockHeld = true;
          return true;
        },
        releaseLock: () => {
          lockHeld = false;
        },
      }),
    },
    Math,
    Number,
    PropertiesService: {
      getDocumentProperties: () => documentProperties,
      getUserProperties: () => userProperties,
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({
        getId: () => activeSpreadsheetId,
      }),
    },
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
          getHeaders: () => next.headers || {},
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
    cache: documentCache,
    cachePuts,
    context,
    documentCache,
    documentPropertyValues,
    queue(status, body, headers = {}) {
      responses.push({ status, body, headers });
    },
    queueError(error) {
      responses.push(error);
    },
    requests,
    setLockAvailable(value) {
      lockAvailable = value;
    },
    setActiveSpreadsheetId(value) {
      activeSpreadsheetId = value;
    },
    switchUserCache() {
      userCache = new Map();
    },
    get userCache() {
      return userCache;
    },
    userPropertyValues,
  };
}

function attachSpreadsheet(harness, existingSheet = true) {
  const calls = [];
  const alerts = [];
  const ranges = [];

  function makeRange(row, column, rowCount, columnCount) {
    const state = { row, column, rowCount, columnCount };
    const range = {
      setBackground(value) {
        state.background = value;
        return range;
      },
      setFontWeight(value) {
        state.fontWeight = value;
        return range;
      },
      setNumberFormat(value) {
        state.numberFormat = value;
        return range;
      },
      setValues(value) {
        state.values = JSON.parse(JSON.stringify(value));
        return range;
      },
    };
    ranges.push(state);
    return range;
  }

  const sheet = {
    activate() {
      calls.push(["activate"]);
    },
    autoResizeColumns(start, count) {
      calls.push(["autoResizeColumns", start, count]);
    },
    clear() {
      calls.push(["clear"]);
    },
    getRange(row, column, rowCount, columnCount) {
      calls.push(["getRange", row, column, rowCount, columnCount]);
      if (rowCount < 1 || columnCount < 1) {
        throw new Error("Spreadsheet ranges must contain at least one cell.");
      }
      return makeRange(row, column, rowCount, columnCount);
    },
  };
  let availableSheet = existingSheet ? sheet : null;
  const spreadsheet = {
    getSheetByName(name) {
      calls.push(["getSheetByName", name]);
      return availableSheet;
    },
    insertSheet(name) {
      calls.push(["insertSheet", name]);
      availableSheet = sheet;
      return sheet;
    },
  };
  harness.context.SpreadsheetApp = {
    getActiveSpreadsheet: () => spreadsheet,
    getUi: () => ({
      alert(message) {
        alerts.push(message);
      },
    }),
  };

  return { alerts, calls, ranges, sheet, spreadsheet };
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
  assert.equal(
    harness.documentPropertyValues.get("OILPRICEAPI_KEY"),
    "test-key-not-a-secret",
  );
  assert.equal(harness.userPropertyValues.has("OILPRICEAPI_KEY"), false);
  assert.deepEqual(
    JSON.parse(JSON.stringify(harness.context.getApiKeyStatus())),
    { configured: true },
  );

  harness.documentPropertyValues.set(
    "OILPRICEAPI_LAST_DIAGNOSTIC",
    JSON.stringify({ code: "OK" }),
  );
  harness.context.deleteApiKey();
  assert.equal(harness.context.getApiKeyStatus().configured, false);
  assert.equal(harness.documentPropertyValues.has("OILPRICEAPI_KEY"), false);
  assert.equal(
    harness.documentPropertyValues.has("OILPRICEAPI_LAST_DIAGNOSTIC"),
    false,
  );
});

test("credential save fails closed without an active spreadsheet ID", () => {
  const harness = createHarness();
  harness.context.SpreadsheetApp.getActiveSpreadsheet = () => null;

  assert.throws(
    () => harness.context.saveApiKey("test-key-not-a-secret"),
    /open the OilPriceAPI sidebar from a Google Sheet/i,
  );
  assert.equal(harness.documentPropertyValues.size, 0);
  assert.equal(harness.userPropertyValues.size, 0);
});

test("spreadsheet-scoped key survives the custom-function user identity boundary", () => {
  const harness = createHarness();
  configure(harness);
  harness.context.PropertiesService.getUserProperties = () => ({
    getProperty: () => null,
    setProperty: () => undefined,
    deleteProperty: () => undefined,
  });
  harness.queue(200, latestBody());

  assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
  assert.match(
    harness.requests[0].options.headers.Authorization,
    /^Token test-key-not-a-secret$/,
  );
});

test("custom functions use a spreadsheet-scoped owner fallback when document properties are unavailable", () => {
  const harness = createHarness();
  harness.context.saveApiKey("test-key-not-a-secret");
  assert.equal(
    harness.userPropertyValues.get("OILPRICEAPI_KEY:sheet-a"),
    "test-key-not-a-secret",
  );
  harness.context.PropertiesService.getDocumentProperties = () => null;
  harness.queue(200, latestBody());

  assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
  assert.equal(
    harness.requests[0].options.headers.Authorization,
    "Token test-key-not-a-secret",
  );
});

test("spreadsheet-scoped owner fallback never crosses into another spreadsheet", () => {
  const harness = createHarness();
  harness.context.saveApiKey("sheet-a-key");
  harness.setActiveSpreadsheetId("sheet-b");
  harness.context.PropertiesService.getDocumentProperties = () => null;

  assert.match(
    harness.context.OILPRICE_PRICE("WTI_USD"),
    /^#AUTH_REQUIRED:/,
  );
});

test("user-scoped fallback credentials never share cached responses", () => {
  const shared = {
    documentCache: new Map(),
    documentPropertyValues: new Map(),
  };
  const firstUser = createHarness(shared);
  firstUser.context.PropertiesService.getDocumentProperties = () => null;
  firstUser.userPropertyValues.set("OILPRICEAPI_KEY:sheet-a", "first-user-key");
  firstUser.queue(200, latestBody({ price: 81.78 }));
  assert.equal(firstUser.context.OILPRICE_PRICE("WTI_USD"), 81.78);

  const secondUser = createHarness(shared);
  secondUser.context.PropertiesService.getDocumentProperties = () => null;
  secondUser.userPropertyValues.set("OILPRICEAPI_KEY:sheet-a", "second-user-key");
  secondUser.queue(200, latestBody({ price: 92.35 }));
  assert.equal(secondUser.context.OILPRICE_PRICE("WTI_USD"), 92.35);
  assert.equal(secondUser.requests.length, 1);
  assert.equal(
    secondUser.requests[0].options.headers.Authorization,
    "Token second-user-key",
  );
  assert.ok(firstUser.cachePuts.every((entry) => entry.scope === "user"));
  assert.ok(secondUser.cachePuts.every((entry) => entry.scope === "user"));
  assert.equal(
    JSON.stringify([
      ...shared.documentCache.entries(),
      ...firstUser.userCache.entries(),
      ...secondUser.userCache.entries(),
    ]).includes("user-key"),
    false,
  );
});

test("user-scoped fallback credentials never share entitlement blocks", () => {
  const shared = {
    documentCache: new Map(),
    documentPropertyValues: new Map(),
  };
  const firstUser = createHarness(shared);
  firstUser.context.PropertiesService.getDocumentProperties = () => null;
  firstUser.userPropertyValues.set("OILPRICEAPI_KEY:sheet-a", "first-user-key");
  firstUser.queue(403, JSON.stringify({ error: "upgrade required" }));
  assert.match(
    firstUser.context.OILPRICE_HISTORY("WTI_USD", 30)[0][0],
    /^#UPGRADE_REQUIRED$/,
  );

  const secondUser = createHarness(shared);
  secondUser.context.PropertiesService.getDocumentProperties = () => null;
  secondUser.userPropertyValues.set("OILPRICEAPI_KEY:sheet-a", "second-user-key");
  secondUser.queue(
    200,
    historyBody([
      {
        code: "WTI_USD",
        price: 92.35,
        currency: "USD",
        unit: "barrel",
        source: "market_reporting",
        created_at: "2026-08-11T12:00:00.000Z",
      },
    ]),
  );
  assert.deepEqual(
    JSON.parse(
      JSON.stringify(secondUser.context.OILPRICE_HISTORY("WTI_USD", 30)),
    ),
    [["2026-08-11T12:00:00.000Z", 92.35]],
  );
  assert.equal(secondUser.requests.length, 1);
});

test("one user's fallback response cache is isolated between spreadsheets", () => {
  const shared = {
    documentCache: new Map(),
    documentPropertyValues: new Map(),
    userCache: new Map(),
    userPropertyValues: new Map([
      ["OILPRICEAPI_KEY:sheet-a", "sheet-a-key"],
      ["OILPRICEAPI_KEY:sheet-b", "sheet-b-key"],
    ]),
  };
  const firstSheet = createHarness(shared);
  firstSheet.context.PropertiesService.getDocumentProperties = () => null;
  firstSheet.queue(200, latestBody({ price: 81.78 }));
  assert.equal(firstSheet.context.OILPRICE_PRICE("WTI_USD"), 81.78);

  const secondSheet = createHarness(shared);
  secondSheet.setActiveSpreadsheetId("sheet-b");
  secondSheet.context.PropertiesService.getDocumentProperties = () => null;
  secondSheet.queue(200, latestBody({ price: 92.35 }));
  assert.equal(secondSheet.context.OILPRICE_PRICE("WTI_USD"), 92.35);
  assert.equal(secondSheet.requests.length, 1);
  assert.equal(
    secondSheet.requests[0].options.headers.Authorization,
    "Token sheet-b-key",
  );
  assert.ok(
    [...shared.userCache.keys()].every(
      (key) => !key.includes("sheet-a") && !key.includes("sheet-b"),
    ),
  );
});

test("one user's fallback entitlement blocks are isolated between spreadsheets", () => {
  const shared = {
    documentCache: new Map(),
    documentPropertyValues: new Map(),
    userCache: new Map(),
    userPropertyValues: new Map([
      ["OILPRICEAPI_KEY:sheet-a", "sheet-a-key"],
      ["OILPRICEAPI_KEY:sheet-b", "sheet-b-key"],
    ]),
  };
  const firstSheet = createHarness(shared);
  firstSheet.context.PropertiesService.getDocumentProperties = () => null;
  firstSheet.queue(403, JSON.stringify({ error: "upgrade required" }));
  assert.match(
    firstSheet.context.OILPRICE_HISTORY("WTI_USD", 30)[0][0],
    /^#UPGRADE_REQUIRED$/,
  );

  const secondSheet = createHarness(shared);
  secondSheet.setActiveSpreadsheetId("sheet-b");
  secondSheet.context.PropertiesService.getDocumentProperties = () => null;
  secondSheet.queue(
    200,
    historyBody([
      {
        code: "WTI_USD",
        price: 92.35,
        currency: "USD",
        unit: "barrel",
        source: "market_reporting",
        created_at: "2026-08-11T12:00:00.000Z",
      },
    ]),
  );
  assert.deepEqual(
    JSON.parse(
      JSON.stringify(secondSheet.context.OILPRICE_HISTORY("WTI_USD", 30)),
    ),
    [["2026-08-11T12:00:00.000Z", 92.35]],
  );
  assert.equal(secondSheet.requests.length, 1);
});

test("deleting a key removes both document and spreadsheet-scoped owner stores", () => {
  const harness = createHarness();
  harness.context.saveApiKey("test-key-not-a-secret");
  harness.context.deleteApiKey();

  assert.equal(harness.documentPropertyValues.has("OILPRICEAPI_KEY"), false);
  assert.equal(
    harness.userPropertyValues.has("OILPRICEAPI_KEY:sheet-a"),
    false,
  );
});

test("an unscoped legacy user-property key cannot authorize another spreadsheet", () => {
  const harness = createHarness();
  harness.userPropertyValues.set("OILPRICEAPI_KEY", "legacy-key");
  harness.setActiveSpreadsheetId("sheet-b");

  assert.match(
    harness.context.OILPRICE_PRICE("WTI_USD"),
    /^#AUTH_REQUIRED:/,
  );
  assert.equal(harness.requests.length, 0);

  harness.context.saveApiKey("spreadsheet-key");
  assert.equal(harness.userPropertyValues.has("OILPRICEAPI_KEY"), false);
  assert.equal(
    harness.documentPropertyValues.get("OILPRICEAPI_KEY"),
    "spreadsheet-key",
  );
});

test("OILPRICE rejects a missing key with a recovery action", () => {
  const harness = createHarness();
  assert.match(
    harness.context.OILPRICE("WTI_USD"),
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
    assert.match(harness.context.OILPRICE("WTI_USD"), pattern);
  });
}

test("OILPRICE maps Apps Script fetch failures to timeout recovery", () => {
  const harness = createHarness();
  configure(harness);
  harness.queueError(new Error("Socket timeout"));
  assert.match(harness.context.OILPRICE("WTI_USD"), /timed out/i);
});

test("OILPRICE rejects malformed JSON and an empty successful response", () => {
  const malformed = createHarness();
  configure(malformed);
  malformed.queue(200, "not-json");
  assert.match(malformed.context.OILPRICE("WTI_USD"), /malformed JSON/i);

  const empty = createHarness();
  configure(empty);
  empty.queue(200, JSON.stringify({ status: "success", data: { prices: [] } }));
  assert.match(empty.context.OILPRICE("WTI_USD"), /no usable price/i);
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
    assert.match(harness.context.OILPRICE("WTI_USD"), pattern);
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

test("a CacheService outage degrades to a live response", () => {
  const harness = createHarness();
  configure(harness);
  const unavailable = () => {
    throw new Error("Cache service unavailable");
  };
  harness.context.CacheService.getDocumentCache = unavailable;
  harness.context.CacheService.getUserCache = unavailable;
  harness.queue(200, latestBody());

  assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
  assert.equal(harness.requests.length, 1);
});

for (const operation of ["get", "put", "remove"]) {
  test(`a cache ${operation} failure does not replace live API data`, () => {
    const harness = createHarness();
    configure(harness);
    const cache = {
      get: () => null,
      put: () => {},
      remove: () => {},
    };
    cache[operation] = () => {
      throw new Error(`Cache ${operation} unavailable`);
    };
    harness.context.CacheService.getDocumentCache = () => cache;
    harness.context.CacheService.getUserCache = () => cache;
    harness.queue(200, latestBody());

    assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
    assert.equal(harness.requests.length, 1);
  });
}

for (const operation of ["getDocumentLock", "tryLock", "releaseLock"]) {
  test(`a lock ${operation} failure degrades to an unlocked live response`, () => {
    const harness = createHarness();
    configure(harness);
    harness.context.LockService.getDocumentLock = () => ({
      tryLock: () => true,
      releaseLock: () => {},
    });
    if (operation === "getDocumentLock") {
      harness.context.LockService.getDocumentLock = () => {
        throw new Error("Lock service unavailable");
      };
    } else {
      harness.context.LockService.getDocumentLock = () => ({
        tryLock: () => {
          if (operation === "tryLock") throw new Error("Lock service unavailable");
          return true;
        },
        releaseLock: () => {
          if (operation === "releaseLock") throw new Error("Lock service unavailable");
        },
      });
    }
    harness.queue(200, latestBody());

    assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
    assert.equal(harness.requests.length, 1);
  });
}

test("quota failures are cached and repeated formulas do not refetch", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(
    402,
    JSON.stringify({
      error_code: "PAYMENT_REQUIRED",
      message: "You have used all requests for the current limit window",
      upgrade_url: "https://www.oilpriceapi.com/pricing",
    }),
    { "X-RateLimit-Reset": String(Math.floor(Date.now() / 1000) + 3600) },
  );

  const first = harness.context.OILPRICE_PRICE("WTI_USD");
  const second = harness.context.OILPRICE_PRICE("WTI_USD");

  assert.match(first, /^#UPGRADE_REQUIRED:.*pricing/i);
  assert.equal(second, first);
  assert.equal(harness.requests.length, 1);
});

test("a connection check bypasses a cached quota wall and clears it after upgrade", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(402, JSON.stringify({ error_code: "PAYMENT_REQUIRED" }));
  assert.match(
    harness.context.OILPRICE_PRICE("WTI_USD"),
    /^#UPGRADE_REQUIRED:/,
  );

  harness.queue(200, latestBody());
  assert.equal(harness.context.testConnection().success, true);
  harness.queue(200, latestBody());
  assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
  assert.equal(harness.requests.length, 3);
});

test("a successful connection check invalidates every cached entitlement block", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(403, JSON.stringify({ error: "upgrade required" }));
  assert.match(
    harness.context.OILPRICE_HISTORY("WTI_USD", 30)[0][0],
    /^#UPGRADE_REQUIRED$/,
  );

  harness.queue(200, latestBody());
  assert.equal(harness.context.testConnection().success, true);
  harness.queue(
    200,
    historyBody([
      {
        code: "WTI_USD",
        price: 82.45,
        currency: "USD",
        unit: "barrel",
        source: "market_reporting",
        created_at: "2026-08-11T12:05:00.000Z",
      },
    ]),
  );
  assert.deepEqual(
    JSON.parse(JSON.stringify(harness.context.OILPRICE_HISTORY("WTI_USD", 30))),
    [["2026-08-11T12:05:00.000Z", 82.45]],
  );
  assert.equal(harness.requests.length, 3);
});

test("a fallback-user connection check invalidates that user's cached blocks", () => {
  const harness = createHarness();
  harness.context.PropertiesService.getDocumentProperties = () => null;
  harness.userPropertyValues.set("OILPRICEAPI_KEY:sheet-a", "fallback-user-key");
  harness.userPropertyValues.set("OILPRICEAPI_CACHE_GENERATION:sheet-a", "100");
  harness.queue(403, JSON.stringify({ error: "upgrade required" }));
  assert.match(
    harness.context.OILPRICE_HISTORY("WTI_USD", 30)[0][0],
    /^#UPGRADE_REQUIRED$/,
  );

  harness.queue(200, latestBody());
  assert.equal(harness.context.testConnection().success, true);
  assert.notEqual(
    harness.userPropertyValues.get("OILPRICEAPI_CACHE_GENERATION:sheet-a"),
    "100",
  );
  harness.queue(
    200,
    historyBody([
      {
        code: "WTI_USD",
        price: 83.1,
        currency: "USD",
        unit: "barrel",
        source: "market_reporting",
        created_at: "2026-08-11T12:10:00.000Z",
      },
    ]),
  );
  assert.deepEqual(
    JSON.parse(JSON.stringify(harness.context.OILPRICE_HISTORY("WTI_USD", 30))),
    [["2026-08-11T12:10:00.000Z", 83.1]],
  );
  assert.equal(harness.requests.length, 3);
});

test("a fallback-user connection check reports cache invalidation failure", () => {
  const harness = createHarness();
  harness.context.PropertiesService.getDocumentProperties = () => null;
  harness.userPropertyValues.set("OILPRICEAPI_KEY:sheet-a", "fallback-user-key");
  harness.context.PropertiesService.getUserProperties = () => ({
    getProperty: (key) => harness.userPropertyValues.get(key) || null,
    setProperty: () => {
      throw new Error("Properties service unavailable");
    },
    deleteProperty: (key) => harness.userPropertyValues.delete(key),
  });
  harness.queue(200, latestBody());

  const result = harness.context.testConnection();
  assert.equal(result.success, false);
  assert.match(result.message, /cached worksheet state.*test connection/i);
});

test("a document connection check fails closed when its generation write fails", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(403, JSON.stringify({ error: "upgrade required" }));
  assert.match(
    harness.context.OILPRICE_HISTORY("WTI_USD", 30)[0][0],
    /^#UPGRADE_REQUIRED$/,
  );

  harness.context.PropertiesService.getDocumentProperties = () => ({
    getProperty: (key) => harness.documentPropertyValues.get(key) || null,
    setProperty: () => {
      throw new Error("Document properties unavailable");
    },
    deleteProperty: (key) => harness.documentPropertyValues.delete(key),
  });
  harness.queue(200, latestBody());

  const result = harness.context.testConnection();
  assert.equal(result.success, false);
  assert.match(result.message, /cached worksheet state.*test connection/i);
});

test("all worksheet formulas return readable errors instead of raw exceptions", () => {
  const harness = createHarness();
  const scalarCalls = [
    () => harness.context.OILPRICE("WTI_USD"),
    () => harness.context.OILPRICE_CONVERT("WTI_USD"),
    () => harness.context.BUNKER_PRICE("SINGAPORE", "VLSFO"),
    () => harness.context.FUTURES_PRICE("BZ"),
    () => harness.context.RIG_COUNT("oil"),
  ];
  const tableCalls = [
    () => harness.context.OILPRICE_HISTORY("WTI_USD", 30),
    () => harness.context.BUNKER_PORT_PRICES("SINGAPORE"),
    () => harness.context.FUTURES_CURVE("BZ"),
  ];

  for (const call of scalarCalls) {
    assert.match(call(), /^#AUTH_REQUIRED:/);
  }
  for (const call of tableCalls) {
    assert.match(call()[0][0], /^#AUTH_REQUIRED$/);
  }
  assert.equal(harness.requests.length, 0);
});

test("OILPRICE_TABLE batches a range and primes the shared latest cache", () => {
  const harness = createHarness();
  configure(harness);
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
      {
        code: "BRENT_CRUDE_USD",
        price: 85.45,
        currency: "USD",
        unit: "barrel",
        source: "market_reporting",
        created_at: "2026-07-19T14:10:50.373Z",
      },
    ]),
    { "X-RateLimit-Tier": "free" },
  );

  assert.deepEqual(
    JSON.parse(
      JSON.stringify(
        harness.context.OILPRICE_TABLE([
          ["wti_usd"],
          ["BRENT_CRUDE_USD"],
          ["WTI_USD"],
          [""],
        ]),
      ),
    ),
    [
      ["Code", "Price", "Currency", "Unit", "Source", "Source Timestamp"],
      ["WTI_USD", 81.78, "USD", "barrel", "market_reporting", "2026-07-19T14:10:50.373Z"],
      ["BRENT_CRUDE_USD", 85.45, "USD", "barrel", "market_reporting", "2026-07-19T14:10:50.373Z"],
    ],
  );
  assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
  assert.equal(harness.requests.length, 1);
  assert.match(harness.requests[0].url, /by_code=WTI_USD%2CBRENT_CRUDE_USD$/);
});

test("OILPRICE_TABLE preserves valid rows when one batch record is malformed", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(
    200,
    historyBody([
      {
        code: "WTI_USD",
        price: "not-a-number",
        currency: "USD",
        unit: "barrel",
        created_at: "2026-07-19T14:10:50.373Z",
      },
      {
        code: "BRENT_CRUDE_USD",
        price: 85.45,
        currency: "USD",
        unit: "barrel",
        source: "market_reporting",
        created_at: "2026-07-19T14:10:50.373Z",
      },
    ]),
  );

  const table = JSON.parse(
    JSON.stringify(
      harness.context.OILPRICE_TABLE([["WTI_USD"], ["BRENT_CRUDE_USD"]]),
    ),
  );
  assert.equal(table[1][0], "WTI_USD");
  assert.match(table[1][1], /^#INVALID_RESPONSE:/);
  assert.deepEqual(table[2], [
    "BRENT_CRUDE_USD",
    85.45,
    "USD",
    "barrel",
    "market_reporting",
    "2026-07-19T14:10:50.373Z",
  ]);
  assert.equal(harness.context.OILPRICE_PRICE("BRENT_CRUDE_USD"), 85.45);
  assert.equal(harness.requests.length, 1);
});

test("history cache reuses the endpoint bucket for equivalent lookbacks", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(
    200,
    historyBody([
      {
        code: "WTI_USD",
        price: 81.78,
        currency: "USD",
        unit: "barrel",
        created_at: "2026-07-19T14:10:50.373Z",
      },
    ]),
  );

  assert.deepEqual(
    JSON.parse(JSON.stringify(harness.context.OILPRICE_HISTORY("WTI_USD", 8))),
    [["2026-07-19T14:10:50.373Z", 81.78]],
  );
  assert.deepEqual(
    JSON.parse(JSON.stringify(harness.context.OILPRICE_HISTORY("WTI_USD", 15))),
    [["2026-07-19T14:10:50.373Z", 81.78]],
  );
  assert.equal(harness.requests.length, 1);
});

test("request block keys include a bounded path prefix as well as a hash", () => {
  const harness = createHarness();
  const keys = harness.context.requestBlockKeys_(
    "/prices/latest?by_code=WTI_USD",
  );

  assert.match(keys.endpoint, /request_block_endpoint__v1_prices_latest_[a-z0-9]+$/);
  assert.match(
    keys.request,
    /request_block_request__prices_latest_by_code_WTI_USD_[a-z0-9]+$/,
  );
  assert.ok(keys.request.length < 180);
});

test("latest values use the document cache across spreadsheet viewers", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(200, latestBody(), { "X-RateLimit-Tier": "developer" });

  assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
  harness.switchUserCache();
  assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
  assert.equal(harness.requests.length, 1);
});

test("cold-cache contention fails readably without issuing a duplicate request", () => {
  const harness = createHarness();
  configure(harness);
  harness.setLockAvailable(false);

  assert.match(
    harness.context.OILPRICE_PRICE("WTI_USD"),
    /^#RETRY_LATER:/,
  );
  assert.equal(harness.requests.length, 0);
});

for (const [tier, expectedTtl] of [
  ["free", 3600],
  ["developer", 300],
  ["enterprise", 60],
]) {
  test(`latest cache TTL follows the ${tier} response tier`, () => {
    const harness = createHarness();
    configure(harness);
    harness.queue(200, latestBody(), { "X-RateLimit-Tier": tier });

    assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
    const latestPut = harness.cachePuts.find((entry) =>
      entry.key.includes("latest_WTI_USD"),
    );
    assert.equal(latestPut.scope, "document");
    assert.equal(latestPut.ttlSeconds, expectedTtl);
  });
}

test("saving a replacement key invalidates terminal errors immediately", () => {
  const harness = createHarness();
  harness.context.saveApiKey("first-test-key");
  harness.queue(401, JSON.stringify({ error: "invalid key" }));
  assert.match(
    harness.context.OILPRICE_PRICE("WTI_USD"),
    /^#AUTH_INVALID:/,
  );

  harness.context.saveApiKey("replacement-test-key");
  harness.queue(200, latestBody());
  assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
  assert.equal(harness.requests.length, 2);
});

test("custom-function endpoint families read through document-scoped caches", () => {
  const generic = createHarness();
  configure(generic);
  generic.queue(200, JSON.stringify({ data: { ok: true } }));
  assert.deepEqual(
    JSON.parse(JSON.stringify(generic.context.OILPRICE_GET("/v1/status", ""))),
    [["Field", "Value"], ["ok", true]],
  );
  assert.deepEqual(
    JSON.parse(JSON.stringify(generic.context.OILPRICE_GET("/v1/status", ""))),
    [["Field", "Value"], ["ok", true]],
  );
  assert.equal(generic.requests.length, 1);

  const bunker = createHarness();
  configure(bunker);
  bunker.queue(200, bunkerBody([bunkerRecord()]));
  assert.equal(bunker.context.BUNKER_PRICE("SINGAPORE", "VLSFO"), 612.5);
  assert.equal(bunker.context.BUNKER_PRICE("SINGAPORE", "VLSFO"), 612.5);
  assert.equal(bunker.requests.length, 1);

  const futures = createHarness();
  configure(futures);
  futures.queue(
    200,
    JSON.stringify({ data: { contracts: [{ month: "2026-09", price: 84.2, change: 0.4 }] } }),
  );
  assert.equal(futures.context.FUTURES_PRICE("BZ"), 84.2);
  assert.equal(futures.context.FUTURES_PRICE("BZ"), 84.2);
  assert.equal(futures.requests.length, 1);

  const rigs = createHarness();
  configure(rigs);
  rigs.queue(
    200,
    JSON.stringify({ data: { oil: 410, gas: 125, total: 535, date: "2026-08-07" } }),
  );
  assert.equal(rigs.context.RIG_COUNT("oil"), 410);
  assert.equal(rigs.context.RIG_COUNT("total"), 535);
  assert.equal(rigs.requests.length, 1);
});

test("cache keys and payloads never contain an API key", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(200, latestBody());
  assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);

  const cachedText = JSON.stringify([
    ...harness.documentCache.entries(),
    ...harness.userCache.entries(),
  ]);
  assert.equal(cachedText.includes("test-key-not-a-secret"), false);
});

test("Excel-equivalent price, status, unit, and info formulas preserve source context", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(
    200,
    latestBody({
      currency: "GBp",
      unit: "therm",
      source: "exchange",
      data_status: "stale",
      stale: true,
      age_days: 2,
      formatted: "81.78 GBp/therm",
      collected_at: "2026-07-19T14:11:00.000Z",
      metadata: { source_description: "Fixture exchange" },
    }),
    { "x-request-id": "request-fixture-123" },
  );

  assert.equal(harness.context.OILPRICE_PRICE("wti_usd"), 81.78);
  assert.equal(harness.context.OILPRICE_STATUS("WTI_USD"), "stale");
  assert.equal(harness.context.OILPRICE_UNIT("WTI_USD"), "GBp/therm");
  assert.deepEqual(
    JSON.parse(JSON.stringify(harness.context.OILPRICE_INFO("WTI_USD"))),
    [
      ["Field", "Value"],
      ["code", "WTI_USD"],
      ["price", 81.78],
      ["currency", "GBp"],
      ["unit", "therm"],
      ["formatted", "81.78 GBp/therm"],
      ["source", "exchange"],
      ["source_description", "Fixture exchange"],
      ["source_timestamp", "2026-07-19T14:10:50.373Z"],
      ["collected_at", "2026-07-19T14:11:00.000Z"],
      ["data_status", "stale"],
      ["stale", true],
      ["age_days", 2],
    ],
  );
  assert.equal(harness.requests.length, 1);

  const diagnostic = harness.context.getLastDiagnostic();
  assert.equal(diagnostic.code, "OK");
  assert.equal(diagnostic.httpStatus, 200);
  assert.equal(diagnostic.endpoint, "/v1/prices/latest");
  assert.equal(diagnostic.requestId, "request-fixture-123");
  assert.equal(JSON.stringify(diagnostic).includes("test-key-not-a-secret"), false);
  assert.equal(JSON.stringify(diagnostic).includes("by_code"), false);
});

test("OILPRICE_GET and OILPRICE_CODES match the Excel allowlisted table contract", () => {
  const getHarness = createHarness();
  configure(getHarness);
  getHarness.queue(
    200,
    JSON.stringify({
      status: "success",
      data: {
        prices: [
          {
            code: "WTI_USD",
            price: "81.78",
            currency: "USD",
            unit: "barrel",
          },
        ],
      },
    }),
  );
  assert.deepEqual(
    JSON.parse(
      JSON.stringify(
        getHarness.context.OILPRICE_GET(
          "/v1/prices/latest",
          "by_code=WTI_USD",
        ),
      ),
    ),
    [
      ["code", "price", "currency", "unit"],
      ["WTI_USD", 81.78, "USD", "barrel"],
    ],
  );

  const codesHarness = createHarness();
  configure(codesHarness);
  codesHarness.queue(
    200,
    JSON.stringify({
      status: "success",
      data: {
        commodities: [
          { code: "WTI_USD", name: "WTI", category: "crude" },
        ],
      },
    }),
  );
  assert.deepEqual(
    JSON.parse(JSON.stringify(codesHarness.context.OILPRICE_CODES())),
    [
      ["Code", "Name", "Category"],
      ["WTI_USD", "WTI", "crude"],
    ],
  );
});

test("OILPRICE_GET rejects unsupported endpoints and credential-shaped query keys before fetch", () => {
  const harness = createHarness();
  configure(harness);

  assert.deepEqual(
    JSON.parse(
      JSON.stringify(harness.context.OILPRICE_GET("/v1/users/me", "")),
    ),
    [
      [
        "#UNSUPPORTED_ENDPOINT",
        "Use a supported OilPriceAPI GET endpoint.",
      ],
    ],
  );
  assert.deepEqual(
    JSON.parse(
      JSON.stringify(
        harness.context.OILPRICE_GET(
          "/v1/prices/latest",
          "by_code=WTI_USD&api_key=must-not-leak",
        ),
      ),
    ),
    [
      [
        "#UNSUPPORTED_QUERY",
        "Do not pass API keys or credentials in query strings.",
      ],
    ],
  );
  assert.equal(harness.requests.length, 0);
});

test("Excel-equivalent formulas return stable worksheet error codes", () => {
  const missing = createHarness();
  assert.match(
    missing.context.OILPRICE_PRICE("WTI_USD"),
    /^#AUTH_REQUIRED:/,
  );

  const invalid = createHarness();
  configure(invalid);
  invalid.queue(
    422,
    JSON.stringify({
      data: {
        error: "invalid_code",
        message: "Unknown code. Did you mean WTI_USD?",
      },
    }),
  );
  assert.match(
    invalid.context.OILPRICE_PRICE("WIT_USD"),
    /^#INVALID_CODE: Unknown code\. Did you mean WTI_USD\?$/,
  );
});

test("OILPRICE_GET accepts the full reviewed Excel endpoint catalog", () => {
  const paths = [
    "/v1/status",
    "/v1/prices",
    "/v1/prices/latest",
    "/v1/prices/past_day",
    "/v1/prices/past_week",
    "/v1/prices/past_month",
    "/v1/prices/past_year",
    "/v1/prices/historical",
    "/v1/prices/all",
    "/v1/prices/all/health",
    "/v1/diesel-prices",
    "/v1/commodities",
    "/v1/commodities/categories",
    "/v1/commodities/BRENT_CRUDE_USD",
    "/v1/futures/ice-brent",
    "/v1/futures/ice-wti",
    "/v1/futures/ice-gasoil",
    "/v1/futures/natural-gas",
    "/v1/futures/eua-carbon",
    "/v1/futures/ice-brent/historical",
    "/v1/futures/ice-brent/ohlc",
    "/v1/futures/ice-brent/intraday",
    "/v1/futures/ice-brent/spreads",
    "/v1/futures/ice-brent/curve",
    "/v1/futures/ice-brent/spread-history",
  ];

  for (const path of paths) {
    const harness = createHarness();
    configure(harness);
    harness.queue(200, JSON.stringify({ status: "success", data: { ok: true } }));
    const result = harness.context.OILPRICE_GET(path, "");
    assert.equal(result[0][0], "Field", path);
    assert.equal(harness.requests.length, 1, path);
  }
});

test("OILPRICE_GET renders nested futures, diesel, and price-hash responses", () => {
  const futures = createHarness();
  configure(futures);
  futures.queue(
    200,
    JSON.stringify({
      contracts: [
        {
          contract_month: "2026-09",
          daily_data: [
            { date: "2026-07-23", open: "81.10", close: "82.25" },
          ],
        },
      ],
    }),
  );
  assert.deepEqual(
    JSON.parse(
      JSON.stringify(
        futures.context.OILPRICE_GET("/v1/futures/ice-brent/ohlc", ""),
      ),
    ),
    [
      ["contract_month", "date", "open", "close"],
      ["2026-09", "2026-07-23", 81.1, 82.25],
    ],
  );

  const diesel = createHarness();
  configure(diesel);
  diesel.queue(
    200,
    JSON.stringify({
      data: {
        regional_average: {
          region: "US",
          price: "3.45",
          currency: "USD",
        },
      },
    }),
  );
  assert.deepEqual(
    JSON.parse(
      JSON.stringify(diesel.context.OILPRICE_GET("/v1/diesel-prices", "")),
    ),
    [
      ["Field", "Value"],
      ["region", "US"],
      ["price", 3.45],
      ["currency", "USD"],
    ],
  );

  const priceHash = createHarness();
  configure(priceHash);
  priceHash.queue(
    200,
    JSON.stringify({
      data: {
        prices: {
          WTI_USD: { price: "81.50", unit: "barrel" },
          BRENT_CRUDE_USD: { price: 84.2, unit: "barrel" },
        },
      },
    }),
  );
  assert.deepEqual(
    JSON.parse(
      JSON.stringify(priceHash.context.OILPRICE_GET("/v1/prices/all", "")),
    ),
    [
      ["Code", "price", "unit"],
      ["WTI_USD", 81.5, "barrel"],
      ["BRENT_CRUDE_USD", 84.2, "barrel"],
    ],
  );
});

test("nested and encoded credential query keys are rejected", () => {
  for (const query of [
    "filters[token]=secret",
    "filters%5Bapi_key%5D=secret",
    "access-token=secret",
    "client.secret=secret",
  ]) {
    const harness = createHarness();
    configure(harness);
    const result = harness.context.OILPRICE_GET("/v1/prices/latest", query);
    assert.equal(result[0][0], "#UNSUPPORTED_QUERY", query);
    assert.equal(harness.requests.length, 0, query);
  }
});

test("a stale cache envelope is discarded before formula refresh", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(200, latestBody({ price: 80 }));
  assert.equal(harness.context.OILPRICE("WTI_USD"), 80);

  const activeCache = harness.documentCache.size
    ? harness.documentCache
    : harness.userCache;
  const cacheKey = [...activeCache.keys()][0];
  const envelope = JSON.parse(activeCache.get(cacheKey));
  envelope.cachedAt = Date.now() - 301000;
  activeCache.set(cacheKey, JSON.stringify(envelope));
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
  assert.match(
    invalid.context.OILPRICE_HISTORY("WTI_USD", 1)[0][1],
    /timestamp/i,
  );
});

test("bunker records preserve the complete source contract", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(200, bunkerBody([bunkerRecord()]));

  assert.deepEqual(
    JSON.parse(JSON.stringify(harness.context.fetchBunkerRecords_(""))),
    [
      {
        port: "SINGAPORE",
        fuelType: "VLSFO",
        price: 612.5,
        currency: "USD",
        unit: "metric_ton",
        region: "Asia",
        source: "fixture_provider",
        timestamp: "2026-07-25T09:30:00.000Z",
      },
    ],
  );
  assert.match(harness.requests[0].url, /\/v1\/prices\/data-connector$/);
});

test("bunker records reject empty, malformed, and incomplete successful responses", () => {
  for (const [body, pattern] of [
    [bunkerBody([]), /no usable bunker price/i],
    [JSON.stringify({ status: "success", data: { prices: [null] } }), /no usable bunker price/i],
    [bunkerBody([bunkerRecord({ price: "not-a-number" })]), /finite price/i],
    [bunkerBody([bunkerRecord({ port: "" })]), /missing port/i],
    [bunkerBody([bunkerRecord({ fuel_type: "" })]), /missing fuel type/i],
    [bunkerBody([bunkerRecord({ currency: "" })]), /missing currency/i],
    [bunkerBody([bunkerRecord({ unit: "" })]), /missing unit/i],
    [bunkerBody([bunkerRecord({ as_of: "not-a-date" })]), /source timestamp/i],
  ]) {
    const harness = createHarness();
    configure(harness);
    harness.queue(200, body);
    assert.throws(() => harness.context.fetchBunkerRecords_(""), pattern);
  }
});

for (const [status, pattern] of [
  [401, /invalid or revoked API key/i],
  [402, /pricing/i],
  [403, /cannot access the requested dataset/i],
  [429, /rate or quota limit/i],
]) {
  test(`BUNKER_PRICE maps HTTP ${status} to actionable recovery`, () => {
    const harness = createHarness();
    configure(harness);
    harness.queue(status, JSON.stringify({ error: "fixture" }));
    assert.match(
      harness.context.BUNKER_PRICE("SINGAPORE", "VLSFO"),
      pattern,
    );
  });
}

test("BUNKER_PRICE maps Apps Script fetch failures to timeout recovery", () => {
  const harness = createHarness();
  configure(harness);
  harness.queueError(new Error("Socket timeout"));
  assert.match(
    harness.context.BUNKER_PRICE("SINGAPORE", "VLSFO"),
    /timed out/i,
  );
});

test("bunker formulas normalize and encode filters and return source-aware values", () => {
  const priceHarness = createHarness();
  configure(priceHarness);
  priceHarness.queue(
    200,
    bunkerBody([bunkerRecord({ port: "NEW-YORK:US", fuel_type: "VLSFO_0-5" })]),
  );
  assert.equal(
    priceHarness.context.BUNKER_PRICE("new-york:us", "vlsfo_0-5"),
    612.5,
  );
  assert.match(
    priceHarness.requests[0].url,
    /\?port=NEW-YORK%3AUS&fuel_type=VLSFO_0-5$/,
  );

  const tableHarness = createHarness();
  configure(tableHarness);
  tableHarness.queue(
    200,
    bunkerBody([
      bunkerRecord(),
      bunkerRecord({
        fuel_type: "MGO",
        price: "785.20",
        unit: "tonne",
        as_of: "2026-07-25T10:00:00.000Z",
      }),
    ]),
  );
  assert.deepEqual(
    JSON.parse(
      JSON.stringify(tableHarness.context.BUNKER_PORT_PRICES("singapore")),
    ),
    [
      ["Fuel Type", "Price", "Currency", "Unit", "Source Timestamp"],
      ["VLSFO", 612.5, "USD", "metric_ton", "2026-07-25T09:30:00.000Z"],
      ["MGO", 785.2, "USD", "tonne", "2026-07-25T10:00:00.000Z"],
    ],
  );
  assert.match(tableHarness.requests[0].url, /\?port=SINGAPORE$/);
});

test("bunker inputs reject unsupported filter characters before fetch", () => {
  const harness = createHarness();
  configure(harness);
  assert.match(
    harness.context.BUNKER_PRICE("Singapore & Johor", "VLSFO"),
    /unsupported characters/i,
  );
  assert.equal(harness.requests.length, 0);
});

test("Data Connector sheet writer creates the nine-column source-aware table", () => {
  const harness = createHarness();
  const spreadsheet = attachSpreadsheet(harness, false);
  const records = [
    {
      port: "SINGAPORE",
      fuelType: "VLSFO",
      price: 612.5,
      currency: "USD",
      unit: "metric_ton",
      region: "Asia",
      source: "fixture_provider",
      timestamp: "2026-07-25T09:30:00.000Z",
    },
  ];

  harness.context.writeToDataConnectorSheet(records);

  assert.deepEqual(spreadsheet.calls.slice(0, 3), [
    ["getSheetByName", "Bunker Prices"],
    ["insertSheet", "Bunker Prices"],
    ["clear"],
  ]);
  const header = spreadsheet.ranges.find(
    (range) => range.row === 1 && range.column === 1,
  );
  assert.deepEqual(header.values, [
    [
      "Port",
      "Fuel Type",
      "Price",
      "Currency",
      "Unit",
      "Region",
      "Source",
      "Source Timestamp",
      "Retrieved At",
    ],
  ]);
  const headerStyle = spreadsheet.ranges.find(
    (range) =>
      range.row === 1 &&
      range.column === 1 &&
      range.fontWeight === "bold",
  );
  assert.equal(headerStyle.background, "#e8f5e9");

  const rows = spreadsheet.ranges.find(
    (range) => range.row === 2 && range.column === 1,
  );
  assert.equal(rows.rowCount, 1);
  assert.equal(rows.columnCount, 9);
  assert.deepEqual(rows.values[0].slice(0, 8), [
    "SINGAPORE",
    "VLSFO",
    612.5,
    "USD",
    "metric_ton",
    "Asia",
    "fixture_provider",
    "2026-07-25T09:30:00.000Z",
  ]);
  assert.ok(Number.isFinite(new Date(rows.values[0][8]).getTime()));

  const priceRange = spreadsheet.ranges.find(
    (range) => range.row === 2 && range.column === 3,
  );
  assert.equal(priceRange.numberFormat, "#,##0.00");
  assert.deepEqual(spreadsheet.calls.slice(-2), [
    ["autoResizeColumns", 1, 9],
    ["activate"],
  ]);
});

test("Data Connector menu flow alerts on success and recovers from empty data", () => {
  const success = createHarness();
  configure(success);
  const successSpreadsheet = attachSpreadsheet(success);
  success.queue(200, bunkerBody([bunkerRecord()]));
  success.context.fetchDataConnectorPrices();
  assert.deepEqual(successSpreadsheet.alerts, [
    "Fetched 1 source-timestamped bunker price records.",
  ]);

  const empty = createHarness();
  configure(empty);
  const emptySpreadsheet = attachSpreadsheet(empty, false);
  empty.queue(200, bunkerBody([]));
  empty.context.fetchDataConnectorPrices();
  assert.equal(emptySpreadsheet.alerts.length, 1);
  assert.match(emptySpreadsheet.alerts[0], /no usable bunker price/i);
  assert.equal(
    emptySpreadsheet.calls.some((call) => call[0] === "insertSheet"),
    false,
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
      sourceDescription: "",
      timestamp: "2026-07-19T14:10:50.373Z",
      collectedAt: "",
      formatted: "",
      dataStatus: "",
      stale: "",
      ageDays: "",
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

test("user info does not infer a quota window from legacy monthly fields", () => {
  const harness = createHarness();
  configure(harness);
  harness.queue(
    200,
    JSON.stringify({
      data: {
        tier: "free",
        request_limit: 50,
        requests_this_month: 5,
      },
    }),
  );

  const info = harness.context.getUserInfo();
  assert.equal(info.tier, "free");
  assert.equal(info.limit, null);
  assert.equal(info.used, null);
  assert.equal(info.window, null);
});
