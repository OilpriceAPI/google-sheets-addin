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

function createHarness() {
  const documentPropertyValues = new Map();
  const userPropertyValues = new Map();
  const cache = new Map();
  const responses = [];
  const requests = [];

  const propertyStore = (values) => ({
    getProperty: (key) => values.get(key) || null,
    setProperty: (key, value) => values.set(key, value),
    deleteProperty: (key) => values.delete(key),
  });
  const documentProperties = propertyStore(documentPropertyValues);
  const userProperties = propertyStore(userPropertyValues);
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
    PropertiesService: {
      getDocumentProperties: () => documentProperties,
      getUserProperties: () => userProperties,
    },
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
    cache,
    context,
    documentPropertyValues,
    queue(status, body, headers = {}) {
      responses.push({ status, body, headers });
    },
    queueError(error) {
      responses.push(error);
    },
    requests,
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

test("legacy user-property key remains readable until it is saved per spreadsheet", () => {
  const harness = createHarness();
  harness.userPropertyValues.set("OILPRICEAPI_KEY", "legacy-key");
  harness.queue(200, latestBody());

  assert.equal(harness.context.OILPRICE_PRICE("WTI_USD"), 81.78);
  assert.equal(
    harness.requests[0].options.headers.Authorization,
    "Token legacy-key",
  );

  harness.context.saveApiKey("spreadsheet-key");
  assert.equal(harness.userPropertyValues.has("OILPRICEAPI_KEY"), false);
  assert.equal(
    harness.documentPropertyValues.get("OILPRICEAPI_KEY"),
    "spreadsheet-key",
  );
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
  [402, /cannot access the requested dataset/i],
  [403, /cannot access the requested dataset/i],
  [429, /rate or quota limit/i],
]) {
  test(`BUNKER_PRICE maps HTTP ${status} to actionable recovery`, () => {
    const harness = createHarness();
    configure(harness);
    harness.queue(status, JSON.stringify({ error: "fixture" }));
    assert.throws(
      () => harness.context.BUNKER_PRICE("SINGAPORE", "VLSFO"),
      pattern,
    );
  });
}

test("BUNKER_PRICE maps Apps Script fetch failures to timeout recovery", () => {
  const harness = createHarness();
  configure(harness);
  harness.queueError(new Error("Socket timeout"));
  assert.throws(
    () => harness.context.BUNKER_PRICE("SINGAPORE", "VLSFO"),
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
  assert.throws(
    () => harness.context.BUNKER_PRICE("Singapore & Johor", "VLSFO"),
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
