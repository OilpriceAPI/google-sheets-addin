const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");
const products = JSON.parse(
  fs.readFileSync(path.join(ROOT, "portfolio", "products.json"), "utf8"),
);
const releases = JSON.parse(
  fs.readFileSync(path.join(ROOT, "portfolio", "releases.json"), "utf8"),
);

function load(productId, overrides = {}) {
  const code = fs.readFileSync(
    path.join(ROOT, "portfolio", "dist", productId, "Code.gs"),
    "utf8",
  );
  const documentValues = new Map();
  const userValues = new Map();
  const requests = [];
  const responses = [];
  let activeSpreadsheetId = "fixture-spreadsheet-id";
  const properties = {
    getProperty: (key) => documentValues.get(key) || null,
    setProperty: (key, value) => documentValues.set(key, value),
    deleteProperty: (key) => documentValues.delete(key),
  };
  const userProperties = {
    getProperty: (key) => userValues.get(key) || null,
    setProperty: (key, value) => userValues.set(key, value),
    deleteProperty: (key) => userValues.delete(key),
  };
  const context = {
    Date,
    Error,
    HtmlService: {},
    JSON,
    Math,
    Number,
    Object,
    PropertiesService: {
      getDocumentProperties: () => properties,
      getUserProperties: () => userProperties,
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => ({ getId: () => activeSpreadsheetId }),
    },
    String,
    UrlFetchApp: {
      fetch(url, options) {
        requests.push({ url, options });
        const response = responses.shift();
        if (!response) throw new Error("No API fixture queued");
        return {
          getResponseCode: () => response.status,
          getContentText: () => response.body,
        };
      },
    },
    decodeURIComponent,
    encodeURIComponent,
    ...overrides,
  };
  vm.createContext(context);
  new vm.Script(code, { filename: `${productId}/Code.gs` }).runInContext(context);
  return {
    context,
    documentValues,
    userValues,
    requests,
    setActiveSpreadsheetId(value) {
      activeSpreadsheetId = value;
    },
    queue(status, body) {
      responses.push({ status, body });
    },
  };
}

function spreadsheetHarness() {
  const sheets = new Map();
  const writes = [];
  function range(reference) {
    const state = { reference };
    writes.push(state);
    const api = {
      setBackground(value) { state.background = value; return api; },
      setFontColor(value) { state.fontColor = value; return api; },
      setFontSize(value) { state.fontSize = value; return api; },
      setFontWeight(value) { state.fontWeight = value; return api; },
      setFormula(value) { state.formula = value; return api; },
      setFormulas(value) { state.formulas = value; return api; },
      setNumberFormat(value) { state.numberFormat = value; return api; },
      setValue(value) { state.value = value; return api; },
      setValues(value) { state.values = value; return api; },
    };
    return api;
  }
  function createSheet(name) {
    return {
      name,
      activate() { return this; },
      autoResizeColumns() { return this; },
      clear() { return this; },
      getRange(...args) { return range(args); },
    };
  }
  const workbook = {
    getId: () => "fixture-spreadsheet-id",
    getSheetByName: (name) => sheets.get(name) || null,
    insertSheet(name) {
      const sheet = createSheet(name);
      sheets.set(name, sheet);
      return sheet;
    },
  };
  return {
    sheets,
    writes,
    SpreadsheetApp: {
      getActiveSpreadsheet: () => workbook,
    },
  };
}

function pricePayload(codes) {
  return JSON.stringify({
    data: {
      prices: codes.map((code, index) => ({
        code,
        price: index + 2.5,
        currency: code === "DUTCH_TTF_EUR" ? "EUR" : "USD",
        unit: code.includes("DIESEL") || code.includes("GASOLINE") || code.includes("HEATING")
          ? "gallon"
          : code.includes("VLSFO") || code.includes("MGO")
            ? "metric_ton"
            : code === "DUTCH_TTF_EUR"
              ? "mwh"
              : "mmbtu",
        source: "fixture",
        created_at: "2026-07-30T00:00:00Z",
      })),
    },
  });
}

test("all five builds compile and expose unique product identities", () => {
  assert.equal(products.length, 5);
  const identities = products.map((product) => {
    const harness = load(product.id);
    assert.equal(typeof harness.context[product.builder], "function");
    assert.equal(harness.context.getSidebarState().product, product.name);
    return product.activationHeader;
  });
  assert.equal(new Set(identities).size, 5);
});

test("release records do not claim superseded immutable packages are current", () => {
  for (const product of products) {
    const release = releases.products[product.id];
    assert.equal(release.stage, "source-candidate");
    assert.equal(release.version, null);
    assert.equal(Number.isInteger(release.previousImmutableVersion), true);
  }
});

test("crack-spread math covers 3-2-1 and 2-1-1 conventions", () => {
  const { context } = load("crack-spread-lab");
  assert.ok(Math.abs(context.calculateCrackSpread_(75, 2.5, 2.4, "3-2-1") - 28.6) < 1e-9);
  assert.ok(Math.abs(context.calculateCrackSpread_(75, 2.5, 2.4, "2-1-1") - 27.9) < 1e-9);
  assert.throws(() => context.calculateCrackSpread_(75, 2.5, 2.4, "5-3-2"), /Supported/);
});

test("bunker voyage model calculates tonnes, blend, and total cost", () => {
  const { context } = load("bunker-voyage-planner");
  const result = context.calculateVoyageFuelCost_(10, 25, 2, 5, 0.8, 600, 800);
  assert.equal(result.tonnes, 260);
  assert.equal(result.blendedPrice, 640);
  assert.equal(result.totalCost, 166400);
  assert.throws(() => context.calculateVoyageFuelCost_(10, 25, 2, 5, 1.2, 600, 800), /between 0 and 1/);
  assert.throws(() => context.calculateVoyageFuelCost_(-1, 25, 2, 5, 0.8, 600, 800), /non-negative/);
  assert.throws(() => context.calculateVoyageFuelCost_(10, 25, 2, 5, 0.8, -600, 800), /non-negative/);
});

test("fuel surcharge schedule never produces a negative surcharge", () => {
  const { context } = load("fuel-surcharge-studio");
  assert.equal(context.surchargePerMile_(3.5, 2.5, 5), 0.2);
  assert.equal(context.surchargePerMile_(2, 2.5, 5), 0);
  const rows = context.surchargeBands_(2.5, 5, 2, 0.25, 4);
  assert.equal(rows.length, 5);
  assert.equal(rows[1][2], 0);
  assert.throws(() => context.surchargePerMile_(-1, 2.5, 5), /non-negative/);
  assert.throws(() => context.surchargeBands_(2.5, 5, 2, 0, 4), /positive/);
});

test("curve builder uses the production endpoint and top-level contracts schema", () => {
  const harness = load("energy-curve-builder");
  const { context } = harness;
  const records = context.curveContracts_({
    analysis_date: "2026-07-29",
    contracts: [
      { contract_month: "2026-09", settlement_price: 70, trading_date: "2026-07-29" },
      { contract_month: "2026-10", settlement_price: 71, trading_date: "2026-07-29" },
    ],
  });
  const rows = context.curveTable_(records);
  assert.equal(rows[2][2], 1);
  assert.equal(rows[2][3], "Contango step");
  assert.throws(
    () => context.curveContracts_({ contracts: [{ contract_month: "2026-09", settlement_price: 70 }] }),
    /at least two/,
  );

  harness.userValues.set(
    "OILPRICEAPI_KEY:fixture-spreadsheet-id",
    "fixture-key",
  );
  harness.queue(
    200,
    JSON.stringify({
      analysis_date: "2026-07-29",
      contracts: [
        { contract_month: "2026-09", settlement_price: 70, trading_date: "2026-07-29" },
        { contract_month: "2026-10", settlement_price: 71, trading_date: "2026-07-29" },
      ],
    }),
  );
  context.fetchCurve_("ice-wti");
  assert.equal(
    harness.requests[0].url,
    "https://api.oilpriceapi.com/v1/futures/ice-wti/curve",
  );
  assert.throws(() => context.fetchCurve_("CL"), /ice-wti and ice-brent/);
});

test("gas spread normalization makes currency and energy conversion explicit", () => {
  const { context } = load("gas-spread-monitor");
  const result = context.normalizeGasMarkets_(3, 34.12141633, 12, 1);
  assert.ok(Math.abs(result.ttf - 10) < 1e-9);
  assert.ok(Math.abs(result.ttfHenryHub - 7) < 1e-9);
  assert.ok(Math.abs(result.jkmTtf - 2) < 1e-9);
});

test("API requests carry measurable product identity without sheet contents", () => {
  const harness = load("crack-spread-lab");
  harness.documentValues.set("OILPRICEAPI_KEY", "fixture-key");
  harness.queue(
    200,
    JSON.stringify({
      data: {
        prices: [
          { code: "WTI_USD", price: 75, currency: "USD", unit: "barrel", created_at: "2026-07-29T00:00:00Z" },
        ],
      },
    }),
  );
  const records = harness.context.latestPrices_(["WTI_USD"]);
  assert.equal(records[0].price, 75);
  assert.equal(
    harness.requests[0].options.headers["X-API-Client"],
    "crack-spread-lab/1.0.0",
  );
  assert.equal(
    harness.requests[0].options.headers["X-OilPriceAPI-Client"],
    undefined,
  );
  assert.equal(JSON.stringify(harness.requests[0]).includes("spreadsheet"), false);
  assert.equal(JSON.stringify(records).includes("fixture-key"), false);
});

test("API keys are spreadsheet-scoped and remain available to custom-function context", () => {
  const harness = load("crack-spread-lab");
  harness.context.saveApiKey(" fixture-key ");
  assert.equal(
    harness.documentValues.get("OILPRICEAPI_KEY"),
    "fixture-key",
  );
  assert.equal(
    harness.userValues.get("OILPRICEAPI_KEY:fixture-spreadsheet-id"),
    "fixture-key",
  );

  harness.documentValues.delete("OILPRICEAPI_KEY");
  assert.equal(harness.context.getApiKey_(), "fixture-key");

  harness.setActiveSpreadsheetId("different-spreadsheet-id");
  assert.equal(harness.context.getApiKey_(), null);

  harness.setActiveSpreadsheetId("fixture-spreadsheet-id");
  harness.context.deleteApiKey();
  assert.equal(
    harness.userValues.has("OILPRICEAPI_KEY:fixture-spreadsheet-id"),
    false,
  );
});

test("unscoped prototype keys cannot authorize another spreadsheet", () => {
  const harness = load("crack-spread-lab");
  harness.userValues.set("OILPRICEAPI_KEY", "prototype-key");
  harness.setActiveSpreadsheetId("sheet-a");
  assert.equal(harness.context.getApiKey_(), null);
  harness.setActiveSpreadsheetId("sheet-b");
  assert.equal(harness.context.getApiKey_(), null);
});

test("API failures produce actionable recovery messages and reject silent bad data", () => {
  const harness = load("crack-spread-lab");
  harness.documentValues.set("OILPRICEAPI_KEY", "fixture-key");

  harness.queue(401, JSON.stringify({ error: "Unauthorized" }));
  assert.throws(
    () => harness.context.latestPrices_(["WTI_USD"]),
    /invalid or revoked/,
  );

  harness.queue(403, JSON.stringify({ error: "Forbidden" }));
  assert.throws(
    () => harness.context.latestPrices_(["WTI_USD"]),
    /not enabled/,
  );

  harness.queue(429, JSON.stringify({ error: "Rate limited" }));
  assert.throws(
    () => harness.context.latestPrices_(["WTI_USD"]),
    /rate or quota limit/,
  );

  harness.queue(200, "not-json");
  assert.throws(
    () => harness.context.latestPrices_(["WTI_USD"]),
    /unreadable response/,
  );

  harness.queue(
    200,
    JSON.stringify({
      data: { prices: [{ code: "WTI_USD", price: null }] },
    }),
  );
  assert.throws(
    () => harness.context.latestPrices_(["WTI_USD"]),
    /finite market price/,
  );
});

test("reviewed market catalogs and first-party URL boundary are enforced", () => {
  const harness = load("gas-spread-monitor");
  harness.documentValues.set("OILPRICEAPI_KEY", "fixture-key");
  assert.throws(() => harness.context.normalizeCode_("WTI_USD"), /outside this product/);
  assert.throws(
    () => harness.context.requestJson_("https://example.com", "fixture-key"),
    /Unsupported OilPriceAPI path/,
  );
  assert.throws(
    () => harness.context.requestJson_("/../users/me", "fixture-key"),
    /Unsupported OilPriceAPI path/,
  );
});

test("all five customer-critical workbook builds complete with expected tabs and formulas", () => {
  for (const product of products) {
    const workbook = spreadsheetHarness();
    const harness = load(product.id, {
      SpreadsheetApp: workbook.SpreadsheetApp,
    });
    harness.documentValues.set("OILPRICEAPI_KEY", "fixture-key");

    if (product.id === "energy-curve-builder") {
      for (const market of ["WTI", "BRENT"]) {
        harness.queue(
          200,
          JSON.stringify({
            analysis_date: "2026-07-29",
            contracts: [
              {
                contract_month: "2026-09",
                contract_code: `${market}_FUTURES_2026_09`,
                settlement_price: 84,
                trading_date: "2026-07-29",
              },
              {
                contract_month: "2026-10",
                contract_code: `${market}_FUTURES_2026_10`,
                settlement_price: 82,
                trading_date: "2026-07-29",
              },
            ],
          }),
        );
      }
    } else {
      harness.queue(200, pricePayload(product.allowedCodes));
    }

    if (product.id === "crack-spread-lab") {
      for (const code of product.allowedCodes) {
        harness.queue(200, pricePayload([code]));
      }
    }

    const result = harness.context[product.builder]();
    assert.equal(result.success, true, product.id);
    assert.deepEqual(
      [...workbook.sheets.keys()].sort(),
      [...product.sheets].sort(),
      `${product.id} sheet set`,
    );
    assert.ok(
      workbook.writes.some((write) => write.formula || write.formulas),
      `${product.id} must create live spreadsheet formulas`,
    );
    assert.ok(
      harness.documentValues.has("OILPRICEAPI_ACTIVATED_AT"),
      `${product.id} activation marker`,
    );
  }
});
