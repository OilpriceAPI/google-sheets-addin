const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");
const products = JSON.parse(
  fs.readFileSync(path.join(ROOT, "portfolio", "products.json"), "utf8"),
);

function load(productId, overrides = {}) {
  const code = fs.readFileSync(
    path.join(ROOT, "portfolio", "dist", productId, "Code.gs"),
    "utf8",
  );
  const documentValues = new Map();
  const requests = [];
  const responses = [];
  const properties = {
    getProperty: (key) => documentValues.get(key) || null,
    setProperty: (key, value) => documentValues.set(key, value),
    deleteProperty: (key) => documentValues.delete(key),
  };
  const context = {
    Date,
    Error,
    HtmlService: {},
    JSON,
    Math,
    Number,
    Object,
    PropertiesService: { getDocumentProperties: () => properties },
    SpreadsheetApp: {},
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
    requests,
    queue(status, body) {
      responses.push({ status, body });
    },
  };
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
});

test("fuel surcharge schedule never produces a negative surcharge", () => {
  const { context } = load("fuel-surcharge-studio");
  assert.equal(context.surchargePerMile_(3.5, 2.5, 5), 0.2);
  assert.equal(context.surchargePerMile_(2, 2.5, 5), 0);
  const rows = context.surchargeBands_(2.5, 5, 2, 0.25, 4);
  assert.equal(rows.length, 5);
  assert.equal(rows[1][2], 0);
});

test("curve parser rejects incomplete data and labels structure", () => {
  const { context } = load("energy-curve-builder");
  const records = context.curveContracts_({
    data: { contracts: [{ month: "2026-09", price: 70 }, { month: "2026-10", price: 71 }] },
  });
  const rows = context.curveTable_(records);
  assert.equal(rows[2][2], 1);
  assert.equal(rows[2][3], "Contango step");
  assert.throws(() => context.curveContracts_({ data: { contracts: [{ month: "2026-09", price: 70 }] } }), /at least two/);
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
    harness.requests[0].options.headers["X-OilPriceAPI-Client"],
    "crack-spread-lab/0.1.0",
  );
  assert.equal(JSON.stringify(harness.requests[0]).includes("spreadsheet"), false);
  assert.equal(JSON.stringify(records).includes("fixture-key"), false);
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
