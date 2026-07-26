#!/usr/bin/env node

const assert = require("node:assert/strict");
const { spawnSync } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const apiKey = process.env.OILPRICEAPI_KEY;
if (!apiKey) {
  throw new Error("Set OILPRICEAPI_KEY before running npm run test:live.");
}
const dataConnectorSmoke =
  process.env.OILPRICEAPI_DATA_CONNECTOR_SMOKE === "1";
const dataConnectorPort = process.env.OILPRICEAPI_DATA_CONNECTOR_PORT;
const dataConnectorFuel = process.env.OILPRICEAPI_DATA_CONNECTOR_FUEL;

function curlConfigValue(value) {
  return String(value).replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}

function fetchWithCurl(url, options) {
  const config = [
    `url = "${curlConfigValue(url)}"`,
    'request = "GET"',
    `header = "Authorization: ${curlConfigValue(options.headers.Authorization)}"`,
    'header = "Accept: application/json"',
    "silent",
    "show-error",
    'write-out = "\\n__HTTP_STATUS__:%{http_code}"',
  ].join("\n");
  const result = spawnSync("curl", ["--config", "-"], {
    encoding: "utf8",
    env: { HOME: process.env.HOME, PATH: process.env.PATH },
    input: config,
  });
  if (result.status !== 0) {
    throw new Error("Production request failed before receiving an HTTP response.");
  }
  const marker = "\n__HTTP_STATUS__:";
  const splitAt = result.stdout.lastIndexOf(marker);
  if (splitAt < 0) throw new Error("Production response did not include status metadata.");
  const body = result.stdout.slice(0, splitAt);
  const status = Number(result.stdout.slice(splitAt + marker.length));
  return {
    getContentText: () => body,
    getResponseCode: () => status,
  };
}

const cache = new Map();
const connectorAlerts = [];
const connectorRanges = [];
const connectorSheet = {
  activate: () => undefined,
  autoResizeColumns: () => undefined,
  clear: () => undefined,
  getRange(row, column, rowCount, columnCount) {
    assert.ok(rowCount > 0 && columnCount > 0);
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
      setValues(values) {
        state.values = values;
        return range;
      },
    };
    connectorRanges.push(state);
    return range;
  },
};
const code = fs.readFileSync(path.join(__dirname, "..", "Code.gs"), "utf8");
const context = {
  CacheService: {
    getUserCache: () => ({
      get: (key) => cache.get(key) || null,
      put: (key, value) => cache.set(key, value),
      remove: (key) => cache.delete(key),
    }),
  },
  Date,
  Error,
  JSON,
  Math,
  Number,
  PropertiesService: {
    getDocumentProperties: () => ({
      getProperty: () => apiKey,
      setProperty: () => undefined,
      deleteProperty: () => undefined,
    }),
    getUserProperties: () => ({
      getProperty: () => null,
      setProperty: () => undefined,
      deleteProperty: () => undefined,
    }),
  },
  SpreadsheetApp: {
    getActiveSpreadsheet: () => ({
      getSheetByName: () => connectorSheet,
      insertSheet: () => connectorSheet,
    }),
    getUi: () => ({
      alert: (message) => connectorAlerts.push(message),
    }),
  },
  String,
  UrlFetchApp: { fetch: fetchWithCurl },
  console,
  decodeURIComponent,
  encodeURIComponent,
};
vm.createContext(context);
new vm.Script(code, { filename: "Code.gs" }).runInContext(context);

const price = context.OILPRICE("WTI_USD");
assert.ok(Number.isFinite(price) && price > 0);
assert.equal(context.OILPRICE("WTI_USD"), price, "cached formula value changed");
assert.equal(context.OILPRICE_PRICE("WTI_USD"), price);

const unit = context.OILPRICE_UNIT("WTI_USD");
assert.ok(typeof unit === "string" && unit.includes("/") && !unit.startsWith("#"));

const info = context.OILPRICE_INFO("WTI_USD");
assert.ok(Array.isArray(info) && info.length > 1);
const infoMap = Object.fromEntries(info.slice(1));
assert.equal(infoMap.price, price);
assert.ok(typeof infoMap.source_timestamp === "string");
assert.ok(Number.isFinite(new Date(infoMap.source_timestamp).getTime()));

const latestTable = context.OILPRICE_GET(
  "/v1/prices/latest",
  "by_code=WTI_USD",
);
assert.ok(Array.isArray(latestTable) && latestTable.length > 1);

const history = context.OILPRICE_HISTORY("WTI_USD", 1);
assert.ok(Array.isArray(history) && history.length > 0);
assert.ok(Number.isFinite(new Date(history[0][0]).getTime()));
assert.ok(Number.isFinite(history[0][1]));

if (dataConnectorSmoke) {
  if (!dataConnectorPort || !dataConnectorFuel) {
    throw new Error(
      "Set OILPRICEAPI_DATA_CONNECTOR_PORT and OILPRICEAPI_DATA_CONNECTOR_FUEL when OILPRICEAPI_DATA_CONNECTOR_SMOKE=1.",
    );
  }

  context.fetchDataConnectorPrices();
  assert.equal(connectorAlerts.length, 1);
  assert.match(connectorAlerts[0], /^Fetched \d+ source-timestamped bunker price records\.$/);
  const headerRange = connectorRanges.find(
    (range) => range.row === 1 && range.column === 1 && range.values,
  );
  const headerStyle = connectorRanges.find(
    (range) => range.row === 1 && range.column === 1 && range.fontWeight,
  );
  const dataRange = connectorRanges.find(
    (range) => range.row === 2 && range.column === 1 && range.values,
  );
  assert.equal(headerRange.columnCount, 9);
  assert.equal(headerStyle.fontWeight, "bold");
  assert.equal(headerStyle.background, "#e8f5e9");
  assert.ok(dataRange.rowCount > 0);
  assert.ok(dataRange.values.every((row) => row.length === 9));

  const bunkerPrice = context.BUNKER_PRICE(
    dataConnectorPort,
    dataConnectorFuel,
  );
  assert.ok(Number.isFinite(bunkerPrice) && bunkerPrice > 0);

  const bunkerTable = context.BUNKER_PORT_PRICES(dataConnectorPort);
  assert.deepEqual(
    JSON.parse(JSON.stringify(bunkerTable[0])),
    ["Fuel Type", "Price", "Currency", "Unit", "Source Timestamp"],
  );
  assert.ok(bunkerTable.length > 1);
  assert.ok(
    bunkerTable.slice(1).every(
      (row) =>
        row.length === 5 &&
        Number.isFinite(row[1]) &&
        typeof row[2] === "string" &&
        typeof row[3] === "string" &&
        Number.isFinite(new Date(row[4]).getTime()),
    ),
  );

  console.log(
    `Production Data Connector smoke passed: menu_rows=${dataRange.rowCount}, formula_price=${bunkerPrice}, port_rows=${bunkerTable.length - 1}`,
  );
} else {
  console.log(
    "Production Data Connector smoke skipped; set OILPRICEAPI_DATA_CONNECTOR_SMOKE=1 with port and fuel inputs to run the entitled-account checks.",
  );
}

console.log(
  `Production formula smoke passed: latest=${price}, unit=${unit}, info_source_timestamp=${infoMap.source_timestamp}, history_records=${history.length}`,
);
