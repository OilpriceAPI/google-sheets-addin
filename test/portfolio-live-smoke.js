#!/usr/bin/env node

const assert = require("node:assert/strict");
const { spawnSync } = require("node:child_process");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");
const apiKey = process.env.OILPRICEAPI_KEY;
if (!apiKey) {
  throw new Error("Set OILPRICEAPI_KEY to a non-customer test key.");
}
const products = JSON.parse(
  fs.readFileSync(path.join(ROOT, "portfolio", "products.json"), "utf8"),
);

function curlConfigValue(value) {
  return String(value).replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}

function fetchWithCurl(url, options) {
  const config = [
    `url = "${curlConfigValue(url)}"`,
    'request = "GET"',
    `header = "Authorization: ${curlConfigValue(options.headers.Authorization)}"`,
    'header = "Accept: application/json"',
    `header = "X-API-Client: ${curlConfigValue(options.headers["X-API-Client"])}"`,
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
  if (splitAt < 0) {
    throw new Error("Production response did not include status metadata.");
  }
  const body = result.stdout.slice(0, splitAt);
  const status = Number(result.stdout.slice(splitAt + marker.length));
  return {
    getContentText: () => body,
    getResponseCode: () => status,
  };
}

function workbookHarness(productId) {
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
      activate() { return this; },
      autoResizeColumns() { return this; },
      clear() { return this; },
      getRange(...args) { return range(args); },
      name,
    };
  }
  const workbook = {
    getId: () => `live-smoke-${productId}`,
    getSheetByName: (name) => sheets.get(name) || null,
    insertSheet(name) {
      const sheet = createSheet(name);
      sheets.set(name, sheet);
      return sheet;
    },
  };
  return { sheets, writes, workbook };
}

for (const product of products) {
  const documentValues = new Map([["OILPRICEAPI_KEY", apiKey]]);
  const userValues = new Map();
  const workbook = workbookHarness(product.id);
  const requests = [];
  const code = fs.readFileSync(
    path.join(ROOT, "portfolio", "dist", product.id, "Code.gs"),
    "utf8",
  );
  const context = {
    Date,
    Error,
    HtmlService: {},
    JSON,
    Math,
    Number,
    Object,
    PropertiesService: {
      getDocumentProperties: () => ({
        getProperty: (key) => documentValues.get(key) || null,
        setProperty: (key, value) => documentValues.set(key, value),
        deleteProperty: (key) => documentValues.delete(key),
      }),
      getUserProperties: () => ({
        getProperty: (key) => userValues.get(key) || null,
        setProperty: (key, value) => userValues.set(key, value),
        deleteProperty: (key) => userValues.delete(key),
      }),
    },
    SpreadsheetApp: {
      getActiveSpreadsheet: () => workbook.workbook,
    },
    String,
    UrlFetchApp: {
      fetch(url, options) {
        requests.push(url);
        return fetchWithCurl(url, options);
      },
    },
    decodeURIComponent,
    encodeURIComponent,
  };
  vm.createContext(context);
  new vm.Script(code, { filename: `${product.id}/Code.gs` }).runInContext(context);

  const connection = context.testConnection();
  assert.equal(connection.success, true, `${product.id} connection`);
  const result = context[product.builder]();
  assert.equal(result.success, true, `${product.id} build`);
  assert.deepEqual(
    [...workbook.sheets.keys()].sort(),
    [...product.sheets].sort(),
    `${product.id} sheet set`,
  );
  assert.ok(
    workbook.writes.some((write) => write.formula || write.formulas),
    `${product.id} formulas`,
  );
  assert.ok(
    workbook.writes.some(
      (write) =>
        Array.isArray(write.values) &&
        write.values.flat().some((value) => Number.isFinite(value)),
    ),
    `${product.id} finite live values`,
  );
  assert.ok(
    documentValues.has("OILPRICEAPI_ACTIVATED_AT"),
    `${product.id} activation marker`,
  );

  console.log(
    `${product.id}: connection=${connection.code}, sheets=${workbook.sheets.size}, requests=${requests.length}`,
  );
}

console.log(`Production portfolio smoke passed for ${products.length} products.`);
