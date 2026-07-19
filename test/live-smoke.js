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
    getUserProperties: () => ({
      getProperty: () => apiKey,
      setProperty: () => undefined,
      deleteProperty: () => undefined,
    }),
  },
  String,
  UrlFetchApp: { fetch: fetchWithCurl },
  console,
  encodeURIComponent,
};
vm.createContext(context);
new vm.Script(code, { filename: "Code.gs" }).runInContext(context);

const price = context.OILPRICE("WTI_USD");
assert.ok(Number.isFinite(price) && price > 0);
assert.equal(context.OILPRICE("WTI_USD"), price, "cached formula value changed");

const history = context.OILPRICE_HISTORY("WTI_USD", 1);
assert.ok(Array.isArray(history) && history.length > 0);
assert.ok(Number.isFinite(new Date(history[0][0]).getTime()));
assert.ok(Number.isFinite(history[0][1]));

console.log(
  `Production formula smoke passed: latest=${price}, history_records=${history.length}, source_timestamp=${history[history.length - 1][0]}`,
);
