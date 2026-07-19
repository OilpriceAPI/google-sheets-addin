#!/usr/bin/env node

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");

function read(relativePath) {
  return fs.readFileSync(path.join(ROOT, relativePath), "utf8");
}

function validateFiles() {
  const required = [
    ".github/workflows/github-pages.yml",
    ".github/workflows/validate.yml",
    "Code.gs",
    "DEPLOYMENT_GUIDE.md",
    "FetchDialog.html",
    "LICENSE",
    "README.md",
    "Sidebar.html",
    "appsscript.json",
    "docs/index.html",
    "package.json",
    "scripts/scan-secrets.sh",
    "test/public-claims.test.js",
    "test/runtime.test.js",
  ];
  for (const file of required) {
    assert.equal(fs.existsSync(path.join(ROOT, file)), true, `missing ${file}`);
  }
}

function validateCode() {
  const code = read("Code.gs");
  new vm.Script(code, { filename: "Code.gs" });

  const requiredFunctions = [
    "onOpen",
    "showSidebar",
    "saveApiKey",
    "deleteApiKey",
    "getApiKeyStatus",
    "testConnection",
    "fetchLatestPrices",
    "OILPRICE",
    "OILPRICE_HISTORY",
    "OILPRICE_CONVERT",
    "BUNKER_PRICE",
    "BUNKER_PORT_PRICES",
    "FUTURES_PRICE",
    "FUTURES_CURVE",
    "RIG_COUNT",
  ];
  for (const name of requiredFunctions) {
    assert.match(code, new RegExp(`function ${name}\\(`), `missing ${name}`);
  }

  assert.match(code, /function getApiKey_\(\)/);
  assert.doesNotMatch(code, /function getApiKey\(\)/);
  assert.match(code, /function requestJson_\(/);
  assert.match(code, /statusCode === 401/);
  assert.match(code, /statusCode === 402 \|\| statusCode === 403/);
  assert.match(code, /statusCode === 429/);
  assert.match(code, /function getCachedValue_\(/);
  assert.match(code, /sourceTimestamp_/);
  assert.match(code, /MAX_BATCH_CODES = 25/);
}

function validateUi() {
  const sidebar = read("Sidebar.html");
  const dialog = read("FetchDialog.html");

  assert.match(sidebar, /\.getApiKeyStatus\(\)/);
  assert.doesNotMatch(sidebar, /\.getApiKey\(\)/);
  assert.match(sidebar, /\.deleteApiKey\(\)/);
  assert.match(sidebar, /type="password"/);
  assert.match(sidebar, /not currently published in the Google Workspace Marketplace/i);
  assert.match(dialog, /fetchLatestPrices\(codes\)/);
  assert.match(dialog, /google\.script\.host\.close/);

  const codes = [...dialog.matchAll(/type="checkbox" value="([A-Z0-9_]+)"/g)].map(
    (match) => match[1],
  );
  assert.ok(codes.length > 0, "dialog needs commodity examples");
  assert.ok(codes.length <= 25, "dialog exceeds batch limit");
  assert.equal(new Set(codes).size, codes.length, "dialog has duplicate codes");
}

function validateManifest() {
  const manifest = JSON.parse(read("appsscript.json"));
  assert.equal(manifest.runtimeVersion, "V8");
  assert.deepEqual(manifest.oauthScopes, [
    "https://www.googleapis.com/auth/spreadsheets.currentonly",
    "https://www.googleapis.com/auth/script.external_request",
  ]);
}

function main() {
  validateFiles();
  validateCode();
  validateUi();
  validateManifest();
  console.log("Apps Script structure and bindings: valid");
}

main();
