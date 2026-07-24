#!/usr/bin/env node

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const root = path.join(__dirname, "..");
const runtimeFiles = [
  "Code.gs",
  "Sidebar.html",
  "FetchDialog.html",
  "appsscript.json",
];
const expectedIgnore = [
  "**",
  "!Code.gs",
  "!Sidebar.html",
  "!FetchDialog.html",
  "!appsscript.json",
];

for (const file of runtimeFiles) {
  assert.equal(fs.existsSync(path.join(root, file)), true, `missing ${file}`);
}

const actualIgnore = fs
  .readFileSync(path.join(root, ".claspignore"), "utf8")
  .split(/\r?\n/)
  .map((line) => line.trim())
  .filter(Boolean);
assert.deepEqual(
  actualIgnore,
  expectedIgnore,
  ".claspignore must expose only reviewed runtime files",
);

const manifest = JSON.parse(
  fs.readFileSync(path.join(root, "appsscript.json"), "utf8"),
);
assert.deepEqual(manifest.oauthScopes, [
  "https://www.googleapis.com/auth/spreadsheets.currentonly",
  "https://www.googleapis.com/auth/script.external_request",
]);
assert.deepEqual(manifest.urlFetchWhitelist, [
  "https://api.oilpriceapi.com/",
]);

console.log(`Deployment package verified: ${runtimeFiles.join(", ")}`);
