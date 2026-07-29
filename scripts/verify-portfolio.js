#!/usr/bin/env node

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const ROOT = path.join(__dirname, "..");
const PORTFOLIO = path.join(ROOT, "portfolio");
const products = JSON.parse(
  fs.readFileSync(path.join(PORTFOLIO, "products.json"), "utf8"),
);
const releases = JSON.parse(
  fs.readFileSync(path.join(PORTFOLIO, "releases.json"), "utf8"),
);
const expectedScopes = [
  "https://www.googleapis.com/auth/spreadsheets.currentonly",
  "https://www.googleapis.com/auth/script.external_request",
  "https://www.googleapis.com/auth/script.container.ui",
];

assert.equal(products.length, 5, "portfolio must contain the five approved products");
for (const field of ["id", "name", "builder", "landingPath", "signupCampaign", "activationHeader", "workflow"]) {
  assert.equal(
    new Set(products.map((product) => product[field])).size,
    products.length,
    `${field} must be unique`,
  );
}

for (const product of products) {
  assert.ok(releases.products[product.id], `${product.id} needs a release record`);
  assert.doesNotMatch(product.name, /\bGoogle\b|\bSheets\b/i, `${product.id} title must not use a Google trademark`);
  assert.ok(product.sheets.length >= 3, `${product.id} must create at least three distinct sheets`);
  const dist = path.join(PORTFOLIO, "dist", product.id);
  const code = fs.readFileSync(path.join(dist, "Code.gs"), "utf8");
  const listing = fs.readFileSync(
    path.join(dist, "MARKETPLACE_LISTING.md"),
    "utf8",
  );
  const manifest = JSON.parse(
    fs.readFileSync(path.join(dist, "appsscript.json"), "utf8"),
  );
  new vm.Script(code, { filename: `${product.id}/Code.gs` });
  assert.match(code, new RegExp(`function ${product.builder}\\(`));
  assert.match(code, new RegExp(`X-OilPriceAPI-Client`));
  assert.match(code, new RegExp(product.activationHeader));
  assert.doesNotMatch(code, /getActiveUser|getEmail|userinfo\.email|userinfo\.profile/);
  assert.deepEqual(manifest.oauthScopes, expectedScopes);
  assert.deepEqual(manifest.urlFetchWhitelist, ["https://api.oilpriceapi.com/"]);
  assert.match(listing, new RegExp(`Application name: \\\`${product.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\\``));
  assert.match(listing, /Google Sheets™ is a trademark of Google LLC\./);
  assert.match(listing, new RegExp(product.signupCampaign));
  assert.match(listing, new RegExp(product.activationHeader));
  assert.doesNotMatch(listing, /\breal[ -]?time\b/i);
}

console.log(`Portfolio verified: ${products.map((product) => product.id).join(", ")}`);
