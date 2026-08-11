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
for (const field of ["id", "name", "builder", "landingPath", "signupCampaign", "activationHeader", "workflow", "cloudProjectId", "iconMark", "brandColor"]) {
  assert.equal(
    new Set(products.map((product) => product[field])).size,
    products.length,
    `${field} must be unique`,
  );
}

for (const product of products) {
  const release = releases.products[product.id];
  assert.ok(release, `${product.id} needs a release record`);
  assert.equal(product.version, "1.0.0", `${product.id} candidate version`);
  assert.ok(product.name.length <= 50, `${product.id} listing name is limited to 50 characters`);
  assert.ok(product.tagline.length <= 200, `${product.id} short description is limited to 200 characters`);
  assert.match(product.cloudProjectId, /^[a-z][a-z0-9-]{4,28}[a-z0-9]$/);
  assert.match(product.brandColor, /^#[0-9A-F]{6}$/);
  assert.equal(release.scriptId.length >= 20, true, `${product.id} script ID`);
  assert.equal(release.stage, "source-candidate", `${product.id} release stage`);
  assert.equal(release.version, null, `${product.id} current immutable version`);
  assert.equal(
    Number.isInteger(release.previousImmutableVersion),
    true,
    `${product.id} historical immutable version`,
  );
  assert.equal(release.cloudProjectId, product.cloudProjectId);
  assert.doesNotMatch(product.name, /\bGoogle\b|\bSheets\b/i, `${product.id} title must not use a Google trademark`);
  assert.ok(product.sheets.length >= 3, `${product.id} must create at least three distinct sheets`);
  const dist = path.join(PORTFOLIO, "dist", product.id);
  const code = fs.readFileSync(path.join(dist, "Code.gs"), "utf8");
  const listing = fs.readFileSync(
    path.join(dist, "MARKETPLACE_LISTING.md"),
    "utf8",
  );
  const sidebar = fs.readFileSync(path.join(dist, "Sidebar.html"), "utf8");
  const manifest = JSON.parse(
    fs.readFileSync(path.join(dist, "appsscript.json"), "utf8"),
  );
  const reviewerGuide = fs.readFileSync(
    path.join(dist, "REVIEWER_GUIDE.md"),
    "utf8",
  );
  const submissionChecklist = fs.readFileSync(
    path.join(dist, "SUBMISSION_CHECKLIST.md"),
    "utf8",
  );
  new vm.Script(code, { filename: `${product.id}/Code.gs` });
  assert.match(code, new RegExp(`function ${product.builder}\\(`));
  assert.match(code, new RegExp(`X-API-Client`));
  assert.doesNotMatch(code, /X-OilPriceAPI-Client/);
  assert.match(code, new RegExp(product.activationHeader));
  assert.doesNotMatch(code, /getActiveUser|getEmail|userinfo\.email|userinfo\.profile/);
  assert.deepEqual(manifest.oauthScopes, expectedScopes);
  assert.deepEqual(manifest.urlFetchWhitelist, ["https://api.oilpriceapi.com/"]);
  assert.match(listing, new RegExp(`Application name: \\\`${product.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\\``));
  assert.match(listing, /Google Sheets™ is a trademark of Google LLC\./);
  assert.match(listing, new RegExp(product.signupCampaign));
  assert.match(listing, new RegExp(product.activationHeader));
  assert.match(listing, /Combined Data Access justification:/);
  assert.match(listing, /not affiliated with or endorsed by Google LLC\./);
  assert.match(listing, /https:\/\/www\.oilpriceapi\.com\/pricing/);
  assert.doesNotMatch(listing, /\breal[ -]?time\b/i);
  assert.match(reviewerGuide, new RegExp(release.scriptId));
  assert.match(sidebar, /Prototype testers must save the key again/i);
  assert.match(reviewerGuide, /Immutable Apps Script version: `New immutable version required`/);
  assert.ok(
    reviewerGuide.includes(
      `Previous immutable version (superseded): \`${release.previousImmutableVersion}\``,
    ),
  );
  assert.match(submissionChecklist, new RegExp(product.cloudProjectId));
  assert.doesNotMatch(
    `${reviewerGuide}\n${submissionChecklist}`,
    /\[(?:record|provide|replace|TODO)[^\]]*\]/i,
  );
}

console.log(`Portfolio verified: ${products.map((product) => product.id).join(", ")}`);
