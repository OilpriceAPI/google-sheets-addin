const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");

const ROOT = path.join(__dirname, "..");
const PUBLIC_FILES = [
  "README.md",
  "DEPLOYMENT_GUIDE.md",
  "TEST_RESULTS.md",
  "Code.gs",
  "Sidebar.html",
  "FetchDialog.html",
  "MARKETPLACE_LISTING.md",
  "docs/index.html",
  "package.json",
  "test/README.md",
];

test("public surfaces identify Marketplace status and canonical facts", () => {
  const text = PUBLIC_FILES.map((file) =>
    fs.readFileSync(path.join(ROOT, file), "utf8"),
  ).join("\n");
  assert.match(text, /Google Workspace Marketplace publication (?:is )?pending/i);
  assert.match(text, /submitted (?:on )?July 26, 2026/i);
  assert.match(text, /rejected (?:on )?July 27/i);
  assert.match(text, /trademark attribution and OAuth verification remediation/i);
  assert.doesNotMatch(text, /Marketplace submission steps remain/i);
  assert.match(text, /https:\/\/api\.oilpriceapi\.com\/product-facts\.json/);
});

test("Marketplace listing gives Google Sheets trademark attribution", () => {
  const listing = fs.readFileSync(
    path.join(ROOT, "MARKETPLACE_LISTING.md"),
    "utf8",
  );
  assert.match(listing, /Application name: `OilPriceAPI for Google Sheets™`/);
  assert.match(listing, /Google Sheets™ is a trademark of Google LLC\./);
  assert.doesNotMatch(listing, /OilPriceAPI for Sheets(?!™)/);
});

test("sidebar gives an in-product privacy notice and policy links", () => {
  const sidebar = fs.readFileSync(path.join(ROOT, "Sidebar.html"), "utf8");

  assert.match(
    sidebar,
    /works only in the spreadsheet where you open it/i,
  );
  assert.match(
    sidebar,
    /sends only your API key and requested market identifiers to OilPriceAPI/i,
  );
  assert.match(
    sidebar,
    /https:\/\/www\.oilpriceapi\.com\/privacy\/google-sheets-addon/,
  );
  assert.match(
    sidebar,
    /https:\/\/www\.oilpriceapi\.com\/terms\/google-sheets-addon/,
  );
  assert.match(sidebar, /Google API Services User Data Policy/i);
  assert.match(sidebar, /Limited Use requirements/i);
});

test("public surfaces contain no unsupported mutable claims", () => {
  const text = PUBLIC_FILES.map((file) =>
    fs.readFileSync(path.join(ROOT, file), "utf8"),
  ).join("\n");
  const unsupported = [
    /\breal[ -]?time\b/i,
    /\b(?:100|200|1,?000|10,?000|50,?000) requests?\b/i,
    /\b20\+? years?\b/i,
    /\b99\.9%\b/i,
    /\b98% less\b/i,
    /updates? every \d+ minutes?/i,
  ];
  for (const pattern of unsupported) assert.doesNotMatch(text, pattern);
});
