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
  "OAUTH_VERIFICATION.md",
  "docs/index.html",
  "package.json",
  "test/README.md",
];

test("public surfaces identify Marketplace status and canonical facts", () => {
  const text = PUBLIC_FILES.map((file) =>
    fs.readFileSync(path.join(ROOT, file), "utf8"),
  ).join("\n");
  const readme = fs.readFileSync(path.join(ROOT, "README.md"), "utf8");
  const runtime = fs.readFileSync(path.join(ROOT, "Code.gs"), "utf8");
  const docs = fs.readFileSync(path.join(ROOT, "docs/index.html"), "utf8");
  assert.match(readme, /publicly available in Google Workspace Marketplace/i);
  assert.match(
    text,
    /https:\/\/workspace\.google\.com\/marketplace\/app\/oilpriceapi_for_google_sheets\/991152473434/,
  );
  assert.doesNotMatch(runtime, /Marketplace publication is pending/i);
  assert.doesNotMatch(text, /Marketplace submission steps remain/i);
  assert.doesNotMatch(docs, /MARKETPLACE REJECTED|remediation in progress/i);
  assert.match(runtime, /listing runtime is managed separately during staged releases/i);
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

test("operator records preserve the exact release and Google submission state", () => {
  const deployment = fs.readFileSync(
    path.join(ROOT, "DEPLOYMENT_GUIDE.md"),
    "utf8",
  );
  const listing = fs.readFileSync(
    path.join(ROOT, "MARKETPLACE_LISTING.md"),
    "utf8",
  );
  const oauth = fs.readFileSync(
    path.join(ROOT, "OAUTH_VERIFICATION.md"),
    "utf8",
  );
  const packageVersion = JSON.parse(
    fs.readFileSync(path.join(ROOT, "package.json"), "utf8"),
  ).version;
  const runtime = fs.readFileSync(path.join(ROOT, "Code.gs"), "utf8");
  const records = `${deployment}\n${listing}\n${oauth}`;

  assert.equal(packageVersion, "1.3.1");
  assert.match(runtime, /ADDON_VERSION = '1\.3\.1'/);
  assert.match(deployment, /Public runtime version: `1\.2\.2`/);
  assert.match(deployment, /Repository release candidate: `1\.3\.1`/);
  assert.match(
    deployment,
    /Latest immutable Apps Script version: `12` \(runtime `1\.3\.0`/,
  );
  assert.match(
    listing,
    /Public Marketplace Apps Script version: `11` \(runtime `1\.2\.2`\)/,
  );
  for (const record of [deployment, listing, oauth]) {
    assert.match(record, /version 12[\s\S]{0,240}never published/i);
  }
  assert.match(records, /publicly available/i);
  assert.match(
    records,
    /https:\/\/workspace\.google\.com\/marketplace\/app\/oilpriceapi_for_google_sheets\/991152473434/,
  );
  assert.doesNotMatch(records, /OAuth verification has \*\*not been submitted\*\*/);
  assert.doesNotMatch(records, /OAuth submission state: \*\*not submitted\*\*/);
  assert.doesNotMatch(records, /awaiting Marketplace publication/i);
  assert.doesNotMatch(
    records,
    /Current immutable Apps Script version: `9`/,
  );
  // Preserve the earlier correction: App Configuration did not lock while the
  // Store Listing was under review.
  assert.doesNotMatch(
    records,
    /Marketplace App Configuration to Apps Script version 10/,
  );
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

test("legacy credential migration requires an explicit spreadsheet reconfigure", () => {
  const readme = fs.readFileSync(path.join(ROOT, "README.md"), "utf8");
  assert.match(readme, /unscoped keys[\s\S]{0,260}save the key again/i);
});

test("validation and production release workflows block on dependency audit", () => {
  for (const workflow of ["validate.yml", "apps-script-release.yml"]) {
    const contents = fs.readFileSync(
      path.join(ROOT, ".github", "workflows", workflow),
      "utf8",
    );
    assert.match(contents, /npm audit --audit-level=moderate/);
  }

  const results = fs.readFileSync(path.join(ROOT, "TEST_RESULTS.md"), "utf8");
  assert.match(results, /npm audit --audit-level=moderate/);
  assert.match(results, /0 vulnerabilities/);
  assert.doesNotMatch(results, /reports moderate transitive[\s\S]{0,120}advisories/i);
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
