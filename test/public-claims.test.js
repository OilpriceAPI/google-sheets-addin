const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const test = require("node:test");
const { parse } = require("yaml");

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
  "GOOGLE_MARKETPLACE_PLAYBOOK.md",
  "YOUTUBE_PROMOTION.md",
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

test("hosted workflows use hardened Node 24 release gates", () => {
  for (const workflow of [
    "validate.yml",
    "github-pages.yml",
    "apps-script-release.yml",
  ]) {
    const document = parse(
      fs.readFileSync(
        path.join(ROOT, ".github", "workflows", workflow),
        "utf8",
      ),
    );
    const steps = Object.values(document.jobs).flatMap((job) => job.steps);
    const gateJobName = workflow === "apps-script-release.yml"
      ? "release"
      : "validate";
    const gateSteps = document.jobs[gateJobName].steps;
    const checkouts = steps.filter((step) =>
      step.uses?.startsWith("actions/checkout@"),
    );
    const setupNode = steps.filter((step) =>
      step.uses?.startsWith("actions/setup-node@"),
    );

    assert.equal(document.permissions.contents, "read");
    assert.ok(checkouts.length > 0);
    for (const checkout of checkouts) {
      assert.equal(checkout.uses, "actions/checkout@v6");
      assert.equal(checkout.with?.["persist-credentials"], false);
    }
    assert.ok(setupNode.length > 0);
    for (const setup of setupNode) {
      assert.equal(setup.uses, "actions/setup-node@v6");
      assert.equal(setup.with?.["node-version"], "24");
    }
    const auditIndex = gateSteps.findIndex(
      (step) => step.run === "npm audit --audit-level=moderate",
    );
    const validationIndex = gateSteps.findIndex(
      (step) => step.run === "npm run validate",
    );
    assert.ok(auditIndex >= 0);
    assert.ok(validationIndex > auditIndex);

    if (workflow === "github-pages.yml") {
      assert.equal(document.jobs.deploy.needs, "validate");
      const pageActions = document.jobs.deploy.steps
        .map((step) => step.uses)
        .filter((uses) =>
          /actions\/(?:configure-pages|upload-pages-artifact|deploy-pages)@/.test(
            uses || "",
          ),
        )
        .sort();
      assert.deepEqual(pageActions, [
        "actions/configure-pages@v6",
        "actions/deploy-pages@v5",
        "actions/upload-pages-artifact@v5",
      ]);
    }
  }

  const results = fs.readFileSync(path.join(ROOT, "TEST_RESULTS.md"), "utf8");
  assert.match(results, /npm audit --audit-level=moderate/);
  assert.match(results, /0 vulnerabilities/);
  assert.doesNotMatch(results, /reports moderate transitive[\s\S]{0,120}advisories/i);
});

test("YouTube attribution separates every video and evaluation checkpoint", () => {
  const promotion = fs.readFileSync(
    path.join(ROOT, "YOUTUBE_PROMOTION.md"),
    "utf8",
  );
  assert.match(promotion, /utm_content=<video-id>_overview_demo/);
  assert.match(promotion, /utm_content=<video-id>_signup_cta/);
  assert.match(promotion, /Day 30[^\n]*interim/i);
  assert.match(promotion, /Day 90[^\n]*final/i);
  assert.doesNotMatch(promotion, /Marketplace publication is pending/i);
  assert.match(promotion, /publicly available in Google Workspace Marketplace/i);
});

test("portfolio readiness does not claim stale immutable packages", () => {
  const readiness = fs.readFileSync(
    path.join(ROOT, "PORTFOLIO_SUBMISSION_READINESS.md"),
    "utf8",
  );
  assert.match(readiness, /new immutable version required/i);
  assert.doesNotMatch(
    readiness,
    /every recorded immutable Apps Script version[\s\S]{0,120}matched/i,
  );
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
