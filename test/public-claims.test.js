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
  "docs/index.html",
  "package.json",
  "test/README.md",
];

test("public surfaces identify Marketplace status and canonical facts", () => {
  const text = PUBLIC_FILES.map((file) =>
    fs.readFileSync(path.join(ROOT, file), "utf8"),
  ).join("\n");
  assert.match(text, /not (?:currently )?published in (?:the )?Google Workspace Marketplace/i);
  assert.match(text, /https:\/\/api\.oilpriceapi\.com\/product-facts\.json/);
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
