#!/usr/bin/env node

const fs = require("node:fs");
const path = require("node:path");

const root = path.join(__dirname, "..");
const scriptId = process.argv[2] || process.env.APPS_SCRIPT_ID;

if (!scriptId) {
  throw new Error(
    "Pass the Apps Script ID: npm run clasp:configure -- YOUR_SCRIPT_ID",
  );
}
if (!/^[A-Za-z0-9_-]{20,}$/.test(scriptId)) {
  throw new Error("The Apps Script ID has an unexpected format.");
}

const target = path.join(root, ".clasp.json");
fs.writeFileSync(
  target,
  `${JSON.stringify({ scriptId, rootDir: "." }, null, 2)}\n`,
  { mode: 0o600 },
);

console.log(
  "Configured .clasp.json locally. It is gitignored and must not be committed.",
);
