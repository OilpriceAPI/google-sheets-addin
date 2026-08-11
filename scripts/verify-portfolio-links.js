#!/usr/bin/env node

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.join(__dirname, "..");
const products = JSON.parse(
  fs.readFileSync(path.join(ROOT, "portfolio", "products.json"), "utf8"),
);
const base = "https://www.oilpriceapi.com";
const urls = [
  ...products.map((product) => `${base}${product.landingPath}`),
  `${base}/auth/signup`,
  `${base}/pricing`,
  `${base}/privacy/workspace-addons`,
  `${base}/terms/workspace-addons`,
  `${base}/support`,
];

async function main() {
  for (const url of urls) {
    const response = await fetch(url, {
      headers: { "User-Agent": "OilPriceAPI-Workspace-release-verifier/1.0" },
      redirect: "follow",
    });
    assert.equal(
      response.ok,
      true,
      `${url} returned HTTP ${response.status}`,
    );
    assert.equal(
      response.url.startsWith(base),
      true,
      `${url} redirected outside the approved first-party domain`,
    );
    console.log(`${response.status} ${url}`);
  }

  console.log(`${urls.length} public submission links are reachable.`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
