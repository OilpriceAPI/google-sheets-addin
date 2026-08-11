#!/usr/bin/env node

const assert = require("node:assert/strict");
const { spawnSync } = require("node:child_process");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");

const ROOT = path.join(__dirname, "..");
const releases = JSON.parse(
  fs.readFileSync(path.join(ROOT, "portfolio", "releases.json"), "utf8"),
);
const clasp = path.join(ROOT, "node_modules", ".bin", "clasp");

for (const [productId, release] of Object.entries(releases.products)) {
  assert.ok(release.scriptId, `${productId} script ID`);
  assert.ok(Number.isInteger(release.version), `${productId} immutable version`);
  const temporary = fs.mkdtempSync(
    path.join(os.tmpdir(), `opa-${productId}-`),
  );
  try {
    const result = spawnSync(
      clasp,
      ["clone", release.scriptId, String(release.version)],
      { cwd: temporary, encoding: "utf8" },
    );
    assert.equal(
      result.status,
      0,
      `${productId} clone failed: ${result.stderr || result.stdout}`,
    );
    const local = path.join(ROOT, "portfolio", "dist", productId);
    const remoteCodeFile = fs.existsSync(path.join(temporary, "Code.js"))
      ? "Code.js"
      : "Code.gs";
    assert.equal(
      fs.readFileSync(path.join(temporary, remoteCodeFile), "utf8"),
      fs.readFileSync(path.join(local, "Code.gs"), "utf8"),
      `${productId} immutable Code differs from local candidate`,
    );
    assert.equal(
      fs.readFileSync(path.join(temporary, "Sidebar.html"), "utf8"),
      fs.readFileSync(path.join(local, "Sidebar.html"), "utf8"),
      `${productId} immutable Sidebar differs from local candidate`,
    );
    assert.deepEqual(
      JSON.parse(
        fs.readFileSync(path.join(temporary, "appsscript.json"), "utf8"),
      ),
      JSON.parse(fs.readFileSync(path.join(local, "appsscript.json"), "utf8")),
      `${productId} immutable manifest differs from local candidate`,
    );
    console.log(
      `${productId}: immutable Apps Script version ${release.version} matches the local package.`,
    );
  } finally {
    fs.rmSync(temporary, { recursive: true, force: true });
  }
}

console.log("All recorded Apps Script candidates match their immutable releases.");
