const assert = require("node:assert/strict");
const fs = require("node:fs");
const os = require("node:os");
const path = require("node:path");
const test = require("node:test");

const {
  findPotentialSecretFiles,
  formatFindings,
} = require("../scripts/scan-secrets.js");

test("secret scan detects credentials without printing their values", (t) => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "opa-secret-scan-"));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  fs.mkdirSync(path.join(root, "nested"));
  fs.writeFileSync(path.join(root, "safe.txt"), "API key is configured in the sidebar.\n");
  const credential = `sk-${"x".repeat(24)}`;
  fs.writeFileSync(path.join(root, "nested", "secret.txt"), `${credential}\n`);

  const findings = findPotentialSecretFiles(root);
  assert.deepEqual(findings, ["nested/secret.txt"]);
  const output = formatFindings(findings);
  assert.match(output, /nested\/secret\.txt/);
  assert.equal(output.includes(credential), false);
});

test("secret scan ignores dependency and scanner implementation paths", (t) => {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "opa-secret-scan-"));
  t.after(() => fs.rmSync(root, { recursive: true, force: true }));
  const credential = `AIza${"x".repeat(32)}`;
  for (const relativePath of [
    "node_modules/package/token.txt",
    "scripts/scan-secrets.js",
    "test/secret-scan.test.js",
    "package-lock.json",
  ]) {
    const target = path.join(root, relativePath);
    fs.mkdirSync(path.dirname(target), { recursive: true });
    fs.writeFileSync(target, credential);
  }

  assert.deepEqual(findPotentialSecretFiles(root), []);
});
