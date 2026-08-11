#!/usr/bin/env node

const fs = require("node:fs");
const path = require("node:path");

const PATTERNS = [
  // Production API keys are bare 64-character hex values. Keep this exact
  // detector; known checksum-bearing files are allowlisted by path below.
  /[0-9a-fA-F]{64}/,
  /sk-[A-Za-z0-9_-]{20,}/,
  /AIza[A-Za-z0-9_-]{30,}/,
  /gh[pousr]_[A-Za-z0-9_]{20,}/,
  /(api[_-]?key|access[_-]?token|secret)[^\r\n]{0,40}[=:]\s*[A-Za-z0-9_-]{32,}/i,
];

const EXCLUDED_DIRECTORIES = new Set([".git", "node_modules"]);
const EXCLUDED_FILES = new Set([
  "package-lock.json",
  "scripts/scan-secrets.js",
  "scripts/scan-secrets.sh",
  "test/secret-scan.test.js",
]);

function normalizedRelativePath(root, filePath) {
  return path.relative(root, filePath).split(path.sep).join("/");
}

function isLikelyBinary(buffer) {
  return buffer.subarray(0, Math.min(buffer.length, 8192)).includes(0);
}

function findPotentialSecretFiles(root) {
  const resolvedRoot = path.resolve(root);
  const findings = [];

  function visit(directory) {
    const entries = fs.readdirSync(directory, { withFileTypes: true });
    for (const entry of entries) {
      if (entry.isDirectory() && EXCLUDED_DIRECTORIES.has(entry.name)) continue;
      const filePath = path.join(directory, entry.name);
      if (entry.isDirectory()) {
        visit(filePath);
        continue;
      }
      if (!entry.isFile()) continue;

      const relativePath = normalizedRelativePath(resolvedRoot, filePath);
      if (EXCLUDED_FILES.has(relativePath)) continue;
      const contents = fs.readFileSync(filePath);
      if (isLikelyBinary(contents)) continue;
      const text = contents.toString("utf8");
      if (PATTERNS.some((pattern) => pattern.test(text))) findings.push(relativePath);
    }
  }

  visit(resolvedRoot);
  return findings.sort();
}

function formatFindings(findings) {
  if (findings.length === 0) return "Secret scan passed (filenames only).\n";
  return `Potential secret pattern found in:\n${findings.join("\n")}\n`;
}

function main() {
  const findings = findPotentialSecretFiles(process.argv[2] || process.cwd());
  process.stdout.write(formatFindings(findings));
  if (findings.length > 0) process.exitCode = 1;
}

if (require.main === module) main();

module.exports = { findPotentialSecretFiles, formatFindings };
