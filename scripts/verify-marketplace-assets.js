#!/usr/bin/env node

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const sharp = require("sharp");

const root = path.join(__dirname, "..", "assets", "marketplace");
const expected = [
  ["app-icon-32.png", 32, 32],
  ["app-icon-128.png", 128, 128],
  ["card-banner-220x140.png", 220, 140],
];

async function main() {
  for (const [file, width, height] of expected) {
    const metadata = await sharp(path.join(root, file)).metadata();
    assert.equal(metadata.format, "png", `${file} must be PNG`);
    assert.equal(metadata.width, width, `${file} width`);
    assert.equal(metadata.height, height, `${file} height`);
  }

  const screenshotsDir = path.join(root, "screenshots");
  const screenshots = fs
    .readdirSync(screenshotsDir)
    .filter((file) => file.toLowerCase().endsWith(".png"));
  assert.ok(screenshots.length > 0, "at least one Marketplace screenshot is required");

  for (const file of screenshots) {
    const metadata = await sharp(path.join(screenshotsDir, file)).metadata();
    assert.equal(metadata.format, "png", `${file} must be PNG`);
    assert.equal(metadata.width, 1280, `${file} width`);
    assert.equal(metadata.height, 800, `${file} height`);
  }

  console.log(
    `Marketplace assets and ${screenshots.length} screenshot(s) have the required formats and dimensions.`,
  );
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
