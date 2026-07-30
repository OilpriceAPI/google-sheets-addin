#!/usr/bin/env node

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const sharp = require("sharp");

const root = path.join(__dirname, "..", "assets", "marketplace");
const repositoryRoot = path.join(root, "..", "..");
const products = JSON.parse(
  fs.readFileSync(
    path.join(repositoryRoot, "portfolio", "products.json"),
    "utf8",
  ),
);
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

  for (const product of products) {
    const productDir = path.join(root, product.id);
    for (const [file, width, height] of expected) {
      const assetPath = path.join(productDir, file);
      assert.equal(fs.existsSync(assetPath), true, `${product.id}/${file} is required`);
      const metadata = await sharp(assetPath).metadata();
      assert.equal(metadata.format, "png", `${product.id}/${file} must be PNG`);
      assert.equal(metadata.width, width, `${product.id}/${file} width`);
      assert.equal(metadata.height, height, `${product.id}/${file} height`);
    }

    const iconSourceText = fs.readFileSync(
      path.join(productDir, "app-icon.svg"),
      "utf8",
    );
    const bannerSourceText = fs.readFileSync(
      path.join(productDir, "card-banner.svg"),
      "utf8",
    );
    assert.match(iconSourceText, new RegExp(product.iconMark));
    assert.match(iconSourceText, new RegExp(product.brandColor, "i"));
    assert.match(bannerSourceText, new RegExp(product.brandColor, "i"));

    const productScreenshots = fs
      .readdirSync(path.join(productDir, "screenshots"))
      .filter((file) => file.toLowerCase().endsWith(".png"));
    for (const file of productScreenshots) {
      const metadata = await sharp(
        path.join(productDir, "screenshots", file),
      ).metadata();
      assert.equal(metadata.format, "png", `${product.id}/${file} must be PNG`);
      assert.equal(metadata.width, 1280, `${product.id}/${file} width`);
      assert.equal(metadata.height, 800, `${product.id}/${file} height`);
    }
  }

  console.log(
    `Original Marketplace assets, ${screenshots.length} original screenshot(s), and ${products.length} distinct pre-submission asset sets are valid.`,
  );
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
