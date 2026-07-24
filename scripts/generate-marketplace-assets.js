#!/usr/bin/env node

const fs = require("node:fs");
const path = require("node:path");
const sharp = require("sharp");

const root = path.join(__dirname, "..");
const assetDir = path.join(root, "assets", "marketplace");
const iconSource = path.join(assetDir, "app-icon.svg");
const bannerSource = path.join(assetDir, "card-banner.svg");

async function main() {
  fs.mkdirSync(assetDir, { recursive: true });
  await sharp(iconSource).resize(32, 32).png().toFile(path.join(assetDir, "app-icon-32.png"));
  await sharp(iconSource).resize(128, 128).png().toFile(path.join(assetDir, "app-icon-128.png"));
  await sharp(bannerSource).resize(220, 140).png().toFile(path.join(assetDir, "card-banner-220x140.png"));
  console.log("Generated Marketplace icons and card banner.");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
