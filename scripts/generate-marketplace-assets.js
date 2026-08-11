#!/usr/bin/env node

const fs = require("node:fs");
const path = require("node:path");
const sharp = require("sharp");

const root = path.join(__dirname, "..");
const assetDir = path.join(root, "assets", "marketplace");
const iconSource = path.join(assetDir, "app-icon.svg");
const bannerSource = path.join(assetDir, "card-banner.svg");
const products = JSON.parse(
  fs.readFileSync(path.join(root, "portfolio", "products.json"), "utf8"),
);

function escapeXml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&apos;");
}

function productIcon(product) {
  return `<svg xmlns="http://www.w3.org/2000/svg" width="128" height="128" viewBox="0 0 128 128">
  <rect width="128" height="128" rx="24" fill="${escapeXml(product.brandColor)}"/>
  <path d="M26 91V68l15-16 15 10 20-28 26 18v39H26Z" fill="#fff" opacity=".20"/>
  <path d="M28 90h72M39 82V68m18 14V57m18 25V45m18 37V59" fill="none" stroke="#fff" stroke-width="6" stroke-linecap="round"/>
  <text x="64" y="111" fill="#fff" font-family="Arial,sans-serif" font-size="19" font-weight="700" text-anchor="middle">${escapeXml(product.iconMark)}</text>
</svg>
`;
}

function productBanner(product) {
  const words = product.menu.split(" ");
  const splitAt = Math.ceil(words.length / 2);
  const lineOne = words.slice(0, splitAt).join(" ");
  const lineTwo = words.slice(splitAt).join(" ");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="220" height="140" viewBox="0 0 220 140">
  <rect width="220" height="140" fill="#F8FAFC"/>
  <rect x="0" y="0" width="12" height="140" fill="${escapeXml(product.brandColor)}"/>
  <circle cx="54" cy="43" r="25" fill="${escapeXml(product.brandColor)}"/>
  <path d="M37 50h34M41 45l8-8 8 6 10-13" fill="none" stroke="#fff" stroke-width="4" stroke-linecap="round" stroke-linejoin="round"/>
  <text x="88" y="37" fill="#0F172A" font-family="Arial,sans-serif" font-size="17" font-weight="700">${escapeXml(lineOne)}</text>
  <text x="88" y="58" fill="#0F172A" font-family="Arial,sans-serif" font-size="17" font-weight="700">${escapeXml(lineTwo)}</text>
  <text x="30" y="104" fill="#475569" font-family="Arial,sans-serif" font-size="13">Purpose-built workflow</text>
  <text x="30" y="123" fill="${escapeXml(product.brandColor)}" font-family="Arial,sans-serif" font-size="13" font-weight="700">by OilPriceAPI</text>
</svg>
`;
}

async function main() {
  fs.mkdirSync(assetDir, { recursive: true });
  await sharp(iconSource).resize(32, 32).png().toFile(path.join(assetDir, "app-icon-32.png"));
  await sharp(iconSource).resize(128, 128).png().toFile(path.join(assetDir, "app-icon-128.png"));
  await sharp(bannerSource).resize(220, 140).png().toFile(path.join(assetDir, "card-banner-220x140.png"));

  for (const product of products) {
    const productDir = path.join(assetDir, product.id);
    const screenshotsDir = path.join(productDir, "screenshots");
    fs.mkdirSync(screenshotsDir, { recursive: true });
    const icon = productIcon(product);
    const banner = productBanner(product);
    fs.writeFileSync(path.join(productDir, "app-icon.svg"), icon);
    fs.writeFileSync(path.join(productDir, "card-banner.svg"), banner);
    fs.writeFileSync(
      path.join(screenshotsDir, "README.md"),
      `# ${product.name} screenshot evidence\n\nCapture at least one full-bleed 1280×800 PNG from the exact immutable installed Marketplace draft after completing the real-account smoke in REVIEWER_GUIDE.md. Do not reuse or simulate another product's screenshot.\n`,
    );
    await sharp(Buffer.from(icon))
      .resize(32, 32)
      .png()
      .toFile(path.join(productDir, "app-icon-32.png"));
    await sharp(Buffer.from(icon))
      .resize(128, 128)
      .png()
      .toFile(path.join(productDir, "app-icon-128.png"));
    await sharp(Buffer.from(banner))
      .resize(220, 140)
      .png()
      .toFile(path.join(productDir, "card-banner-220x140.png"));
  }

  console.log(`Generated original assets and ${products.length} distinct product asset sets.`);
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
