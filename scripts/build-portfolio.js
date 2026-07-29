#!/usr/bin/env node

const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.join(__dirname, "..");
const PORTFOLIO = path.join(ROOT, "portfolio");
const DIST = path.join(PORTFOLIO, "dist");
const products = JSON.parse(
  fs.readFileSync(path.join(PORTFOLIO, "products.json"), "utf8"),
);
const core = fs.readFileSync(
  path.join(PORTFOLIO, "runtime", "Core.gs"),
  "utf8",
);
const sidebar = fs.readFileSync(
  path.join(PORTFOLIO, "runtime", "Sidebar.html"),
  "utf8",
);
const scopes = [
  "https://www.googleapis.com/auth/spreadsheets.currentonly",
  "https://www.googleapis.com/auth/script.external_request",
  "https://www.googleapis.com/auth/script.container.ui",
];

function listing(product) {
  return `# ${product.name} — Marketplace listing

Status: release package validated locally. Do not claim Marketplace availability until Google approves and publishes this distinct listing.

## App details

- Application name: \`${product.name}\`
- Category: ${product.category}
- Pricing: Free of charge with paid features
- Developer: \`OilPriceAPI\`

Short description:

> ${product.tagline}

Detailed description:

> ${product.name} creates a purpose-built workbook inside Google Sheets™. ${product.workflow}
>
> The add-on requests market data only after the user configures an OilPriceAPI key and chooses the build action. Dataset access and freshness depend on the configured account, source, and entitlement. Values retain their source timestamp where provided.
>
> The API key is stored in Apps Script document properties for the current spreadsheet. It is not written to cells, URLs, diagnostics, or browser-side HTML. Requests send only reviewed market identifiers and a product/version header used for first-party activation and reliability measurement. Spreadsheet contents, formulas, and cell values are not sent for analytics.
>
> Google Sheets™ is a trademark of Google LLC.

## Distinct workflow

- Generated sheets: ${product.sheets.map((sheet) => `\`${sheet}\``).join(", ")}
- Primary action: \`${product.builder}\`
- Differentiation: ${product.workflow}

## Measurement

- Marketplace discovery: Google Workspace Marketplace SDK impressions and install events.
- Activation: first successful OilPriceAPI request carrying \`X-OilPriceAPI-Client: ${product.activationHeader}/<version>\`.
- Signup: \`utm_source=workspace_marketplace&utm_medium=addon&utm_campaign=${product.signupCampaign}\`.
- North-star rate: activated workbooks per 100 listing views.

## OAuth scopes

| Scope | Justification |
| --- | --- |
| \`${scopes[0]}\` | Create and format only the workbook where the user runs the add-on. |
| \`${scopes[1]}\` | Request the product's reviewed market data from \`api.oilpriceapi.com\`. |
| \`${scopes[2]}\` | Display the add-on menu, key sidebar, build action, and recovery messages. |

Google's mandatory \`userinfo.email\` and \`userinfo.profile\` defaults may appear in Cloud configuration. Product behavior does not use them. No Drive-wide scope is requested.

## Support links

- Product guide: \`https://www.oilpriceapi.com${product.landingPath}\`
- Signup: \`https://www.oilpriceapi.com/auth/signup?utm_source=workspace_marketplace&utm_medium=addon&utm_campaign=${product.signupCampaign}\`
- Privacy: \`https://www.oilpriceapi.com/privacy/workspace-addons\`
- Terms: \`https://www.oilpriceapi.com/terms/workspace-addons\`
- Support: \`https://www.oilpriceapi.com/support\`
`;
}

for (const product of products) {
  const source = fs.readFileSync(
    path.join(PORTFOLIO, "products", product.id, "Product.gs"),
    "utf8",
  );
  const destination = path.join(DIST, product.id);
  fs.mkdirSync(destination, { recursive: true });
  const header = `const OPA_PRODUCT = Object.freeze(${JSON.stringify(product, null, 2)});\n\n`;
  fs.writeFileSync(path.join(destination, "Code.gs"), `${header}${core}\n${source}`);
  fs.writeFileSync(path.join(destination, "Sidebar.html"), sidebar);
  fs.writeFileSync(
    path.join(destination, "appsscript.json"),
    `${JSON.stringify(
      {
        timeZone: "America/New_York",
        dependencies: {},
        exceptionLogging: "STACKDRIVER",
        runtimeVersion: "V8",
        oauthScopes: scopes,
        urlFetchWhitelist: ["https://api.oilpriceapi.com/"],
        sheets: { macros: [] },
      },
      null,
      2,
    )}\n`,
  );
  fs.writeFileSync(
    path.join(destination, ".claspignore"),
    "**\n!Code.gs\n!Sidebar.html\n!appsscript.json\n",
  );
  fs.writeFileSync(
    path.join(destination, "MARKETPLACE_LISTING.md"),
    listing(product),
  );
}

console.log(`Built ${products.length} distinct Apps Script release packages.`);
