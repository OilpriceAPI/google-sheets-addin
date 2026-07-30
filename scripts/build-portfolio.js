#!/usr/bin/env node

const fs = require("node:fs");
const path = require("node:path");

const ROOT = path.join(__dirname, "..");
const PORTFOLIO = path.join(ROOT, "portfolio");
const DIST = path.join(PORTFOLIO, "dist");
const products = JSON.parse(
  fs.readFileSync(path.join(PORTFOLIO, "products.json"), "utf8"),
);
const releases = JSON.parse(
  fs.readFileSync(path.join(PORTFOLIO, "releases.json"), "utf8"),
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

Status: pre-submission package. Do not claim Marketplace availability until Google approves and publishes this distinct listing.

## App details

- Application name: \`${product.name}\`
- OAuth application name: \`${product.name}\`
- Category: ${product.category}
- Pricing: Free of charge with paid features
- Developer: \`OilPriceAPI\`
- Version: \`${product.version}\`
- Product guide: \`https://www.oilpriceapi.com${product.landingPath}\`
- Pricing details: \`https://www.oilpriceapi.com/pricing\`

Short description:

> ${product.tagline}

Detailed description:

> ${product.name} creates a purpose-built workbook inside Google Sheets™. ${product.workflow}
>
> The add-on requests market data only after the user configures an OilPriceAPI key and chooses the build action. Dataset access and freshness depend on the configured account, source, and entitlement. Values retain their source timestamp where provided.
>
> The API key is stored in Apps Script document properties for the current spreadsheet. It is not written to cells, URLs, diagnostics, or browser-side HTML. Requests send only reviewed market identifiers and a product/version header used for first-party activation and reliability measurement. Spreadsheet contents, formulas, and cell values are not sent for analytics.
>
> Google Sheets™ is a trademark of Google LLC. ${product.name} is not affiliated with or endorsed by Google LLC.

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

Combined Data Access justification:

> This Editor add-on uses spreadsheets.currentonly only to create and format the named workbook tabs in the spreadsheet where the user explicitly runs ${product.menu}; it cannot browse or modify other spreadsheets. It uses script.external_request only to send the user-configured OilPriceAPI key and the product's reviewed market identifiers to api.oilpriceapi.com after the user selects Test connection or Build. It uses script.container.ui only to add the ${product.menu} menu and display its key-management sidebar, About dialog, build status, and recovery messages. No narrower scopes support these visible features. The add-on does not read Google account identity, browse Drive, send spreadsheet contents for analytics, or place API keys in cells or URLs.

## Support links

- Product guide: \`https://www.oilpriceapi.com${product.landingPath}\`
- Signup: \`https://www.oilpriceapi.com/auth/signup?utm_source=workspace_marketplace&utm_medium=addon&utm_campaign=${product.signupCampaign}\`
- Pricing: \`https://www.oilpriceapi.com/pricing\`
- Privacy: \`https://www.oilpriceapi.com/privacy/workspace-addons\`
- Terms: \`https://www.oilpriceapi.com/terms/workspace-addons\`
- Support: \`https://www.oilpriceapi.com/support\`
- Setup: \`https://www.oilpriceapi.com${product.landingPath}\`
- Help: \`https://www.oilpriceapi.com${product.landingPath}\`
- Report issue: \`https://www.oilpriceapi.com/support\`

## Submission assets

- 32px icon: \`assets/marketplace/${product.id}/app-icon-32.png\`
- 128px icon: \`assets/marketplace/${product.id}/app-icon-128.png\`
- Card banner: \`assets/marketplace/${product.id}/card-banner-220x140.png\`
- Screenshot: capture the exact immutable installed build at 1280×800 after the real-account smoke test; do not reuse another product's screenshot.
- OAuth demo video: record the exact OAuth consent screen, requested scopes, key configuration, connection test, and workbook build for this product.
`;
}

function reviewerGuide(product, release) {
  return `# ${product.name} — reviewer guide

Status: pre-submission. Evidence that depends on a real installed Marketplace draft is explicitly gated below.

## Reviewer prerequisites

- Google Workspace host: Google Sheets
- Access model: the reviewer installs the Editor add-on and supplies an OilPriceAPI test key provided privately in the Marketplace review instructions.
- No Google account identity match is required. The OilPriceAPI key may belong to a different email address.
- The reviewer fixture must be synthetic, non-customer, active for the review window, and entitled to: ${product.allowedCodes.length ? product.allowedCodes.map((code) => `\`${code}\``).join(", ") : "`ice-wti` and `ice-brent` futures curves"}.

## End-to-end review flow

1. Install the unpublished Marketplace draft in a blank spreadsheet.
2. In **Extensions → Add-ons → Manage add-ons**, select ${product.name} and choose **Use in this document**.
3. Refresh the spreadsheet once.
4. Open **Extensions → ${product.menu} → Configure OilPriceAPI key**.
5. Paste the private reviewer key and select **Save key**.
6. Select **Test connection** and confirm the green success result.
7. Select **Build workbook**.
8. Confirm these product-specific tabs exist: ${product.sheets.map((sheet) => `\`${sheet}\``).join(", ")}.
9. Confirm source timestamps and units are visible and no API key is written to a cell.
10. Delete the stored key and confirm the sidebar reports that no key is configured.

## OAuth scope demonstration

- \`spreadsheets.currentonly\`: steps 7–9 create and format only the current spreadsheet.
- \`script.external_request\`: steps 6–7 request only the reviewed OilPriceAPI market data.
- \`script.container.ui\`: steps 2–7 display the add-on menu, sidebar, status, and recovery UI.

## Evidence to attach

- Apps Script ID: \`${release.scriptId || "Not created"}\`
- Immutable Apps Script version: \`${release.version || "Not created"}\`
- Marketplace draft install: Pending the separate Cloud project and draft integration.
- Clean test spreadsheet URL: Provide privately after installed-draft smoke.
- OAuth demo video URL: Record after the exact OAuth branding and scopes are configured.
- Reviewed 1280×800 screenshot: Capture after installed-draft smoke.
- Reviewer test credential: Provide privately; never commit.

## Expected recovery behavior

- Missing key: asks the reviewer to configure a key or create an account.
- Invalid or revoked key: asks the reviewer to replace it in the sidebar.
- Dataset unavailable: explains the entitlement problem and points to pricing.
- Rate or quota limit: asks the reviewer to retry later or review the account limit.
- Timeout/network failure: asks the reviewer to check the connection and retry.
- Malformed or incomplete success response: rejects the data instead of building a misleading workbook.
`;
}

function submissionChecklist(product, release) {
  const scriptProjectReady = Boolean(release.scriptId);
  const immutableVersionReady = Number.isInteger(release.version);
  return `# ${product.name} — pre-submission checklist

Target Cloud project ID: \`${product.cloudProjectId}\`

## Automated package — complete before deployment

- [x] Distinct product title without a Google trademark
- [x] Product version \`${product.version}\`
- [x] Three least-privilege Apps Script scopes
- [x] First-party API fetch allowlist
- [x] Spreadsheet-scoped key storage with installed-custom-function compatibility
- [x] Product-specific activation header and UTM campaign
- [x] Trademark attribution and non-affiliation wording
- [x] Product-specific listing copy, reviewer guide, and graphic assets
- [x] Pricing, privacy, terms, setup, help, and support links
- [x] Unit, negative-path, claims, package, asset, and secret validation

## Google account work — prepare, but do not submit yet

- [${scriptProjectReady ? "x" : " "}] Create or confirm the standalone Apps Script project
- [${immutableVersionReady ? "x" : " "}] Push the exact validated package and create an immutable version
- [ ] Create the separate standard Google Cloud project \`${product.cloudProjectId}\`
- [ ] Enable Apps Script API and Google Workspace Marketplace SDK
- [ ] Link the Apps Script project to the standard Cloud project
- [ ] Configure External OAuth branding with the exact application name
- [ ] Add only the three manifest scopes and reconcile the Marketplace integration scopes
- [ ] Keep OAuth in Testing while preparing the installed draft
- [ ] Configure a Public Marketplace draft; remember visibility cannot be changed after saving
- [ ] Install the Marketplace draft, select **Use in this document**, refresh, and run the reviewer guide
- [ ] Inspect Apps Script executions for new exceptions, retries, or unexpected 4xx/5xx responses
- [ ] Capture a unique full-bleed 1280×800 screenshot from the immutable installed build
- [ ] Record the OAuth demonstration video for this exact app name and scope set
- [ ] Move OAuth to In production and prepare verification only when this product reaches its rollout slot

## Hold point

- [ ] Confirm the original OilPriceAPI listing has cleared OAuth and Marketplace review
- [ ] Apply any reviewer feedback to the shared runtime and every pending package
- [ ] Rebuild, create a new immutable version if anything changed, and repeat the real-account smoke
- [ ] Submit this listing only in the approved rollout order
`;
}

for (const product of products) {
  const release = releases.products[product.id];
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
  fs.writeFileSync(
    path.join(destination, "REVIEWER_GUIDE.md"),
    reviewerGuide(product, release),
  );
  fs.writeFileSync(
    path.join(destination, "SUBMISSION_CHECKLIST.md"),
    submissionChecklist(product, release),
  );
}

console.log(`Built ${products.length} distinct Apps Script release packages.`);
