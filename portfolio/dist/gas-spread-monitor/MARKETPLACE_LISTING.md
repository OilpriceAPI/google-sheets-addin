# Gas Spread Monitor by OilPriceAPI — Marketplace listing

Status: pre-submission package. Do not claim Marketplace availability until Google approves and publishes this distinct listing.

## App details

- Application name: `Gas Spread Monitor by OilPriceAPI`
- OAuth application name: `Gas Spread Monitor by OilPriceAPI`
- Category: Accounting and Finance
- Pricing: Free of charge with paid features
- Developer: `OilPriceAPI`
- Version: `1.0.0`
- Product guide: `https://www.oilpriceapi.com/integrations/gas-spread-monitor`
- Pricing details: `https://www.oilpriceapi.com/pricing`

Short description:

> Normalize Henry Hub, TTF, and JKM to compare global gas spreads.

Detailed description:

> Gas Spread Monitor by OilPriceAPI creates a purpose-built workbook inside Google Sheets™. Creates a cross-market gas monitor that normalizes currencies and energy units before calculating TTF-Henry Hub and JKM-Henry Hub spreads.
>
> The add-on requests market data only after the user configures an OilPriceAPI key and chooses the build action. Dataset access and freshness depend on the configured account, source, and entitlement. Values retain their source timestamp where provided.
>
> The API key is stored in Apps Script document properties for the current spreadsheet. It is not written to cells, URLs, diagnostics, or browser-side HTML. Requests send only reviewed market identifiers and a product/version header used for first-party activation and reliability measurement. Spreadsheet contents, formulas, and cell values are not sent for analytics.
>
> Google Sheets™ is a trademark of Google LLC. Gas Spread Monitor by OilPriceAPI is not affiliated with or endorsed by Google LLC.

## Distinct workflow

- Generated sheets: `Gas Spread Monitor`, `Market Data`, `Conversion Audit`
- Primary action: `buildGasSpreadWorkbook`
- Differentiation: Creates a cross-market gas monitor that normalizes currencies and energy units before calculating TTF-Henry Hub and JKM-Henry Hub spreads.

## Measurement

- Marketplace discovery: Google Workspace Marketplace SDK impressions and install events.
- Activation: first successful OilPriceAPI request carrying `X-API-Client: gas-spread-monitor/<version>`.
- Signup: `utm_source=workspace_marketplace&utm_medium=addon&utm_campaign=gas_spread_monitor`.
- North-star rate: activated workbooks per 100 listing views.

## OAuth scopes

| Scope | Justification |
| --- | --- |
| `https://www.googleapis.com/auth/spreadsheets.currentonly` | Create and format only the workbook where the user runs the add-on. |
| `https://www.googleapis.com/auth/script.external_request` | Request the product's reviewed market data from `api.oilpriceapi.com`. |
| `https://www.googleapis.com/auth/script.container.ui` | Display the add-on menu, key sidebar, build action, and recovery messages. |

Google's mandatory `userinfo.email` and `userinfo.profile` defaults may appear in Cloud configuration. Product behavior does not use them. No Drive-wide scope is requested.

Combined Data Access justification:

> This Editor add-on uses spreadsheets.currentonly only to create and format the named workbook tabs in the spreadsheet where the user explicitly runs Gas Spread Monitor; it cannot browse or modify other spreadsheets. It uses script.external_request only to send the user-configured OilPriceAPI key and the product's reviewed market identifiers to api.oilpriceapi.com after the user selects Test connection or Build. It uses script.container.ui only to add the Gas Spread Monitor menu and display its key-management sidebar, About dialog, build status, and recovery messages. No narrower scopes support these visible features. The add-on does not read Google account identity, browse Drive, send spreadsheet contents for analytics, or place API keys in cells or URLs.

## Support links

- Product guide: `https://www.oilpriceapi.com/integrations/gas-spread-monitor`
- Signup: `https://www.oilpriceapi.com/auth/signup?utm_source=workspace_marketplace&utm_medium=addon&utm_campaign=gas_spread_monitor`
- Pricing: `https://www.oilpriceapi.com/pricing`
- Privacy: `https://www.oilpriceapi.com/privacy/workspace-addons`
- Terms: `https://www.oilpriceapi.com/terms/workspace-addons`
- Support: `https://www.oilpriceapi.com/support`
- Setup: `https://www.oilpriceapi.com/integrations/gas-spread-monitor`
- Help: `https://www.oilpriceapi.com/integrations/gas-spread-monitor`
- Report issue: `https://www.oilpriceapi.com/support`

## Submission assets

- 32px icon: `assets/marketplace/gas-spread-monitor/app-icon-32.png`
- 128px icon: `assets/marketplace/gas-spread-monitor/app-icon-128.png`
- Card banner: `assets/marketplace/gas-spread-monitor/card-banner-220x140.png`
- Screenshot: capture the exact immutable installed build at 1280×800 after the real-account smoke test; do not reuse another product's screenshot.
- OAuth demo video: record the exact OAuth consent screen, requested scopes, key configuration, connection test, and workbook build for this product.
