# Energy Curve Builder by OilPriceAPI — Marketplace listing

Status: pre-submission package. Do not claim Marketplace availability until Google approves and publishes this distinct listing.

## App details

- Application name: `Energy Curve Builder by OilPriceAPI`
- OAuth application name: `Energy Curve Builder by OilPriceAPI`
- Category: Accounting and Finance
- Pricing: Free of charge with paid features
- Developer: `OilPriceAPI`
- Version: `1.0.0`
- Product guide: `https://www.oilpriceapi.com/integrations/energy-curve-builder`
- Pricing details: `https://www.oilpriceapi.com/pricing`

Short description:

> Create futures curves, calendar spreads, and structure signals.

Detailed description:

> Energy Curve Builder by OilPriceAPI creates a purpose-built workbook inside Google Sheets™. Creates WTI and Brent term-structure tables, month-on-month calendar spreads, and contango/backwardation signals.
>
> The add-on requests market data only after the user configures an OilPriceAPI key and chooses the build action. Dataset access and freshness depend on the configured account, source, and entitlement. Values retain their source timestamp where provided.
>
> The API key is stored in Apps Script document properties for the current spreadsheet. It is not written to cells, URLs, diagnostics, or browser-side HTML. Requests send only reviewed market identifiers and a product/version header used for first-party activation and reliability measurement. Spreadsheet contents, formulas, and cell values are not sent for analytics.
>
> Google Sheets™ is a trademark of Google LLC. Energy Curve Builder by OilPriceAPI is not affiliated with or endorsed by Google LLC.

## Distinct workflow

- Generated sheets: `Curve Dashboard`, `WTI Curve`, `Brent Curve`, `Calendar Spreads`
- Primary action: `buildEnergyCurveWorkbook`
- Differentiation: Creates WTI and Brent term-structure tables, month-on-month calendar spreads, and contango/backwardation signals.

## Measurement

- Marketplace discovery: Google Workspace Marketplace SDK impressions and install events.
- Activation: first successful OilPriceAPI request carrying `X-OilPriceAPI-Client: energy-curve-builder/<version>`.
- Signup: `utm_source=workspace_marketplace&utm_medium=addon&utm_campaign=energy_curve_builder`.
- North-star rate: activated workbooks per 100 listing views.

## OAuth scopes

| Scope | Justification |
| --- | --- |
| `https://www.googleapis.com/auth/spreadsheets.currentonly` | Create and format only the workbook where the user runs the add-on. |
| `https://www.googleapis.com/auth/script.external_request` | Request the product's reviewed market data from `api.oilpriceapi.com`. |
| `https://www.googleapis.com/auth/script.container.ui` | Display the add-on menu, key sidebar, build action, and recovery messages. |

Google's mandatory `userinfo.email` and `userinfo.profile` defaults may appear in Cloud configuration. Product behavior does not use them. No Drive-wide scope is requested.

Combined Data Access justification:

> This Editor add-on uses spreadsheets.currentonly only to create and format the named workbook tabs in the spreadsheet where the user explicitly runs Energy Curve Builder; it cannot browse or modify other spreadsheets. It uses script.external_request only to send the user-configured OilPriceAPI key and the product's reviewed market identifiers to api.oilpriceapi.com after the user selects Test connection or Build. It uses script.container.ui only to add the Energy Curve Builder menu and display its key-management sidebar, About dialog, build status, and recovery messages. No narrower scopes support these visible features. The add-on does not read Google account identity, browse Drive, send spreadsheet contents for analytics, or place API keys in cells or URLs.

## Support links

- Product guide: `https://www.oilpriceapi.com/integrations/energy-curve-builder`
- Signup: `https://www.oilpriceapi.com/auth/signup?utm_source=workspace_marketplace&utm_medium=addon&utm_campaign=energy_curve_builder`
- Pricing: `https://www.oilpriceapi.com/pricing`
- Privacy: `https://www.oilpriceapi.com/privacy/workspace-addons`
- Terms: `https://www.oilpriceapi.com/terms/workspace-addons`
- Support: `https://www.oilpriceapi.com/support`
- Setup: `https://www.oilpriceapi.com/integrations/energy-curve-builder`
- Help: `https://www.oilpriceapi.com/integrations/energy-curve-builder`
- Report issue: `https://www.oilpriceapi.com/support`

## Submission assets

- 32px icon: `assets/marketplace/energy-curve-builder/app-icon-32.png`
- 128px icon: `assets/marketplace/energy-curve-builder/app-icon-128.png`
- Card banner: `assets/marketplace/energy-curve-builder/card-banner-220x140.png`
- Screenshot: capture the exact immutable installed build at 1280×800 after the real-account smoke test; do not reuse another product's screenshot.
- OAuth demo video: record the exact OAuth consent screen, requested scopes, key configuration, connection test, and workbook build for this product.
