# OilPriceAPI Workspace portfolio

This repository contains five distinct Google Workspace Marketplace candidates.
They share a security and release runtime, but each product creates a different
finished workbook for a different buyer and workflow.

| Product | Primary buyer | Activated outcome | Acquisition keyword wedge |
| --- | --- | --- | --- |
| Crack Spread Lab by OilPriceAPI | Refinery analyst or energy trader | Live refinery-margin workbook with history and sensitivity | crack spread spreadsheet |
| Bunker Voyage Planner by OilPriceAPI | Charterer, bunker buyer, or vessel operator | Port-comparison voyage fuel budget | bunker fuel voyage calculator |
| Fuel Surcharge Studio by OilPriceAPI | Carrier, broker, or shipper | Auditable diesel-index surcharge schedule | fuel surcharge spreadsheet |
| Energy Curve Builder by OilPriceAPI | Commodity analyst or risk team | WTI/Brent curves and calendar-spread signals | oil futures curve spreadsheet |
| Gas Spread Monitor by OilPriceAPI | LNG or gas-market analyst | HH/TTF/JKM comparison normalized to USD/MMBtu | natural gas spread spreadsheet |

## Policy gates

Every release package must pass `npm run validate`.

- Product titles do not contain `Google`, `Google Sheets`, or another Google
  product trademark.
- Descriptive listing copy may say Google Sheets™ and includes
  `Google Sheets™ is a trademark of Google LLC.`
- Each package declares only `spreadsheets.currentonly`,
  `script.external_request`, and `script.container.ui`.
- Each package uses its own Cloud project, OAuth identity, Apps Script project,
  immutable version, Marketplace listing, landing page, and screenshot/video
  evidence.
- No listing may be submitted until the exact deployed build passes a real
  Editor add-on smoke test and its OAuth configuration is In production.
- The five listings must not reuse screenshots or detailed descriptions.
  Shared infrastructure is acceptable; duplicate user experiences are not.

## Measurement contract

The north-star metric is **activated workbooks per 100 Marketplace listing
views**, segmented by product.

1. Listing discovery and install counts come from Marketplace SDK Analytics.
2. A successful product build makes an authenticated OilPriceAPI request with
   `X-OilPriceAPI-Client: <product-id>/<version>`. API logs can therefore count
   first data activation without collecting spreadsheet contents.
3. Sidebar signup links use:
   `utm_source=workspace_marketplace`,
   `utm_medium=addon`, and a unique `utm_campaign`.
4. Signup and paid conversion are joined to that campaign in the existing
   first-party attribution pipeline.

The add-ons do not send spreadsheet contents, formulas, cell values, Google
account identifiers, or API keys for analytics.

## Build and verify

```bash
npm ci
npm run portfolio:build
npm run test:portfolio
npm run portfolio:verify
npm run validate
```

Deployable Apps Script roots are written to `portfolio/dist/<product-id>/`.
Run clasp from the selected product directory so its `.claspignore` exposes only
`Code.gs`, `Sidebar.html`, and `appsscript.json`.

## Release order

1. Crack Spread Lab
2. Fuel Surcharge Studio
3. Energy Curve Builder
4. Bunker Voyage Planner
5. Gas Spread Monitor

The first two have the clearest non-overlapping search intent and quickest
time-to-value. Start their test deployments first, measure activation, then
promote the strongest funnel before submitting the next Marketplace listing.
