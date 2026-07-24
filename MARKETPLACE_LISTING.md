# Google Workspace Marketplace Listing

Status: submission-ready copy for the Editor add-on. Do not claim Marketplace
availability until Google publishes the listing.

## App details

- Default language: English
- Application name: `OilPriceAPI for Sheets`
- Category: Business Tools
- Pricing: Free of charge with paid features
- Developer name: `OilPriceAPI`
- Developer website: `https://www.oilpriceapi.com`

Short description:

> Pull source-timestamped energy prices, units, freshness, futures, and catalog data into spreadsheet formulas.

Detailed description:

> OilPriceAPI for Sheets adds source-aware energy data formulas to Google
> Sheets. Configure an OilPriceAPI key once in the sidebar, then request the
> latest available value, its currency and unit, source timestamp, freshness
> state, or an allowlisted API table.
>
> Core formulas include OILPRICE_PRICE, OILPRICE_INFO, OILPRICE_STATUS,
> OILPRICE_UNIT, OILPRICE_CODES, and OILPRICE_GET. Existing OILPRICE,
> OILPRICE_HISTORY, futures, bunker-price, rig-count, and reference conversion
> formulas remain available.
>
> The add-on returns API values and metadata without inventing missing prices,
> timestamps, currencies, units, account limits, or fallback exchange rates.
> Dataset access, history, and freshness depend on the configured API key,
> source, and account entitlement.
>
> The API key is stored in per-user Apps Script properties. It is not written
> to spreadsheet cells, URLs, diagnostics, or browser-side HTML. Generic GET
> requests are restricted to a reviewed endpoint catalog and reject
> credential-shaped query parameters.

## Support links

- Terms: `https://www.oilpriceapi.com/terms/google-sheets-addon`
- Privacy: `https://www.oilpriceapi.com/privacy/google-sheets-addon`
- Support: `https://www.oilpriceapi.com/support`
- Setup: `https://oilpriceapi.github.io/google-sheets-addin/`
- Help: `https://docs.oilpriceapi.com`
- Report issue: `https://github.com/OilpriceAPI/google-sheets-addin/issues`

## OAuth scopes and justification

Use the same two scopes in the Apps Script manifest, OAuth consent screen, and
Marketplace SDK:

| Scope | Justification |
| --- | --- |
| `https://www.googleapis.com/auth/spreadsheets.currentonly` | Read and write only the spreadsheet where the user runs the add-on, including inserting formulas and writing requested data tables. |
| `https://www.googleapis.com/auth/script.external_request` | Send authenticated HTTPS GET requests to `api.oilpriceapi.com` for data explicitly requested by the user. |

The add-on does not request Drive-wide access, email, profile, or user-info
scopes.

## Graphic assets

- `assets/marketplace/app-icon-32.png`
- `assets/marketplace/app-icon-128.png`
- `assets/marketplace/card-banner-220x140.png`

Generate and verify them with:

```bash
npm run assets
npm run verify:assets
```

## Required real screenshots

Capture these from the test deployment at 1280x800 with square corners and no
padding. Do not use mock data or expose an API key.

1. Sidebar showing a configured key state and a successful connection test.
2. A sheet with `OILPRICE_PRICE`, `OILPRICE_UNIT`, and `OILPRICE_STATUS`.
3. A spilled `OILPRICE_INFO` table showing source timestamp and freshness.
4. A spilled `OILPRICE_CODES` or allowlisted `OILPRICE_GET` table.

At least one real screenshot is required for submission.
