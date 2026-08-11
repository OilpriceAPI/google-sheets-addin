# Google Workspace Marketplace Listing

Status: **publicly available** at
`https://workspace.google.com/marketplace/app/oilpriceapi_for_google_sheets/991152473434`.

Public Marketplace Apps Script version: `11` (runtime `1.2.2`). Runtime
`1.3.1` remains a release candidate until an immutable Apps Script version is
cut, smoke-tested through an installed Marketplace add-on, and selected in App
Configuration.

## App details

- Default language: English
- Application name: `OilPriceAPI for Google Sheets™`
- Category: Accounting and Finance
- Pricing: Free of charge with paid features
- Developer name: `OilPriceAPI`
- Developer website: `https://www.oilpriceapi.com`

Short description:

> Pull source-timestamped energy prices, units, freshness, futures, and catalog data into spreadsheet formulas.

Detailed description:

> OilPriceAPI for Google Sheets™ adds source-aware energy data formulas to
> Google Sheets™. Configure an OilPriceAPI key once in the sidebar, then
> request the latest available value, its currency and unit, source timestamp,
> freshness state, or an allowlisted API table.
>
> Core formulas include OILPRICE_PRICE, OILPRICE_TABLE, OILPRICE_INFO, OILPRICE_STATUS,
> OILPRICE_UNIT, OILPRICE_CODES, and OILPRICE_GET. Existing OILPRICE,
> OILPRICE_HISTORY, futures, bunker-price, rig-count, and reference conversion
> formulas remain available.
>
> The add-on returns API values and metadata without inventing missing prices,
> timestamps, currencies, units, account limits, or fallback exchange rates.
> Dataset access, history, and freshness depend on the configured API key,
> source, and account entitlement.
>
> The API key is stored in Apps Script properties scoped to the current
> spreadsheet so its formulas can retrieve data. The primary copy uses
> document properties. A compatibility copy uses the spreadsheet owner's user
> properties with the spreadsheet ID in the property name for Google's
> custom-function authorization context. It is not written to spreadsheet
> cells, URLs, diagnostics, or browser-side HTML. Spreadsheet editors can cause
> installed add-on formulas to make requests using the configured key. Generic
> GET requests are restricted to a reviewed endpoint catalog and reject
> credential-shaped query parameters.
>
> Google Sheets™ is a trademark of Google LLC.

## Support links

- Terms: `https://www.oilpriceapi.com/terms/google-sheets-addon`
- Privacy: `https://www.oilpriceapi.com/privacy/google-sheets-addon`
- Support: `https://www.oilpriceapi.com/support`
- Setup: `https://oilpriceapi.github.io/google-sheets-addin/`
- Help: `https://docs.oilpriceapi.com`
- Report issue: `https://github.com/OilpriceAPI/google-sheets-addin/issues`

## OAuth scopes and justification

The Apps Script manifest declares these three functional scopes:

| Scope                                                      | Justification                                                                                                                       |
| ---------------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------- |
| `https://www.googleapis.com/auth/spreadsheets.currentonly` | Read and write only the spreadsheet where the user runs the add-on, including inserting formulas and writing requested data tables. |
| `https://www.googleapis.com/auth/script.external_request`  | Send authenticated HTTPS GET requests to `api.oilpriceapi.com` for data explicitly requested by the user.                           |
| `https://www.googleapis.com/auth/script.container.ui`      | Display the add-on menu, API-key sidebar, help alerts, and data-fetch dialog inside the spreadsheet where the user runs the add-on. |

The submitted OAuth/Marketplace configuration also displays Google's mandatory
`userinfo.email` and `userinfo.profile` defaults. The add-on does not use those
identity defaults for product behavior and does not request Drive-wide access.

## Current publication receipt

- Google Cloud project: `oilpriceapi-sheets-addon` (`991152473434`)
- Public listing:
  `https://workspace.google.com/marketplace/app/oilpriceapi_for_google_sheets/991152473434`
- Public Marketplace Apps Script version: `11`
- Latest immutable Apps Script version: `12`
- Runtime release represented by version 11: `1.2.2`
- Version 12 represented runtime `1.3.0`, failed the user-scoped cache-isolation
  review, and was never published. Runtime `1.3.1` is its replacement candidate.
- Superseded: version 10 (`1.2.1`), cut before the PR #22 credential fix
- Integration: Google Sheets Editor add-on
- Install modes: individual and administrator
- Regions: all regions
- Review state: **Published — independently verified August 10, 2026**
- Rejection email received: **July 27, 2026**
- Historical Google Cloud receipt: **“The draft is in review and can't be edited.”**
- Verification-page production deployment:
  `https://github.com/OilpriceAPI/website-clean/actions/runs/30434284989`
- OAuth verification was submitted July 30, 2026 with the reviewer-accessible
  demonstration `https://youtu.be/FakNSmBddhE`; the later public listing is the
  customer-visible approval evidence.

The OAuth consent screen is **In production**. The manifest and published
submission use the three functional scopes above. Google's default `userinfo.email` and
`userinfo.profile` scopes remain in place. The public homepage, privacy policy,
and terms were corrected and deployed from website PR 1461. Cache-busted
production checks returned HTTP 200 from the canonical domain and found the
expected disclosure text on all three pages.

Do not infer any private Cloud-console field beyond the receipts above. The
public listing and installed-add-on smoke are the release gates for customer
availability; future runtime revisions still require their own immutable
version and installed formula smoke.

## Graphic assets

- `assets/marketplace/app-icon-32.png`
- `assets/marketplace/app-icon-128.png`
- `assets/marketplace/card-banner-220x140.png`
- `assets/marketplace/screenshots/sheets-addon-sidebar-prices-1280x800.png`

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

The first reviewed real screenshot is
`assets/marketplace/screenshots/sheets-addon-sidebar-prices-1280x800.png`. It
shows the installed sidebar, stored-key state, successful connection test,
live prices, units, source labels, source timestamps, and retrieval timestamps
without displaying the API key or a Google account identifier.
