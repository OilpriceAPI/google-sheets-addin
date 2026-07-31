# Google Workspace Marketplace Listing

Status: the Marketplace Store Listing draft was resubmitted July 29, 2026 and
remains in Google review. That draft references Apps Script version 9, which
predates both PR #19 (OAuth verification prep) and PR #22 (custom-function
credential fix).

**Version 11 (runtime `1.2.2`) is the release candidate** and should replace
version 9 in App Configuration before publishing. App Configuration is
editable during review — only the Store Listing tab locks (verified
2026-07-31; see `OAUTH_VERIFICATION.md`), so this repin does not have to wait
for Google.

OAuth branding and data-access verification have not yet been submitted.
Do not claim Marketplace availability until Google approves and publishes the
listing.

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

## Submission receipt

- Google Cloud project: `oilpriceapi-sheets-addon` (`991152473434`)
- Marketplace draft Apps Script version: `9` (stale - repin to `11`)
- Latest immutable Apps Script version: `11`
- Runtime release represented by version 11: `1.2.2`
- Superseded: version 10 (`1.2.1`), cut before the PR #22 credential fix
- Integration: Google Sheets Editor add-on
- Install modes: individual and administrator
- Regions: all regions
- Review state: **In review — resubmitted July 29, 2026**
- Rejection email received: **July 27, 2026**
- Google Cloud receipt: **“The draft is in review and can't be edited.”**
- Verification-page production deployment:
  `https://github.com/OilpriceAPI/website-clean/actions/runs/30434284989`
- OAuth submission state: **not submitted**
- Canonical remaining-work issue:
  `https://github.com/OilpriceAPI/google-sheets-addin/issues/20`

The OAuth consent screen is **In production**. The manifest and prepared
submission use the three functional scopes above; the locked Marketplace draft
uses the same scopes. Google's default `userinfo.email` and
`userinfo.profile` scopes remain in place. The public homepage, privacy policy,
and terms were corrected and deployed from website PR 1461. Cache-busted
production checks returned HTTP 200 from the canonical domain and found the
expected disclosure text on all three pages.

Google Auth Platform still reports branding and data access as unverified.
Before submission, an owner/editor of Cloud project `991152473434` must confirm
Search Console ownership for `oilpriceapi.com`, record the required continuous
OAuth demonstration, enter version 11 in App Configuration (already editable -
it does not lock during review), submit branding and data-access
verification, and preserve the
resulting receipt. Track those actions only in issue 20 rather than opening
parallel submission issues.

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
