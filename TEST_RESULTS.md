# Verification Results

Verification date: 2026-07-26

## Apps Script scope correction

The sidebar and dialog smoke test exposed a release-manifest mismatch: the
runtime calls `SpreadsheetApp.getUi().showSidebar(...)` and
`showModalDialog(...)`, but immutable versions `2`, `3`, and `4` declared only
the spreadsheet and external-request scopes.

Direct `clasp` inspection on 2026-07-25 confirmed:

- versions `2`, `3`, and `4` contain two OAuth scopes and must not be submitted;
- versions `5` and `6` add
  `https://www.googleapis.com/auth/script.container.ui`, but still resolve
  formula credentials through the spreadsheet owner's user properties;
- version `7` stores the credential in Apps Script document properties scoped
  to the spreadsheet, with a legacy user-property migration fallback;
- version `7` contains the same `Code`, `Sidebar`, `FetchDialog`, and manifest
  source as the reviewed local runtime.

## Automated release validation

Command:

```bash
npm run validate
```

Result:

- 39 runtime/public-claims tests passed, 0 failed.
- Apps Script syntax, required functions, UI bindings, scopes, and fetch
  allowlist are valid.
- The deployment package exposes only `Code.gs`, `Sidebar.html`,
  `FetchDialog.html`, and `appsscript.json`.
- Marketplace icons are valid 32x32 and 128x128 PNGs.
- The Marketplace card banner is a valid 220x140 PNG.
- The reviewed Marketplace screenshot is a valid 1280x800 PNG showing the
  installed sidebar, successful connection, live data, units, and source
  timestamps without a visible key or account identifier.
- Unsupported mutable-claim checks passed.
- Filename-only secret scan passed.

Covered behavior includes:

- spreadsheet-scoped credential save, legacy user-property migration, status,
  and deletion without returning the stored value;
- missing/invalid key, locked dataset, rate limit, timeout, and server recovery;
- malformed JSON, empty success, and schema-drift rejection;
- stable worksheet error codes, including API invalid-code suggestions;
- Excel-equivalent PRICE, GET, CODES, STATUS, UNIT, and INFO formulas;
- all reviewed GET endpoint patterns and credential-query rejection;
- flat price, keyed price, diesel, and nested futures table rendering;
- source timestamp, unit, freshness, and diagnostic preservation;
- fresh/stale cache envelopes, history, batch limit, conversion behavior, and
  no fabricated exchange-rate fallbacks or account limits;
- Data Connector response validation, filter normalization and encoding,
  account-lock/rate-limit/timeout recovery, both bunker formulas, the
  nine-column green-header sheet writer, success alert, and empty-data
  recovery.

## Dependency audit

Command:

```bash
npm audit --audit-level=moderate
```

Result on 2026-08-11: 0 vulnerabilities. The audit includes the development
toolchain used to validate and publish the Apps Script package. `@google/clasp`
remains pinned to `3.3.0`; the lockfile override resolves its transitive `uuid`
dependency to `11.1.1` without downgrading the release CLI.

## Production API smoke

Command:

```bash
OILPRICEAPI_KEY="non-customer test key" npm run test:live
```

Result on 2026-07-24:

- `OILPRICE("WTI_USD")` and `OILPRICE_PRICE("WTI_USD")` returned `88.96`.
- A second latest-price call used the cached value.
- `OILPRICE_UNIT("WTI_USD")` returned `USD/barrel`.
- `OILPRICE_INFO("WTI_USD")` retained source timestamp
  `2026-07-24T17:43:28.389Z`.
- Allowlisted `OILPRICE_GET("/v1/prices/latest", "by_code=WTI_USD")` returned a
  non-empty table.
- `OILPRICE_HISTORY("WTI_USD", 1)` returned 100 timestamped records.

The smoke used the existing non-customer key from the local environment. The
script passes the key to `curl` through standard input and never prints it.

The Data Connector portion is intentionally opt-in because it requires an
entitled account and known valid port/fuel filters:

```bash
OILPRICEAPI_KEY="non-customer entitled test key" \
OILPRICEAPI_DATA_CONNECTOR_SMOKE=1 \
OILPRICEAPI_DATA_CONNECTOR_PORT="known valid port" \
OILPRICEAPI_DATA_CONNECTOR_FUEL="known valid fuel" \
npm run test:live
```

When enabled, the smoke exercises the menu fetch, validates the nine-column
sheet and green header, and runs both bunker custom functions against the
production API. This entitled live smoke has not yet been run.

## Account-bound release and submission gates

Completed on 2026-07-24:

- created the publisher-owned standalone Apps Script project
  `OilPriceAPI for Sheets`;
- pushed exactly `Code.gs`, `Sidebar.html`, `FetchDialog.html`, and
  `appsscript.json`;
- created immutable Apps Script versions `1` through `7`;
- configured the `apps-script-production` GitHub environment for `main` only;
- stored clasp credentials as environment secrets;
- ran the production release workflow successfully:
  `https://github.com/OilpriceAPI/google-sheets-addin/actions/runs/30116108120`.

Production release target:

- Script ID:
  `1rlVWvciYu-wzqnY009I3oW-08ZPazYK1snrrMg9NNY7c5WBSkUK8W2Hb`
- Current immutable version: `7`
- Apps Script editor:
  `https://script.google.com/d/1rlVWvciYu-wzqnY009I3oW-08ZPazYK1snrrMg9NNY7c5WBSkUK8W2Hb/edit`

Completed on 2026-07-26:

- linked the script to Google Cloud project `oilpriceapi-sheets-addon`;
- installed and smoked the Editor add-on test deployment in a clean Sheet;
- uploaded the reviewed 1280x800 runtime screenshot and all required graphics;
- configured the Marketplace SDK with Apps Script version `7`;
- submitted the public listing for review;
- recorded the locked state: **“The draft is in review and can't be edited.”**

Remaining Google-controlled gates:

- complete OAuth verification if Google requests it;
- wait for Marketplace approval and publication;
- after approval, install from the public listing with a separate clean account
  and repeat the customer-critical smoke before changing availability claims.

Remaining Data Connector acceptance gate:

- run the opt-in production smoke with a Data Connector-enabled non-customer
  account, then repeat the menu item and both formulas inside the submitted
  Apps Script version in a clean spreadsheet.

See [DEPLOYMENT_GUIDE.md](DEPLOYMENT_GUIDE.md).
