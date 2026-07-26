# Verification Results

Verification date: 2026-07-25

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

- 28 runtime/public-claims tests passed, 0 failed.
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
  no fabricated exchange-rate fallbacks or account limits.

## Dependency audit

Command:

```bash
npm audit --omit=dev
```

Result: 0 runtime vulnerabilities.

The current official clasp development dependency reports moderate transitive
development-tool advisories. It is not shipped to Apps Script. Do not use clasp
to serve untrusted local files; update it when Google publishes a dependency
refresh.

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

## Account-bound release gates

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

Remaining Google account/Marketplace gates:

- link that script to the standard Google Cloud project;
- install and smoke the Editor add-on test deployment in a clean Sheet;
- capture at least one real 1280x800 screenshot;
- configure OAuth and Marketplace SDK with the Script ID and version;
- complete any required OAuth verification and submit the public listing.

See [DEPLOYMENT_GUIDE.md](DEPLOYMENT_GUIDE.md).
