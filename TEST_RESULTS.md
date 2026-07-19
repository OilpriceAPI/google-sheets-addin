# Verification Results

Verification date: 2026-07-19

## Automated Runtime

Command:

```bash
npm test
```

Result: 19 tests passed, 0 failed.

Covered behavior:

- credential save, status, and deletion without returning the stored value;
- missing key and HTTP `401`, `403`, and `429` recovery;
- Apps Script fetch timeout handling;
- malformed JSON and empty successful responses;
- successful records missing price, currency, unit, or timestamp;
- current flat latest-price response parsing;
- fresh and stale cache envelopes;
- historical source timestamp preservation;
- 25-code batch limit;
- batch metadata writing and reference conversion behavior;
- missing exchange-rate data without fabricated fallbacks;
- schema-valid connection test; and
- no fabricated account tier, usage, or limits.

## Structure And Claims

Command:

```bash
npm run validate
```

Result:

- Apps Script syntax and UI bindings valid.
- Public reference and Marketplace status checks passed.
- Unsupported mutable-claim checks passed.
- Filename-only secret scan passed.
- `appsscript.json` retains only current-sheet and external-request scopes.

## Production Formula Path

Command shape:

```bash
OILPRICEAPI_KEY="..." npm run test:live
```

The key is passed by environment, sent to `curl` through standard input, removed
from the child environment, and never printed.

Result:

- `OILPRICE("WTI_USD")` returned a finite positive production value.
- A second call returned the cached value.
- `OILPRICE_HISTORY("WTI_USD", 1)` returned 16 records.
- The most recent returned source timestamp was
  `2026-07-19T14:10:50.373Z`.

## Deployment Scope

No `.clasp.json` or local `clasp` credentials are stored in this repository, so
the source checkout cannot silently deploy into a maintainer's Apps Script
project. Manual clean-sheet verification follows
[DEPLOYMENT_GUIDE.md](DEPLOYMENT_GUIDE.md). This repository is not currently
published in the Google Workspace Marketplace.
