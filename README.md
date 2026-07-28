# OilPriceAPI for Google Sheets™

Deployment-ready Editor add-on for source-aware OilPriceAPI formulas in Google
Sheets™.

> Google Workspace Marketplace publication is pending. The public listing was
> submitted on July 26, 2026 and rejected on July 27 pending trademark
> attribution and OAuth verification remediation. Do not claim Marketplace
> availability until Google approves and publishes the listing.

Dataset access, history, freshness, and limits depend on the API key, source,
and account entitlement. Review the
[versioned product-facts contract](https://api.oilpriceapi.com/product-facts.json)
before publishing derived product claims.

## Excel-equivalent formulas

Google Sheets custom-function names cannot contain a dot. These underscore
names are the direct equivalents of the Excel add-in surface:

| Google Sheets | Excel | Result |
| --- | --- | --- |
| `OILPRICE_PRICE(code)` | `OILPRICE.PRICE(code)` | Latest numeric API value |
| `OILPRICE_GET(path, query)` | `OILPRICE.GET(path, query)` | Allowlisted API table |
| `OILPRICE_CODES()` | `OILPRICE.CODES()` | Available commodity-code table |
| `OILPRICE_STATUS(code)` | `OILPRICE.STATUS(code)` | API freshness state |
| `OILPRICE_UNIT(code)` | `OILPRICE.UNIT(code)` | Currency/unit |
| `OILPRICE_INFO(code)` | `OILPRICE.INFO(code)` | Source, timestamp, unit, and freshness table |

Examples:

```text
=OILPRICE_PRICE("WTI_USD")
=OILPRICE_INFO("WTI_USD")
=OILPRICE_GET("/v1/prices/latest", "by_code=WTI_USD")
```

The original `OILPRICE(code)` formula remains supported for existing sheets.

## Additional Sheets formulas

| Function | Behavior | Cache |
| --- | --- | --- |
| `OILPRICE(code)` | Backward-compatible numeric latest price | 5 minutes |
| `OILPRICE_HISTORY(code, days)` | Source timestamp and price rows | 1 hour |
| `OILPRICE_CONVERT(code)` | Reference USD/MMBtu conversion for documented mappings | Latest-price cache |
| `BUNKER_PRICE(port, fuel)` | Numeric Data Connector bunker price | None |
| `BUNKER_PORT_PRICES(port)` | Bunker-price table with units and timestamp | None |
| `FUTURES_PRICE(contract)` | Numeric first-contract price | 5 minutes |
| `FUTURES_CURVE(contract)` | Month, price, and change rows | 5 minutes |
| `RIG_COUNT(type)` | Oil, gas, total, or source-dated table | 1 hour |

## Runtime and security contract

- API keys are stored in Apps Script document properties scoped to the current
  spreadsheet so custom formulas can use them. Editors of that spreadsheet can
  cause add-on formulas to make requests with the configured key.
- The sidebar receives only configured/not-configured state; it never reads the
  stored key into browser-side HTML.
- Generic GET calls are restricted to the same reviewed endpoint catalog as
  the Excel preview.
- Credential-shaped query keys are rejected before any network request.
- Missing, invalid, locked, rate-limited, timed-out, malformed, and empty
  responses fail with worksheet-readable recovery text.
- Latest-request diagnostics contain endpoint path, status, duration,
  timestamp, and optional request ID—never the API key or query string.
- The manifest requests only current-sheet, external-request, and container-UI
  scopes and restricts URL fetches to `api.oilpriceapi.com`.

## Validate locally

Node.js 20 or newer is required.

```bash
npm ci
npm test
npm run validate
```

The validation suite covers formula parity, credential lifecycle, negative
auth/entitlement/quota paths, response-shape drift, stale cache, source
metadata, Data Connector filtering and sheet output, endpoint/query
allowlisting, deployment packaging, Marketplace asset dimensions, public
claims, and secret scanning.

For a production API smoke:

```bash
OILPRICEAPI_KEY="your non-customer test key" npm run test:live
```

The standard smoke skips the account-gated Data Connector checks. Run those
with an entitled non-customer account and known valid filters:

```bash
OILPRICEAPI_KEY="your non-customer test key" \
OILPRICEAPI_DATA_CONNECTOR_SMOKE=1 \
OILPRICEAPI_DATA_CONNECTOR_PORT="SINGAPORE" \
OILPRICEAPI_DATA_CONNECTOR_FUEL="VLSFO" \
npm run test:live
```

The live-smoke script does not print the key or filter values.

## Deploy

Follow [DEPLOYMENT_GUIDE.md](DEPLOYMENT_GUIDE.md). The short operator sequence
after the Apps Script project exists is:

```bash
npm ci
npm run clasp:login
read -r "OPA_SCRIPT_ID?Apps Script ID: "
npm run clasp:configure -- "$OPA_SCRIPT_ID"
npm run deploy:push
npm run deploy:version -- "OilPriceAPI for Google Sheets 1.2.0"
```

Editor add-on publication uses the Apps Script **script ID and version
number**, not a web-app deployment ID. Test the Editor add-on before entering
that version in the Marketplace SDK.

Prepared listing copy, scope justifications, required screenshot shots, and
generated assets are in [MARKETPLACE_LISTING.md](MARKETPLACE_LISTING.md).

## Canonical links

- [Product facts](https://api.oilpriceapi.com/product-facts.json)
- [API documentation](https://docs.oilpriceapi.com)
- [Pricing and dataset access](https://www.oilpriceapi.com/pricing)
- [Data usage](https://www.oilpriceapi.com/legal/data-usage)
- [Apps Script quotas](https://developers.google.com/apps-script/guides/services/quotas)

## License

MIT
