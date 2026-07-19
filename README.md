# OilPriceAPI Google Sheets Reference

Source-timestamped OilPriceAPI records in Google Sheets through a manually
installed Apps Script project.

> This is a reference implementation. It is not currently published in the
> Google Workspace Marketplace. Do not search for or install an add-on that
> claims to be this repository.

Dataset access, history, freshness, and limits depend on the API key, source,
and account entitlement. Review the
[versioned product-facts contract](https://api.oilpriceapi.com/product-facts.json)
before publishing derived product claims.

## Install In A Test Sheet

1. Create a blank Google Sheet.
2. Open **Extensions > Apps Script**.
3. Replace the default script with [Code.gs](Code.gs).
4. Add HTML files named `Sidebar` and `FetchDialog`, using
   [Sidebar.html](Sidebar.html) and [FetchDialog.html](FetchDialog.html).
5. Enable the manifest in Project Settings and replace it with
   [appsscript.json](appsscript.json).
6. Save, return to the sheet, and reload it.
7. Open **OilPriceAPI > Configure API Key** and save an API key.

Create or manage keys at
[oilpriceapi.com/auth/signup](https://www.oilpriceapi.com/auth/signup?utm_source=google_sheets&utm_medium=reference&utm_campaign=readme).
The script stores the key in Apps Script user properties and never writes it to
a cell, URL, log, or browser-side HTML.

## Verify One Formula

Enter this formula in a cell:

```text
=OILPRICE("WTI_USD")
```

The cell returns the numeric value from the latest available source record. To
inspect its source, unit, currency, and timestamp, use **OilPriceAPI > Fetch
Latest Available Prices**. The resulting `Data` sheet separates `Source
Timestamp` from the client-side `Retrieved At` timestamp.

## Functions

| Function | Executable behavior | Reference cache |
| --- | --- | --- |
| `OILPRICE(code)` | Numeric latest available price | 5 minutes |
| `OILPRICE_HISTORY(code, days)` | `[source timestamp, price]` rows; `days` selects a versioned lookback endpoint from 1 through 365 | 1 hour |
| `OILPRICE_CONVERT(code)` | Reference USD/MMBtu conversion for codes in `COMMODITY_MAP` | Shares latest-price cache |
| `BUNKER_PRICE(port, fuel)` | Numeric Data Connector price | None |
| `BUNKER_PORT_PRICES(port)` | Fuel, price, currency, unit, and source timestamp rows | None |
| `FUTURES_PRICE(contract)` | Numeric first contract price | 5 minutes |
| `FUTURES_CURVE(contract)` | Month, price, and change rows | 5 minutes |
| `RIG_COUNT(type)` | Oil, gas, total, or a source-dated table | 1 hour |

The dialog contains example commodity codes, not a guarantee of universal
catalog access. Check the current
[commodity catalog](https://www.oilpriceapi.com/commodities) and your account
entitlements.

Google Sheets controls formula recalculation. This script does not poll in the
background. A recalculation inside a cache window returns the cached source
record; a stale or malformed cache envelope is discarded before a new request.

## Failure Behavior

All network paths use the same response validator:

- Missing key links to key management.
- `401` identifies an invalid or revoked key.
- `402` and `403` link to current dataset access options.
- `429` explains that the rate or quota window must recover before retrying.
- Fetch exceptions expose timeout recovery without logging credentials.
- Empty or malformed successful responses fail instead of returning zero,
  guessed currency or units, fallback exchange rates, or the current time.
- Batch refresh accepts at most 25 selected codes per request.

## Conversion Scope

`OILPRICE_CONVERT` is a reference calculation. Its heat-content factors are
defined in `Code.gs`, and its currency conversion requires API-provided
`GBP_USD` and `EUR_USD` records when needed. It does not use fallback exchange
rates. Verify the factor and source unit for your analytical or commercial use.

## Security

- Keys are stored in per-user Apps Script properties, not encrypted by this
  repository.
- The sidebar receives only `{ configured: true|false }`; it cannot load the
  stored value back into the browser.
- Delete the key from the sidebar before sharing or transferring a test sheet.
- Do not place a key in cells, script source, screenshots, URLs, logs, issues,
  or analytics.
- Standard API plans do not imply unrestricted source-data redistribution.
  Review the [data-usage policy](https://www.oilpriceapi.com/legal/data-usage).

## Validate Locally

Node.js 20 or newer is required. The test suite runs `Code.gs` inside a mocked
Apps Script runtime.

```bash
npm test
npm run validate
```

The suite covers the current flat latest-price response, missing/invalid keys,
locked data, `429`, timeout, stale cache, empty and malformed responses,
timestamp preservation, and the batch limit. See
[DEPLOYMENT_GUIDE.md](DEPLOYMENT_GUIDE.md) for manual Apps Script verification.

## Canonical Links

- [Product facts](https://api.oilpriceapi.com/product-facts.json)
- [API documentation](https://docs.oilpriceapi.com)
- [Pricing and dataset access](https://www.oilpriceapi.com/pricing)
- [Data usage](https://www.oilpriceapi.com/legal/data-usage)
- [Apps Script quotas](https://developers.google.com/apps-script/guides/services/quotas)

## License

This repository is provided under the [MIT License](LICENSE).
