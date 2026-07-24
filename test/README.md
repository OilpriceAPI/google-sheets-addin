# Test Suite

Run from the repository root with Node.js 20 or newer:

```bash
npm test
npm run validate
```

`runtime.test.js` executes `Code.gs` in a VM with deterministic mocks for user
properties, user cache, HTTP responses, and Apps Script services. It covers:

- credential save/status/delete without returning the key;
- missing, invalid, and revoked key recovery;
- locked dataset, rate or quota limit, and timeout behavior;
- malformed JSON, empty `200`, and record schema drift;
- the production flat latest-price response;
- Excel-equivalent PRICE, GET, CODES, STATUS, UNIT, and INFO behavior;
- all reviewed generic GET endpoints and sensitive-query rejection;
- nested futures, diesel, and keyed-price table rendering;
- secret-free request diagnostics;
- fresh and stale cache behavior;
- historical source timestamps;
- the 25-code batch guard; and
- user metadata without invented tier or limits.

`public-claims.test.js` checks every indexed/add-on surface for explicit
Marketplace status, a product-facts link, and banned mutable claims.

`validate_code.js` validates the Apps Script manifest, JavaScript syntax, UI
bindings, private key getter, and required release files.

`scripts/verify-deploy-package.js` proves clasp can expose only the four
reviewed runtime files. `scripts/verify-marketplace-assets.js` proves the icon
and card-banner PNGs have Google's required formats and dimensions.

`scripts/scan-secrets.sh` prints filenames only and fails on common API-key
patterns. It does not print matched values.

The automated suite does not replace the customer-critical manual Sheet smoke
in [DEPLOYMENT_GUIDE.md](../DEPLOYMENT_GUIDE.md).
