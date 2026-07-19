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
- fresh and stale cache behavior;
- historical source timestamps;
- the 25-code batch guard; and
- user metadata without invented tier or limits.

`public-claims.test.js` checks every indexed/reference surface for explicit
Marketplace status, a product-facts link, and banned mutable claims.

`validate_code.js` validates the Apps Script manifest, JavaScript syntax, UI
bindings, private key getter, and required reference files.

`scripts/scan-secrets.sh` prints filenames only and fails on common API-key
patterns. It does not print matched values.

The automated suite does not replace the customer-critical manual Sheet smoke
in [DEPLOYMENT_GUIDE.md](../DEPLOYMENT_GUIDE.md).
