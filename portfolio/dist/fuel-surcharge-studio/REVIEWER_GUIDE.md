# Fuel Surcharge Studio by OilPriceAPI — reviewer guide

Status: pre-submission. Evidence that depends on a real installed Marketplace draft is explicitly gated below.

## Reviewer prerequisites

- Google Workspace host: Google Sheets
- Access model: the reviewer installs the Editor add-on and supplies an OilPriceAPI test key provided privately in the Marketplace review instructions.
- No Google account identity match is required. The OilPriceAPI key may belong to a different email address.
- The reviewer fixture must be synthetic, non-customer, active for the review window, and entitled to: `DIESEL_RETAIL_USD`, `DIESEL_RETAIL_EAST_COAST_USD`, `DIESEL_RETAIL_MIDWEST_USD`, `DIESEL_RETAIL_GULF_COAST_USD`, `DIESEL_RETAIL_ROCKY_MOUNTAIN_USD`, `DIESEL_RETAIL_WEST_COAST_USD`.

## End-to-end review flow

1. Install the unpublished Marketplace draft in a blank spreadsheet.
2. In **Extensions → Add-ons → Manage add-ons**, select Fuel Surcharge Studio by OilPriceAPI and choose **Use in this document**.
3. Refresh the spreadsheet once.
4. Open **Extensions → Fuel Surcharge Studio → Configure OilPriceAPI key**.
5. Paste the private reviewer key and select **Save key**.
6. Select **Test connection** and confirm the green success result.
7. Select **Build workbook**.
8. Confirm these product-specific tabs exist: `Surcharge Calculator`, `Regional Indexes`, `Surcharge Table`.
9. Confirm source timestamps and units are visible and no API key is written to a cell.
10. Delete the stored key and confirm the sidebar reports that no key is configured.

## OAuth scope demonstration

- `spreadsheets.currentonly`: steps 7–9 create and format only the current spreadsheet.
- `script.external_request`: steps 6–7 request only the reviewed OilPriceAPI market data.
- `script.container.ui`: steps 2–7 display the add-on menu, sidebar, status, and recovery UI.

## Evidence to attach

- Apps Script ID: `1Mii2a-nGgRmrnsV1rl_9wElmZfBrmhJgufYPqvDEjvN_s9xTQ8WHRtxN`
- Immutable Apps Script version: `New immutable version required`
- Previous immutable version (superseded): `4`
- Marketplace draft install: Pending the separate Cloud project and draft integration.
- Clean test spreadsheet URL: Provide privately after installed-draft smoke.
- OAuth demo video URL: Record after the exact OAuth branding and scopes are configured.
- Reviewed 1280×800 screenshot: Capture after installed-draft smoke.
- Reviewer test credential: Provide privately; never commit.

## Expected recovery behavior

- Missing key: asks the reviewer to configure a key or create an account.
- Invalid or revoked key: asks the reviewer to replace it in the sidebar.
- Dataset unavailable: explains the entitlement problem and points to pricing.
- Rate or quota limit: asks the reviewer to retry later or review the account limit.
- Timeout/network failure: asks the reviewer to check the connection and retry.
- Malformed or incomplete success response: rejects the data instead of building a misleading workbook.
