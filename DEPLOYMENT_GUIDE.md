# Google Sheets Reference Deployment

## Status

This repository is a manually installed reference implementation. It is not
currently published in the Google Workspace Marketplace. A GitHub release is a
source release, not a Marketplace listing or an installed Apps Script project.

## Prerequisites

- A Google account allowed to create Apps Script projects.
- A blank test spreadsheet that contains no customer data.
- An OilPriceAPI key with access to the dataset used in the smoke test.
- The four runtime files from one tagged repository release: `Code.gs`,
  `Sidebar.html`, `FetchDialog.html`, and `appsscript.json`.

Create or manage a key at https://www.oilpriceapi.com/auth/signup. Never paste a
real key into this repository, a cell, an issue, or a screenshot.

## Clean Setup

1. Create a blank Google Sheet and open **Extensions > Apps Script**.
2. Replace the default `.gs` file with `Code.gs` from the tagged release.
3. Add an HTML file named `Sidebar` and paste `Sidebar.html`.
4. Add an HTML file named `FetchDialog` and paste `FetchDialog.html`.
5. In Apps Script Project Settings, enable the manifest file.
6. Replace the manifest with `appsscript.json`.
7. Save the project, return to the spreadsheet, and reload it.
8. Confirm that the `OilPriceAPI` menu appears.
9. Open **OilPriceAPI > Configure API Key**, paste the key, select **Save**, and
   clear the clipboard.
10. Select **Test connection**. A valid response must pass both HTTP and schema
    validation.

## Customer-Critical Smoke

1. In an empty cell enter `=OILPRICE("WTI_USD")`.
2. Confirm that the cell returns a finite numeric value rather than zero, an
   empty cell, or an error.
3. Open **OilPriceAPI > Fetch Latest Available Prices** and select WTI.
4. Confirm that the `Data` sheet contains code, finite price, currency, unit,
   source timestamp, and a distinct retrieval timestamp.
5. Change the stored key to an invalid value and confirm the formula identifies
   an invalid or revoked key.
6. Delete the stored key from the sidebar and confirm the formula provides the
   key-management recovery action.

## Cache Verification

Latest price and futures records use a five-minute user cache. Historical and
rig-count records use a one-hour user cache. Google Sheets decides when a
formula recalculates; the script does not run a background refresh. For a clean
refresh test, wait for the relevant cache window or use a new Apps Script user
context.

## Entitlement Verification

Use a dataset the test account cannot access and confirm the response links to
https://www.oilpriceapi.com/pricing. Do not document a plan name, catalog total,
fixed source cadence, or history promise in deployment metadata. Those facts
come from https://api.oilpriceapi.com/product-facts.json and current API docs.

## Key Cleanup

Before sharing, duplicating, or transferring the sheet:

1. Open the sidebar and select **Delete**.
2. Confirm that the sidebar reports no stored key.
3. Review Apps Script execution logs and the spreadsheet for accidental values.
4. Run `npm run validate` in the source checkout.

## Marketplace Publication

Marketplace publication is outside this reference deployment. Do not describe
the project as available in Marketplace unless an installable listing has been
published and independently verified. Any future listing must use current
product facts, declare external API calls and user-property key storage, and
complete Google's OAuth and security review requirements.

## Local Verification

```bash
npm test
npm run validate
```

`clasp` deployment is intentionally not configured in this repository. That
avoids committing a script ID or using one maintainer's Apps Script project as
the reference installation.
