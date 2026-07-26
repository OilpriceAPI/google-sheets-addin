# Weekend action: submit OilPriceAPI for Sheets

Target date: Saturday, July 25, 2026

Completed: Sunday, July 26, 2026

## Outcome

Finish the Google-account work that cannot be automated, submit the public
Google Workspace Marketplace listing, and save proof of the submission.

The standalone repository and Apps Script release are already deployed. Use:

- Apps Script project:
  `https://script.google.com/d/1rlVWvciYu-wzqnY009I3oW-08ZPazYK1snrrMg9NNY7c5WBSkUK8W2Hb/edit`
- Script ID:
  `1rlVWvciYu-wzqnY009I3oW-08ZPazYK1snrrMg9NNY7c5WBSkUK8W2Hb`
- Immutable Apps Script version: `7`
- GitHub release:
  `https://github.com/OilpriceAPI/google-sheets-addin/releases/tag/v1.2.0`
- Full operator runbook: `DEPLOYMENT_GUIDE.md`
- Approved listing copy and scope justifications: `MARKETPLACE_LISTING.md`

Do not deploy the older `OilPriceAPI Sheets…` Apps Script project. The project
linked above is the reviewed production target.

## Checklist

### 1. Link the standard Google Cloud project

- [x] Sign in with the publisher-controlled Google account.
- [x] Open the Apps Script project above and select **Project Settings**.
- [x] Replace the default Apps Script Cloud project with the organization-owned
      standard Google Cloud project's numeric project number.
- [x] Confirm the Apps Script API and Google Workspace Marketplace SDK are
      enabled in that Cloud project.
- [x] Add a backup organization-controlled collaborator to the Apps Script and
      Cloud projects.

### 2. Configure OAuth

- [x] Configure the OAuth consent screen as **External**.
- [x] Use app name `OilPriceAPI for Sheets`.
- [x] Use an organization-controlled support and developer contact.
- [x] Use:
      `https://www.oilpriceapi.com/privacy/google-sheets-addon`
      and
      `https://www.oilpriceapi.com/terms/google-sheets-addon`.
- [x] Declare these three functional scopes:
      `https://www.googleapis.com/auth/spreadsheets.currentonly`,
      `https://www.googleapis.com/auth/script.external_request`, and
      `https://www.googleapis.com/auth/script.container.ui`.
- [x] Accept Google's two mandatory identity defaults,
      `userinfo.email` and `userinfo.profile`, in the submitted consent/listing
      configuration. The Apps Script manifest still declares only the three
      functional scopes above.
- [x] Add the publisher and clean smoke-test accounts as test users while OAuth
      remains in Testing status.

Do not add Drive-wide scopes or use the Google identity defaults for product
behavior.

### 3. Create and smoke-test the Editor add-on

- [x] Create a blank spreadsheet containing no customer data.
- [x] In Apps Script, select **Deploy → Test deployments → Editor add-on**.
- [x] Test **Latest Code** against the blank spreadsheet.
- [x] Open the sidebar and confirm it initially reports no stored API key.
- [x] Save a non-customer OilPriceAPI test key and test the connection.
- [x] Run:

```text
=OILPRICE_PRICE("WTI_USD")
=OILPRICE_UNIT("WTI_USD")
=OILPRICE_INFO("WTI_USD")
=OILPRICE_STATUS("WTI_USD")
=OILPRICE_CODES()
=OILPRICE_GET("/v1/prices/latest","by_code=WTI_USD")
```

- [x] Confirm unsupported endpoints and query keys such as `api_key`, `token`,
      and `password` fail before a request.
- [x] Delete the stored key and confirm the sidebar returns to no-key state.
- [x] Review Apps Script Executions for exceptions, retries, unexpected 4xx/5xx
      responses, or any credential/query-string leakage.

### 4. Capture submission proof

- [x] Capture at least one real 1280×800 PNG showing the installed add-on,
      sidebar, and actual formula results.
- [x] Include unit, source timestamp, or freshness context in a screenshot.
- [x] Confirm no key, customer data, account identifier, or clipboard content is
      visible.
- [x] Save reviewed screenshots under `assets/marketplace/screenshots/`:
      `sheets-addon-sidebar-prices-1280x800.png`.

### 5. Configure and submit Marketplace

- [x] In Google Workspace Marketplace SDK, add an **Editor add-on** integration
      for Google Sheets.
- [x] Enter the Script ID above and immutable version `7`.
- [x] Enter the three functional OAuth scopes plus Google's two mandatory
      identity defaults.
- [x] Choose **Public** visibility before saving. Google treats this choice as
      permanent; do not use Private as temporary staging.
- [x] Upload:
      `assets/marketplace/app-icon-32.png`,
      `assets/marketplace/app-icon-128.png`,
      `assets/marketplace/card-banner-220x140.png`,
      and the approved screenshot.
- [x] Paste the approved copy and links from `MARKETPLACE_LISTING.md`.
- [x] Submit the Marketplace listing for review.
- [ ] Complete Google's OAuth verification if requested during review.

## Done when

- [x] The Marketplace console shows a submitted/review state.
- [x] The locked-review confirmation is recorded below.
- [x] The exact submitted Script ID and version are recorded as:
      `1rlVWvciYu-wzqnY009I3oW-08ZPazYK1snrrMg9NNY7c5WBSkUK8W2Hb`, version `7`.
- [x] No public page claims Marketplace availability before Google approves the
      listing.
- [ ] After approval, a separate clean account installs the public listing and
      repeats the customer-critical smoke test before marketing copy changes.

## Submission receipt

Submitted on July 26, 2026:

- listing: `OilPriceAPI for Sheets`;
- state: **In review**;
- console proof: **“The draft is in review and can't be edited”**, with the
  review-cancellation control present and draft saving disabled;
- Google Cloud project: `oilpriceapi-sheets-addon` (`991152473434`);
- Apps Script version: `7`;
- integration: Google Sheets Editor add-on;
- installation: individual and administrator install;
- regions: all regions.

Review caveats:

- the OAuth consent screen was still in Testing when submitted; Google may
  require it to move to In production;
- two submitted identity scopes are marked unverified, so Google may route the
  listing through OAuth verification;
- neither caveat blocked submission, and the review must not be canceled merely
  to change repository copy.

Versions `2`, `3`, and `4` do not include the required
`script.container.ui` scope. Versions `5` and `6` still read formula
credentials from the spreadsheet owner's user properties instead of the
spreadsheet-scoped credential. Do not submit or restore versions `2` through
`6`.
