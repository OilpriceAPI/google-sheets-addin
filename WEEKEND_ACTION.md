# Saturday action: submit OilPriceAPI for Sheets

Target date: Saturday, July 25, 2026

## Outcome

Finish the Google-account work that cannot be automated, submit the public
Google Workspace Marketplace listing, and save proof of the submission.

The standalone repository and Apps Script release are already deployed. Use:

- Apps Script project:
  `https://script.google.com/d/1rlVWvciYu-wzqnY009I3oW-08ZPazYK1snrrMg9NNY7c5WBSkUK8W2Hb/edit`
- Script ID:
  `1rlVWvciYu-wzqnY009I3oW-08ZPazYK1snrrMg9NNY7c5WBSkUK8W2Hb`
- Immutable Apps Script version: `2`
- GitHub release:
  `https://github.com/OilpriceAPI/google-sheets-addin/releases/tag/v1.2.0`
- Full operator runbook: `DEPLOYMENT_GUIDE.md`
- Approved listing copy and scope justifications: `MARKETPLACE_LISTING.md`

Do not deploy the older `OilPriceAPI Sheets…` Apps Script project. The project
linked above is the reviewed production target.

## Checklist

### 1. Link the standard Google Cloud project

- [ ] Sign in as `karl.waldman@gmail.com`.
- [ ] Open the Apps Script project above and select **Project Settings**.
- [ ] Replace the default Apps Script Cloud project with the organization-owned
      standard Google Cloud project's numeric project number.
- [ ] Confirm the Apps Script API and Google Workspace Marketplace SDK are
      enabled in that Cloud project.
- [ ] Add a backup organization-controlled collaborator to the Apps Script and
      Cloud projects.

### 2. Configure OAuth

- [ ] Configure the OAuth consent screen as **External**.
- [ ] Use app name `OilPriceAPI for Sheets`.
- [ ] Use `support@oilpriceapi.com` for user support and developer contact.
- [ ] Use:
      `https://www.oilpriceapi.com/privacy/google-sheets-addon`
      and
      `https://www.oilpriceapi.com/terms/google-sheets-addon`.
- [ ] Declare exactly these scopes:
      `https://www.googleapis.com/auth/spreadsheets.currentonly`
      and
      `https://www.googleapis.com/auth/script.external_request`.
- [ ] Add the publisher and clean smoke-test accounts as test users while OAuth
      remains in Testing status.

Do not add Drive-wide, email, profile, or `userinfo.email` scopes.

### 3. Create and smoke-test the Editor add-on

- [ ] Create a blank spreadsheet containing no customer data.
- [ ] In Apps Script, select **Deploy → Test deployments → Editor add-on**.
- [ ] Test **Latest Code** against the blank spreadsheet.
- [ ] Open the sidebar and confirm it initially reports no stored API key.
- [ ] Save a non-customer OilPriceAPI test key and test the connection.
- [ ] Run:

```text
=OILPRICE_PRICE("WTI_USD")
=OILPRICE_UNIT("WTI_USD")
=OILPRICE_INFO("WTI_USD")
=OILPRICE_STATUS("WTI_USD")
=OILPRICE_CODES()
=OILPRICE_GET("/v1/prices/latest","by_code=WTI_USD")
```

- [ ] Confirm unsupported endpoints and query keys such as `api_key`, `token`,
      and `password` fail before a request.
- [ ] Delete the stored key and confirm the sidebar returns to no-key state.
- [ ] Review Apps Script Executions for exceptions, retries, unexpected 4xx/5xx
      responses, or any credential/query-string leakage.

### 4. Capture submission proof

- [ ] Capture at least one real 1280×800 PNG showing the installed add-on,
      sidebar, and actual formula results.
- [ ] Include unit, source timestamp, or freshness context in a screenshot.
- [ ] Confirm no key, customer data, account identifier, or clipboard content is
      visible.
- [ ] Save reviewed screenshots under `assets/marketplace/screenshots/`.

### 5. Configure and submit Marketplace

- [ ] In Google Workspace Marketplace SDK, add an **Editor add-on** integration
      for Google Sheets.
- [ ] Enter the Script ID above and immutable version `2`.
- [ ] Enter the same two OAuth scopes.
- [ ] Choose **Public** visibility before saving. Google treats this choice as
      permanent; do not use Private as temporary staging.
- [ ] Upload:
      `assets/marketplace/app-icon-32.png`,
      `assets/marketplace/app-icon-128.png`,
      `assets/marketplace/card-banner-220x140.png`,
      and the approved screenshot.
- [ ] Paste the approved copy and links from `MARKETPLACE_LISTING.md`.
- [ ] Submit OAuth verification if Google requests it.
- [ ] Submit the Marketplace listing for review.

## Done when

- [ ] The Marketplace console shows a submitted/review state.
- [ ] The submission confirmation or receipt is saved.
- [ ] The exact submitted Script ID and version are recorded as:
      `1rlVWvciYu-wzqnY009I3oW-08ZPazYK1snrrMg9NNY7c5WBSkUK8W2Hb`, version `2`.
- [ ] No public page claims Marketplace availability before Google approves the
      listing.
- [ ] After approval, a separate clean account installs the public listing and
      repeats the customer-critical smoke test before marketing copy changes.

