# Deploy OilPriceAPI for Google Sheets™

This is the production operator runbook for the Google Sheets Editor add-on.
Code packaging is automated; Google account ownership, Cloud/OAuth settings,
test installation, screenshots, and submission require the publisher account.

## Current release gate

- Runtime version: `1.2.0`
- Production Apps Script ID:
  `1rlVWvciYu-wzqnY009I3oW-08ZPazYK1snrrMg9NNY7c5WBSkUK8W2Hb`
- Current immutable Apps Script version: `9`
- GitHub production release workflow:
  `https://github.com/OilpriceAPI/google-sheets-addin/actions/workflows/apps-script-release.yml`
- Marketplace status: resubmitted July 29, 2026; draft in Google review
- Runtime push/version and local deployment checks: complete
- Marketplace review receipt and a real 1280x800 screenshot: complete
- Separate OAuth verification dependency: add the end-to-end authorization
  demo-video link to the prepared manual appeal and confirm it

Do not claim Marketplace availability until Google publishes the listing.

## 1. Choose the publisher and Cloud project

Use an organization-controlled publisher account. Create or select a
**standard** Google Cloud project for the add-on; the Apps Script default Cloud
project cannot be used for publication.

Recommended app identity:

- OAuth app name: `OilPriceAPI for Google Sheets™`
- Marketplace app name: `OilPriceAPI for Google Sheets™`
- User support email: `support@oilpriceapi.com`
- Developer contact: `support@oilpriceapi.com`
- App domain: `oilpriceapi.com`
- Privacy policy:
  `https://www.oilpriceapi.com/privacy/google-sheets-addon`
- Terms:
  `https://www.oilpriceapi.com/terms/google-sheets-addon`

Add at least one backup collaborator to the Apps Script and Cloud projects.

## 2. Create the standalone Apps Script project

1. Open `https://script.google.com` with the publisher account.
2. Create a **standalone** project named `OilPriceAPI for Google Sheets™`.
3. Open Project Settings and copy its Script ID.
4. Under Google Cloud Platform Project, switch from the default project to the
   standard Cloud project's numeric project number.
5. At `https://script.google.com/home/usersettings`, enable the Apps Script API
   for the publisher account.

The add-on must remain a standalone Apps Script project. Do not create it as a
script bound to one spreadsheet.

## 3. Connect this repository and push reviewed runtime files

From this repository:

```bash
npm ci
npm run validate
npm run clasp:login
read -r "OPA_SCRIPT_ID?Apps Script ID: "
npm run clasp:configure -- "$OPA_SCRIPT_ID"
npm run clasp:status
npm run deploy:push
```

`npm run deploy:push` reruns validation first. `.claspignore` permits only these
files to reach Apps Script:

- `Code.gs`
- `Sidebar.html`
- `FetchDialog.html`
- `appsscript.json`

`.clasp.json` and clasp OAuth credentials are gitignored. Never commit either.

Open the remote project and confirm the files:

```bash
npx clasp open-script
```

## 4. Configure OAuth

In the linked standard Cloud project:

1. Configure the OAuth consent screen as an External app for a public release.
2. Use the app identity and legal URLs from step 1.
3. Add the publisher and smoke-test accounts as test users while the app is in
   Testing status.
4. Declare exactly these scopes:

```text
https://www.googleapis.com/auth/spreadsheets.currentonly
https://www.googleapis.com/auth/script.external_request
https://www.googleapis.com/auth/script.container.ui
```

The manifest, OAuth consent screen, and Marketplace SDK scope lists must match.
Google adds the default `userinfo.email` and `userinfo.profile` identity scopes
to the displayed OAuth configuration; preserve those defaults. Do not add
Drive-wide access or additional functional scopes.

If Google requires OAuth verification, submit the requested scope
justifications and a demo video showing installation, authorization, API-key
configuration, a formula result, and key deletion. The prepared justification
text is in `MARKETPLACE_LISTING.md`.

## 5. Install a test Editor add-on

Create a blank spreadsheet containing no customer data.

In the Apps Script editor:

1. Select **Deploy > Test deployments**.
2. Enable deployment types and choose **Editor add-on**.
3. Create a test using **Latest Code**.
4. Select the blank spreadsheet as the test document.
5. Save and execute the test.
6. Refresh the spreadsheet if the add-on menu is not immediately visible.

Test both authorization states where available: installed-but-not-enabled and
installed-and-enabled.

## 6. Customer-critical smoke

Use a non-customer OilPriceAPI test key. Never place the key in a cell, source
file, screenshot, URL, issue, or log.

Run this checklist:

- [ ] The add-on appears under Extensions and opens the OilPriceAPI sidebar.
- [ ] The sidebar initially reports no stored key.
- [ ] Saving a key clears the input and never displays the stored value.
- [ ] The saved key is scoped to the current spreadsheet through Apps Script
      document properties; an editor can use formulas but cannot read the key
      through the sidebar.
- [ ] Test connection validates both HTTP status and the response schema.
- [ ] `=OILPRICE_PRICE("WTI_USD")` returns a finite number.
- [ ] `=OILPRICE_UNIT("WTI_USD")` returns a currency/unit value.
- [ ] `=OILPRICE_INFO("WTI_USD")` spills source and timestamp context.
- [ ] `=OILPRICE_STATUS("WTI_USD")` returns an API freshness state when present.
- [ ] `=OILPRICE_CODES()` spills a readable catalog table.
- [ ] `=OILPRICE_GET("/v1/prices/latest","by_code=WTI_USD")` spills a readable table.
- [ ] `=OILPRICE_GET("/v1/users/me","")` fails as unsupported without a request.
- [ ] A query containing `api_key`, `token`, or `password` fails before a request.
- [ ] An invalid key produces an actionable auth error.
- [ ] An inaccessible dataset produces an entitlement recovery action.
- [ ] A rate-limited account produces a retry-later action.
- [ ] Deleting the key returns the sidebar to no-key state.
- [ ] Last-request diagnostics contain no key and no query string.
- [ ] Legacy `OILPRICE`, history, futures, bunker, rig-count, and conversion paths
      used by customers still work for the test account's entitlements.

Review Apps Script Executions after the smoke. There should be no raw key,
unexpected exception loop, repeated retries, or unexplained high-volume calls.

## 7. Capture submission proof

Capture real 1280x800 PNG screenshots from the test deployment. Required shots
are listed in `MARKETPLACE_LISTING.md`. Use square corners and full bleed.

Before saving each image:

- remove the key from the input and clipboard;
- ensure no customer data or account identifier is visible;
- show actual API results rather than example numbers;
- include source timestamp/unit/freshness context in at least one shot.

Place approved screenshots in `assets/marketplace/screenshots/`. Do not commit
screenshots until they have been reviewed for secrets and customer data.

## 8. Create the immutable Editor add-on version

After the smoke passes against the exact pushed source:

```bash
npm run deploy:version -- "OilPriceAPI for Google Sheets 1.2.0"
npm run deploy:list
```

Record the version number printed by clasp. Editor add-on publication uses the
**Script ID and version number**. It does not use a web-app deployment ID.

If code changes after the version is created, push, retest, and create a new
version. Never submit an older version number by accident.

### Optional GitHub release workflow

After the first manual setup works, configure the protected
`apps-script-production` GitHub environment with:

- `CLASPRC_JSON`: the publisher account's clasp credential JSON;
- `CLASP_JSON`: the production project's local `.clasp.json`.

The manual **Create Apps Script release version** workflow validates, pushes
only the reviewed runtime files, and creates an immutable version. Treat
`CLASPRC_JSON` as a sensitive refresh credential, restrict environment access,
and rotate it if exposed.

## 9. Configure Google Workspace Marketplace SDK

In the same standard Cloud project:

1. Enable **Google Workspace Marketplace SDK**.
2. Open **App Configuration**.
3. Add an **Editor add-on** integration for Google Sheets.
4. Enter the Apps Script project Script ID and the tested version number.
5. Enable individual installation and, if desired, administrator installation.
6. Enter the same three OAuth scopes used by the manifest and OAuth consent
   screen.
7. Choose visibility deliberately:
   - Public: any Marketplace user after Google review.
   - Private: only the selected Workspace organization, with no public review.

Google treats the saved visibility choice as permanent. For a public launch,
choose Public before saving; do not use Private as a temporary staging mode.

## 10. Complete the store listing

Use the exact copy and links in `MARKETPLACE_LISTING.md`.

Upload:

- `assets/marketplace/app-icon-32.png`
- `assets/marketplace/app-icon-128.png`
- `assets/marketplace/card-banner-220x140.png`
- at least one approved real screenshot

Verify locally before upload:

```bash
npm run verify:assets
```

Save the listing as a draft while OAuth verification or screenshot review is
pending.

## 11. Submit and post-publication smoke

1. Submit OAuth verification if required.
2. Submit the public Marketplace listing for review.
3. Track the Marketplace SDK publication status and review email sent to
   `support@oilpriceapi.com`.
4. After approval, install the public listing using a separate clean account.
5. Repeat the customer-critical smoke against the published version.
6. Review Apps Script execution logs for new errors, retries, unexpected
   authorization failures, and noisy request patterns.
7. Only then update public marketing copy to say the add-on is available from
   the Google Workspace Marketplace and add the real listing URL.

## Updating an approved release

Push and smoke-test the new code, create a new Apps Script version, then update
the version number on the Marketplace SDK App Configuration page. Do not create
a new app integration or change the Script ID.

If scopes change, update all three scope lists and complete any required OAuth
reverification before publishing the new version.
