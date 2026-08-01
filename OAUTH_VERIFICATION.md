# OAuth verification submission packet

This packet is for the production Google Cloud project
`oilpriceapi-sheets-addon` (`991152473434`) and the original
`OilPriceAPI for Google Sheets™` add-on.

## Status as of July 31, 2026

Release evidence below is from July 29, 2026. The "Current Google state"
section was re-verified against the live Cloud console on July 31, 2026.

Completed release evidence:

- Add-on PR 19 merged as
  `799f49f46059340ed332431f5c7ac87f5c91a695`.
- All 49 runtime, recovery, disclosure, deployment-package, asset, portfolio,
  and secret-scan checks passed.
- The exact merged runtime was pushed to the production Apps Script project.
- Immutable Apps Script version 10 was created with description
  `OilPriceAPI for Google Sheets 1.2.1 OAuth verification`.
- A fresh clone of version 10 matched the four-file reviewed release package
  exactly.
- **Superseded by version 11.** PR #22 (custom-function credential fix) merged
  2026-07-29 22:01 UTC, after version 10 was cut, and changed `Code.gs` and
  `Sidebar.html`. Immutable version 11 was cut the same minute (18:01 EDT) and
  carries runtime `1.2.2`; its `Code.gs` reads `ADDON_VERSION = '1.2.2'`
  (verified 2026-07-31). Version 11 is the release candidate.
- Website PR 1461 merged as
  `c3acb510680992538315781fb0ce3dcec335bf20`.
- Production deployment
  `https://github.com/OilpriceAPI/website-clean/actions/runs/30434284989`
  completed successfully, including production and money-page smoke checks
  and a Cloudflare purge.
- Cache-busted checks returned HTTP 200 without a cross-domain redirect for
  the homepage, privacy policy, and terms. The responses contained the
  expected scope, Limited Use, and current-formula disclosures.
- DigitalOcean deployment `65aaf4e7-df6a-44aa-a89b-de796963e442` is ACTIVE.
  Its first 500 runtime log lines contained no matched errors, warnings,
  retries, timeouts, or 5xx responses.

Current Google state:

- The Marketplace **Store Listing** draft is in review. That tab reports
  "The draft is in review and can't be edited" and exposes a "Cancel review"
  control.
- The Marketplace **App Configuration** tab is _editable_ during that review.
  Verified 2026-07-31 by DOM inspection of the Cloud console: every input
  reports `disabled: false`, `readOnly: false`, with no `aria-disabled`. The
  Version field is a free-text `<input type="text">`, not a dropdown, and
  currently holds `9`. "Save Draft" is greyed only for want of unsaved
  changes.
  **Correction:** earlier revisions of this document asserted the App
  Configuration was locked during review. That is wrong, and it nearly drove
  an unnecessary cancel-and-recut. The accurate rule is: **Store Listing locks
  during review; App Configuration does not.**
- Apps Script **version 11** (runtime `1.2.2`) is the current release
  candidate and is not yet selected in App Configuration, which still points
  at version 9.
- OAuth publishing status is **In production**.
- OAuth branding is **not verified**.
- OAuth data access is **not verified**.
- OAuth verification has **not been submitted**.
- No public OAuth demonstration URL or Google submission receipt exists yet.

Remaining owner-session work:

1. Confirm that a Cloud project owner/editor is a verified Search Console owner
   for `oilpriceapi.com`.
2. Record and publish the continuous end-to-end OAuth demonstration below.
3. Update Marketplace App Configuration to Apps Script version **11**. This
   does not have to wait for Google - App Configuration is editable while the
   Store Listing is in review.
4. Submit OAuth branding and data-access verification with the exact scopes,
   justifications, and public video URL.
5. Capture the confirmation text, date, case/reference ID if present, and
   redacted screenshots in issue 20:
   `https://github.com/OilpriceAPI/google-sheets-addin/issues/20`.

## Branding values

Use the same values everywhere Google displays or reviews the app:

- App name: `OilPriceAPI for Google Sheets™`
- Homepage:
  `https://www.oilpriceapi.com/integrations/google-sheets`
- Privacy policy:
  `https://www.oilpriceapi.com/privacy/google-sheets-addon`
- Terms:
  `https://www.oilpriceapi.com/terms/google-sheets-addon`
- Authorized domain: `oilpriceapi.com`
- User support email: `support@oilpriceapi.com`
- Developer contact: `support@oilpriceapi.com`

The homepage is public without authentication, identifies the submitted app,
explains each requested permission, links the same privacy policy configured
on the consent screen, and states the Google Workspace API Limited Use
commitment.

Before submission, confirm that an owner or editor of Cloud project
`991152473434` is also a verified owner of the `oilpriceapi.com` domain
property in Google Search Console.

## Functional scopes

The Apps Script manifest, OAuth Data Access page, and Workspace Marketplace SDK
must contain the same three functional scopes:

| Scope                                                      | Reviewer justification                                                                                                                                                                                                                                                 |
| ---------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `https://www.googleapis.com/auth/spreadsheets.currentonly` | The add-on reads only user-selected inputs required for an invoked feature and writes requested formulas, market-data tables, formatting, and conversion outputs in the spreadsheet where the add-on is open. It does not request broad Google Drive access.           |
| `https://www.googleapis.com/auth/script.external_request`  | The add-on sends authenticated HTTPS GET requests to `api.oilpriceapi.com` for market data explicitly requested by the user. Requests contain the user's OilPriceAPI key and reviewed market identifiers or filters; general spreadsheet contents are not transferred. |
| `https://www.googleapis.com/auth/script.container.ui`      | The add-on displays its menu, API-key sidebar, price-selection dialog, informational alerts, diagnostics, and recovery actions inside the current spreadsheet.                                                                                                         |

Google may display default `userinfo.email` and `userinfo.profile` scopes. The
add-on does not use those identity scopes for product behavior and does not
read, store, or transmit the user's Google email or profile.

## Demo video runbook

Record one continuous end-to-end video using the exact production-branded app
and a non-customer test spreadsheet and OilPriceAPI key.

1. Show the public homepage and open its privacy-policy link.
2. Open a blank Google spreadsheet with no customer data.
3. Start the add-on authorization flow with an account that has not already
   granted the production app's permissions.
4. Set the consent-screen language to English and show the complete consent
   screen, exact app name, and all requested permissions before granting them.
5. Open the installed add-on sidebar and show its current-spreadsheet data
   notice plus privacy and terms links.
6. Save the non-customer OilPriceAPI key. Show that the input clears and only
   the configured/not-configured state is returned.
7. Run Test connection and show a successful, schema-validated response.
8. Enter `=OILPRICE_INFO("WTI_USD")` and show price, currency, unit, source,
   source timestamp, and freshness output.
9. Use Fetch Latest Available Prices to show the user-selected identifiers
   and the add-on writing the requested table into the current spreadsheet.
10. Show the last-request diagnostic without a key or query string.
11. Choose Delete API Key and show that the stored-key and diagnostic states
    are cleared.

The recording must not expose an API key, Google account identifier, customer
data, browser password manager, clipboard contents, or unrelated tabs.

## Submission order

1. Deploy the public homepage, privacy policy, and terms above.
2. Confirm all three URLs return `200` without authentication or redirects to
   another domain.
3. Confirm Search Console ownership for `oilpriceapi.com`.
4. Push and smoke the exact reviewed Apps Script source.
5. Create a new immutable Apps Script version and enter that version in the
   Marketplace SDK.
6. Record and upload the demo video with link visibility enabled for the
   Google review team.
7. In Google Auth Platform, verify branding first, then submit Data Access
   verification with the scope justifications and demo link.
8. Keep the Workspace Marketplace listing in review only after the OAuth
   verification request is accepted for review.

Do not claim that the add-on is publicly installable until Google approves and
publishes the Marketplace listing.
