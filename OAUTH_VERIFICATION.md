# OAuth verification submission packet

This packet is for the production Google Cloud project
`oilpriceapi-sheets-addon` (`991152473434`) and the original
`OilPriceAPI for Google Sheets™` add-on.

## Status as of July 30, 2026

Completed release evidence:

- Add-on PR 19 merged as
  `799f49f46059340ed332431f5c7ac87f5c91a695`.
- All 54 runtime, recovery, disclosure, deployment-package, asset, portfolio,
  and secret-scan checks passed for the current release.
- The exact merged runtime was pushed to the production Apps Script project.
- Immutable Apps Script version 11 was created for runtime `1.2.2`, merge
  `3747fef2d09474c5b610bfa7154c7134a16e6a9f`.
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

- Google rejected/paused the Marketplace version 9 submission pending OAuth
  approval and instructed the publisher not to resubmit Marketplace yet.
- Apps Script version 11 is the prepared release and must replace version 9 in
  Marketplace App Configuration after OAuth approval.
- OAuth publishing status is **In production**.
- OAuth branding is **verified and shown to users**.
- OAuth data access was **submitted July 30, 2026 and is under review**.
- Google reports that Homepage requirements are under review.
- Reviewer video: `https://youtu.be/FakNSmBddhE`.
- Google confirmed that Trust and Safety received the form, expects the first
  email within 3–5 days, and says full review can take up to 4–6 weeks.
- Submission receipt:
  `https://github.com/OilpriceAPI/google-sheets-addin/issues/20#issuecomment-5131610169`.

Remaining owner-session work:

1. Monitor the project contact email and answer Trust and Safety in the same
   case.
2. Keep the reviewer video, reviewer key, branding, URLs, and scope list stable
   during review.
3. After OAuth approval, update Marketplace App Configuration from version 9
   to version 11 and resubmit once.
4. Preserve OAuth approval and Marketplace resubmission receipts in issue 20:
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

| Scope | Reviewer justification |
| --- | --- |
| `https://www.googleapis.com/auth/spreadsheets.currentonly` | The add-on reads only user-selected inputs required for an invoked feature and writes requested formulas, market-data tables, formatting, and conversion outputs in the spreadsheet where the add-on is open. It does not request broad Google Drive access. |
| `https://www.googleapis.com/auth/script.external_request` | The add-on sends authenticated HTTPS GET requests to `api.oilpriceapi.com` for market data explicitly requested by the user. Requests contain the user's OilPriceAPI key and reviewed market identifiers or filters; general spreadsheet contents are not transferred. |
| `https://www.googleapis.com/auth/script.container.ui` | The add-on displays its menu, API-key sidebar, price-selection dialog, informational alerts, diagnostics, and recovery actions inside the current spreadsheet. |

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

The recording must not expose an API key, customer data, browser password
manager, clipboard contents, or unrelated tabs. Use a dedicated demo Google
account because its identifier can appear as part of the required consent
screen, and keep that reviewer copy Unlisted.

The completed reviewer video is `https://youtu.be/FakNSmBddhE`. Keep it
Unlisted. Create a different sanitized public acquisition video using
`YOUTUBE_PROMOTION.md`.

## Submission order

1. Deploy the public homepage, privacy policy, and terms above.
2. Confirm all three URLs return `200` without authentication or redirects to
   another domain.
3. Confirm Search Console ownership for `oilpriceapi.com`.
4. Push and smoke the exact reviewed Apps Script source.
5. Create a new immutable Apps Script version.
6. Record and upload the demo video with link visibility enabled for the
   Google review team.
7. In Google Auth Platform, verify and publish branding first, then submit Data
   Access verification with the scope justifications and demo link.
8. When Google requires OAuth approval first, wait for it before entering the
   new Apps Script version and resubmitting Marketplace.

Do not claim that the add-on is publicly installable until Google approves and
publishes the Marketplace listing.
