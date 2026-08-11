# OAuth verification submission packet

This packet is for the production Google Cloud project
`oilpriceapi-sheets-addon` (`991152473434`) and the original
`OilPriceAPI for Google Sheets™` add-on.

## Status as of August 11, 2026

Customer-visible Google evidence:

- The listing is publicly available at
  `https://workspace.google.com/marketplace/app/oilpriceapi_for_google_sheets/991152473434`.
- The public Marketplace configuration points to immutable Apps Script version
  **11**, runtime `1.2.2`.
- OAuth verification was submitted on July 30, 2026 for Cloud project
  `991152473434`. The reviewer-accessible continuous demonstration is
  `https://youtu.be/FakNSmBddhE`.
- A Marketplace draft install previously proved custom-function registration,
  spreadsheet-scoped key save, connection/schema validation,
  `OILPRICE_PRICE("WTI_USD")`, `OILPRICE_CODES()`, and sidebar batch fetch.
- Public availability was independently rechecked unauthenticated on August
  10, 2026; the canonical URL returned the OilPriceAPI listing while a bogus
  application ID returned Google error 400.

Historical release evidence remains relevant: PR #22 supplied the
custom-function credential-context fix, immutable version 11 was cut after it,
and the three functional scopes below were used for submission. Earlier draft
versions 9 and 10 are superseded.

Private Cloud-console fields are not inferred from public availability. The
customer release gate is the public listing plus an installed-add-on formula
smoke. Immutable version 12 (runtime `1.3.0`) failed cache-isolation review and
was never published. Runtime `1.3.1` must therefore remain a release candidate until it is
merged, pushed, cut as a new immutable Apps Script version, installed through
Marketplace, and smoke-tested before App Configuration is updated.

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

## Future release order

1. Validate and merge the exact reviewed source.
2. Push it to the production Apps Script project and create a new immutable
   version without changing the public Marketplace configuration.
3. Install and smoke that candidate with a non-customer account.
4. Select the candidate in Marketplace App Configuration only after key,
   formula, batching, quota-recovery, and key-deletion checks pass.
5. Repeat the smoke through the public listing and review Apps Script logs.

Do not claim that a new runtime is public before its selected-version smoke.
