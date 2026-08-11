# Google Sheets add-on launch playbook

This playbook captures the production lessons from the original
`OilPriceAPI for Google Sheets™` submission so later OilPriceAPI add-ons can
reuse the proven path without repeating the same diagnostics.

Use this together with:

- `DEPLOYMENT_GUIDE.md` for the release commands and Google configuration;
- `MARKETPLACE_LISTING.md` for reviewed listing copy and assets;
- `OAUTH_VERIFICATION.md` for the current submission record;
- `PORTFOLIO.md` for product order and acquisition measurement.

## Proven sequence

Follow this order. Reordering the Google steps creates ambiguous failures and
avoidable review delays.

1. Create a standalone Apps Script project and link it to a standard,
   organization-controlled Google Cloud project.
2. Deploy the exact reviewed source, run automated validation, smoke it with a
   non-customer key, and create an immutable Apps Script version.
3. Configure identical functional scopes in the Apps Script manifest, Google
   Auth Platform Data Access page, and Workspace Marketplace SDK.
4. Create public app-specific homepage, privacy, terms, support, and opt-out
   URLs. Verify the authorized domain through the same organization-controlled
   account used for the Cloud project.
5. Configure OAuth as External and move it to In production.
6. Publish and verify OAuth branding before trying to prepare Data Access
   verification.
7. Add the publisher or dedicated reviewer account as a Marketplace draft
   tester.
8. Install the real Marketplace draft, select **Use in this document**, refresh
   the spreadsheet, and run the installed-add-on smoke.
9. Record a clean OAuth review video from that installed draft.
10. Submit OAuth Data Access verification once and preserve the receipt.
11. Wait for OAuth approval before submitting or resubmitting the Marketplace
    listing when Google has explicitly required that sequence.
12. Update the Marketplace integration to the exact tested Apps Script version,
    submit once, and repeat the installed-add-on smoke after publication.

## Scope contract

Every OilPriceAPI Sheets add-on currently needs only these three functional
scopes:

```text
https://www.googleapis.com/auth/spreadsheets.currentonly
https://www.googleapis.com/auth/script.external_request
https://www.googleapis.com/auth/script.container.ui
```

Google may display `userinfo.email` and `userinfo.profile` as identity defaults.
Do not add those defaults to product logic, and do not add Drive-wide access.

Use this combined Data Access justification:

> OilPriceAPI for Google Sheets uses spreadsheets.currentonly to read
> user-selected inputs and write formulas and requested market-data tables only
> in the spreadsheet where the add-on is open; it does not request Drive-wide
> access. It uses script.external_request only to send HTTPS GET requests to
> api.oilpriceapi.com for market data the user explicitly requests. Requests
> include the user's stored OilPriceAPI key and selected commodity identifiers
> or filters; general spreadsheet contents are not transferred. It uses
> script.container.ui to display the add-on menu, API-key sidebar,
> price-selection dialog, help alerts, diagnostics, and recovery actions inside
> Google Sheets. These are the narrowest scopes available: currentonly limits
> spreadsheet access to the active document, external_request is required to
> call the OilPriceAPI service, and container.ui is required for the in-sheet
> sidebar and dialogs.

## The two Google test paths are not equivalent

An Apps Script Editor add-on test deployment can prove menu, sidebar, dialog,
authorization, and API request behavior. It did not register the custom
function namespace in fresh test documents during this submission:

- sidebar key save and Fetch Latest Available Prices worked;
- formula autocomplete was absent;
- even a service-free diagnostic custom function returned `#NAME?`;
- no formula execution appeared in Apps Script logs.

That combination proves a host registration problem, not an OilPriceAPI
authentication failure.

The real Marketplace draft install registered the functions correctly only
after this sequence:

1. Add the Google account as a Marketplace draft tester.
2. Install the draft from its Marketplace page.
3. Open a new blank spreadsheet.
4. Select **Extensions → Add-ons → Manage add-ons**.
5. Open the OilPriceAPI add-on menu and select **Use in this document**.
6. Refresh the spreadsheet.

After that, autocomplete appeared and `OILPRICE_PRICE`, `OILPRICE_CODES`, and
the other functions executed normally.

Do not spend API debugging time on a `#NAME?` result when:

- the sidebar fetch works;
- formula autocomplete is absent; and
- Apps Script shows no formula invocation.

First confirm installed-add-on activation and namespace registration.

## Credential-context lesson

Sidebar handlers and custom functions can run in different Apps Script
authorization contexts. A key that works in the sidebar does not by itself
prove a formula can retrieve it.

The production runtime therefore stores:

- a primary copy in document properties; and
- a compatibility copy in the spreadsheet owner's user properties, keyed by
  spreadsheet ID.

The two prototype candidates also wrote an unscoped user-property key. Current
packages intentionally ignore that value because it cannot be tied to its
original spreadsheet. Prototype testers must save the key again in every
spreadsheet they continue to use.

The spreadsheet owner should configure the key. The OilPriceAPI account email
does not need to match the Google account email. The sidebar must expose only
configured/not-configured state; never return the stored credential to HTML,
cells, diagnostics, logs, screenshots, URLs, or issue trackers.

Test all four credential states:

1. no key;
2. valid key;
3. invalid or revoked key;
4. deleted key with diagnostic state cleared.

## Avoid duplicate add-on contexts

Simultaneous Apps Script test deployments and Marketplace draft installs can
produce duplicate extension entries and make results impossible to interpret.

Before the final smoke:

1. remove obsolete Apps Script test deployments;
2. uninstall stale draft installs if necessary;
3. confirm only one OilPriceAPI add-on is installed;
4. create a new spreadsheet;
5. enable the installed draft for that document;
6. refresh before testing formulas.

Name test spreadsheets with the product, version, and path, for example:
`OPA Marketplace Draft v11 Smoke`.

## Reviewer fixture

Create one persistent synthetic reviewer user and API key per listing.

- Use a clearly synthetic, non-customer email identity.
- Suppress lifecycle email, billing, and product analytics for the fixture.
- Give it only the datasets and conservative quota needed for review.
- Use a descriptive key name with the product and date.
- Verify one valid request and one invalid-key request before recording.
- Keep the key active through review, but never commit or post it.
- Deliver it only through Google's private reviewer field or existing review
  email thread.

Do not rotate or delete the fixture while a case is open unless exposure is
suspected.

## Review video checklist

The reviewer video and the public marketing video are different assets.
Keep the reviewer video Unlisted when it shows an account identifier or
unverified-app screen.

Before recording:

- enable Do Not Disturb;
- close unrelated tabs and applications;
- disable visible password-manager prompts;
- use a blank spreadsheet and synthetic key;
- set the OAuth consent language to English;
- record at 1080p or better;
- rehearse the exact flow once.

Show:

1. Marketplace installation and the complete OAuth consent flow;
2. the exact app name and every requested permission;
3. **Use in this document** activation;
4. the sidebar and masked/configured credential state;
5. successful connection and response-schema validation;
6. custom-function autocomplete and a live numeric result;
7. sidebar batch fetch writing into the active spreadsheet;
8. source, unit, timestamp, and freshness metadata.

Before upload, sample the full rendered video at frequent intervals. Remove:

- raw keys, passwords, cookies, or clipboard contents;
- browser autofill and password-manager overlays;
- notification panels;
- Terminal windows and unrelated tabs;
- macOS capture controls;
- customer or personal spreadsheet data.

Upload the reviewer copy as Unlisted and verify the URL while signed out before
submitting it.

## Google Auth submission notes

- Branding must be published and verified before Data Access verification is
  enabled.
- The Verification Center may describe unverified branding as “not being shown
  to users” even when the underlying fields are complete. Open Branding,
  publish it, then return to Verification Center.
- The Data Access form accepts one combined justification of up to 1,000
  characters and a reviewer-accessible YouTube URL.
- Confirm that Restricted scopes shows no rows.
- Save Data Access changes before returning to Verification Center.
- Submit once. Repeated submissions can trigger a cooldown and delay review.
- Preserve the exact confirmation text, date, screenshot, case ID, and review
  email thread.

For the original add-on, Google confirmed receipt on July 30, 2026, said the
first Trust and Safety email should arrive within 3–5 days, and warned that the
full review can take up to 4–6 weeks.

## What can and cannot be automated

Automate:

- manifest and scope parity;
- packaging and secret scans;
- API success and negative-path tests;
- asset dimensions;
- legal/disclosure URL checks;
- immutable version records;
- API-client attribution and activation measurement.

Keep a real-account smoke for:

- Marketplace draft tester installation;
- **Use in this document** activation;
- custom-function registration and autocomplete;
- OAuth consent presentation;
- Google review submission and receipt capture.

Google does not expose every Marketplace draft-install or custom-function
registration behavior through public Apps Script or Sheets APIs. API-written
formulas cannot bypass an unregistered custom-function namespace.

## Per-add-on evidence packet

Create a separate packet for every portfolio product:

- Cloud project ID and numeric project number;
- Apps Script project ID and immutable version;
- Git commit and validation output;
- exact three-scope comparison;
- public homepage, privacy, terms, support, and opt-out URLs;
- dedicated reviewer fixture record;
- installed-draft smoke spreadsheet;
- real listing screenshots;
- unlisted OAuth review video;
- OAuth receipt and case thread;
- Marketplace receipt and published listing URL;
- post-publication smoke and log review.

Do not reuse screenshots, detailed listing descriptions, reviewer keys, Cloud
projects, OAuth identities, or Apps Script projects across portfolio listings.
