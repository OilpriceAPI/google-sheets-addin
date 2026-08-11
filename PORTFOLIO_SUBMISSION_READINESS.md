# Google Workspace portfolio submission readiness

Reviewed: August 11, 2026

## Release candidates

All five products have distinct code, titles, workflows, Apps Script projects,
listing copy, reviewer guides, least-privilege manifests, activation headers,
UTM campaigns, and graphic asset sets. None has been submitted to Google
Workspace Marketplace. The July immutable versions are retained as historical
receipts but contain the retired unscoped credential fallback and old client
header; a new immutable version is required for every current package.

| Rollout | Product | Apps Script candidate | Planned Cloud project | Remaining account-bound evidence |
| --- | --- | --- | --- | --- |
| 1 | Crack Spread Lab by OilPriceAPI | [Script](https://script.google.com/d/1c2O84bJoprkUtyo8eHb-yYmz1Mtttpo-8miUrKf3o3rzrVXUBfFU8M5C/edit), new immutable version required (previous 3) | `oilpriceapi-crack-spread` | Immutable version, Cloud/OAuth/Marketplace draft, installed-draft smoke, screenshot, video |
| 2 | Fuel Surcharge Studio by OilPriceAPI | [Script](https://script.google.com/d/1Mii2a-nGgRmrnsV1rl_9wElmZfBrmhJgufYPqvDEjvN_s9xTQ8WHRtxN/edit), new immutable version required (previous 4) | `oilpriceapi-fuel-surcharge` | Immutable version, Cloud/OAuth/Marketplace draft, installed-draft smoke, screenshot, video |
| 3 | Energy Curve Builder by OilPriceAPI | [Script](https://script.google.com/d/1q3YQIyE17nv4uNLQV7Pdw3DwKFu8Yrp9Uwylxw6WPmWhwyKzv4xZZDLZ/edit), new immutable version required (previous 1) | `oilpriceapi-energy-curve` | Immutable version, Cloud/OAuth/Marketplace draft, installed-draft smoke, screenshot, video |
| 4 | Bunker Voyage Planner by OilPriceAPI | [Script](https://script.google.com/d/1aQLBBWhyd_ffw1h9gmVFromI-5MwQsGd2v9REoUhXOHziDgbTA2eqY-u/edit), new immutable version required (previous 2) | `oilpriceapi-bunker-voyage` | Immutable version, Cloud/OAuth/Marketplace draft, installed-draft smoke, screenshot, video |
| 5 | Gas Spread Monitor by OilPriceAPI | [Script](https://script.google.com/d/1Od5dLY-A8l-sULQidIuso3rJRjJvWIaI54JZpPQlCVyAk32phv2D1YSU/edit), new immutable version required (previous 2) | `oilpriceapi-gas-spread` | Immutable version, Cloud/OAuth/Marketplace draft, installed-draft smoke, screenshot, video |

## Verified now

- The complete shared runtime and every product-specific builder compile.
- Every customer-critical build produces the documented workbook tabs and
  editable Google Sheets formulas.
- Missing, revoked, unentitled, rate-limited, timed-out, malformed, and
  incomplete API responses produce recovery messages or fail closed.
- The Energy Curve Builder uses the production
  `/v1/futures/ice-wti/curve` and `/v1/futures/ice-brent/curve` contracts.
- The July production API smoke built all five workbook models with
  source-timestamped non-customer test data; it must be repeated for the current
  hardened source.
- Remote verification proved the July immutable versions no longer match the
  current hardened packages, so they cannot be submitted.
- All product landing, signup, pricing, privacy, terms, and support links return
  successful first-party responses.
- Each application name is within the 50-character Marketplace limit; every
  short description is within the 200-character limit.
- Each manifest requests only:
  `spreadsheets.currentonly`, `script.external_request`, and
  `script.container.ui`.
- Each product has unique 32px and 128px icons and a unique 220×140 card
  banner. Screenshots are intentionally not simulated or reused.

## Remaining console pass

Repeat this sequence for one product at a time:

1. Push the exact merged package, cut a new immutable version, and require
   `npm run portfolio:verify:remote` to match it byte-for-byte.
2. Create the separate standard Google Cloud project shown in the table.
3. Enable the Apps Script API and Google Workspace Marketplace SDK.
4. Link the standalone Apps Script project to the standard Cloud project's
   numeric project number.
5. Configure External OAuth branding using the exact Marketplace application
   name, an organization-controlled support contact, the product guide, and the
   shared Workspace privacy and terms pages.
6. Add only the three manifest scopes. Keep the project in Testing while the
   draft and reviewer fixture are prepared.
7. Configure a Public Editor add-on Marketplace draft with the recorded Script
   ID and immutable version. Public visibility is permanent after it is saved.
8. Install the Marketplace draft—not only the Apps Script editor test
   deployment—select **Use in this document**, and refresh the blank test
   spreadsheet.
9. Run the product's `REVIEWER_GUIDE.md` with a synthetic non-customer API key,
   inspect Apps Script executions, and confirm all documented recovery states.
10. Capture a unique, full-bleed 1280×800 PNG from the exact installed build and
   record a scope-complete OAuth demonstration video.
11. Stop before OAuth or Marketplace submission until the preceding product's
    Google review feedback has been applied across the shared runtime.

## Hold and rollout decision

The original OilPriceAPI listing is public and remains the policy canary. Do not
submit these five listings in parallel. Before the next listing:

1. Convert any reviewer feedback into shared tests and rebuild all pending
   packages.
2. Take Crack Spread Lab through OAuth verification and Marketplace review.
3. Measure listing views, installs, attributed signups, and first workbook
   activations.
4. Continue in rollout order only if the prior product is policy-clean and the
   acquisition funnel is producing qualified activity.

This sequencing preserves the search-intent experiment without multiplying an
unknown policy or runtime defect across five live listings.

## Current Google references

- [Publish an add-on](https://developers.google.com/workspace/add-ons/how-tos/publish-add-on-overview)
- [Create a Marketplace store listing](https://developers.google.com/workspace/marketplace/create-listing)
- [Marketplace app review requirements](https://developers.google.com/workspace/marketplace/about-app-review)
- [OAuth verification requirements](https://support.google.com/cloud/answer/13464321)
