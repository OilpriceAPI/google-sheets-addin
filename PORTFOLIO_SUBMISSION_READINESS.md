# Google Workspace portfolio submission readiness

Reviewed: July 30, 2026

## Release candidates

All five products have distinct code, titles, workflows, Apps Script projects,
immutable versions, listing copy, reviewer guides, least-privilege manifests,
activation headers, UTM campaigns, and graphic asset sets. None has been
submitted to Google Workspace Marketplace.

| Rollout | Product | Apps Script candidate | Planned Cloud project | Remaining account-bound evidence |
| --- | --- | --- | --- | --- |
| 1 | Crack Spread Lab by OilPriceAPI | [Script](https://script.google.com/d/1c2O84bJoprkUtyo8eHb-yYmz1Mtttpo-8miUrKf3o3rzrVXUBfFU8M5C/edit), version 3 | `oilpriceapi-crack-spread` | Cloud/OAuth/Marketplace draft, installed-draft smoke, screenshot, video |
| 2 | Fuel Surcharge Studio by OilPriceAPI | [Script](https://script.google.com/d/1Mii2a-nGgRmrnsV1rl_9wElmZfBrmhJgufYPqvDEjvN_s9xTQ8WHRtxN/edit), version 4 | `oilpriceapi-fuel-surcharge` | Cloud/OAuth/Marketplace draft, installed-draft smoke, screenshot, video |
| 3 | Energy Curve Builder by OilPriceAPI | [Script](https://script.google.com/d/1q3YQIyE17nv4uNLQV7Pdw3DwKFu8Yrp9Uwylxw6WPmWhwyKzv4xZZDLZ/edit), version 1 | `oilpriceapi-energy-curve` | Cloud/OAuth/Marketplace draft, installed-draft smoke, screenshot, video |
| 4 | Bunker Voyage Planner by OilPriceAPI | [Script](https://script.google.com/d/1aQLBBWhyd_ffw1h9gmVFromI-5MwQsGd2v9REoUhXOHziDgbTA2eqY-u/edit), version 2 | `oilpriceapi-bunker-voyage` | Cloud/OAuth/Marketplace draft, installed-draft smoke, screenshot, video |
| 5 | Gas Spread Monitor by OilPriceAPI | [Script](https://script.google.com/d/1Od5dLY-A8l-sULQidIuso3rJRjJvWIaI54JZpPQlCVyAk32phv2D1YSU/edit), version 2 | `oilpriceapi-gas-spread` | Cloud/OAuth/Marketplace draft, installed-draft smoke, screenshot, video |

## Verified now

- The complete shared runtime and every product-specific builder compile.
- Every customer-critical build produces the documented workbook tabs and
  editable Google Sheets formulas.
- Missing, revoked, unentitled, rate-limited, timed-out, malformed, and
  incomplete API responses produce recovery messages or fail closed.
- The Energy Curve Builder uses the production
  `/v1/futures/ice-wti/curve` and `/v1/futures/ice-brent/curve` contracts.
- A production API smoke built all five workbook models with current,
  source-timestamped non-customer test data.
- Remote verification re-cloned every recorded immutable Apps Script version
  and matched its code, sidebar, and manifest to the local release package.
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

1. Create the separate standard Google Cloud project shown in the table.
2. Enable the Apps Script API and Google Workspace Marketplace SDK.
3. Link the standalone Apps Script project to the standard Cloud project's
   numeric project number.
4. Configure External OAuth branding using the exact Marketplace application
   name, an organization-controlled support contact, the product guide, and the
   shared Workspace privacy and terms pages.
5. Add only the three manifest scopes. Keep the project in Testing while the
   draft and reviewer fixture are prepared.
6. Configure a Public Editor add-on Marketplace draft with the recorded Script
   ID and immutable version. Public visibility is permanent after it is saved.
7. Install the Marketplace draft—not only the Apps Script editor test
   deployment—select **Use in this document**, and refresh the blank test
   spreadsheet.
8. Run the product's `REVIEWER_GUIDE.md` with a synthetic non-customer API key,
   inspect Apps Script executions, and confirm all documented recovery states.
9. Capture a unique, full-bleed 1280×800 PNG from the exact installed build and
   record a scope-complete OAuth demonstration video.
10. Stop before OAuth or Marketplace submission until the preceding product's
    Google review feedback has been applied across the shared runtime.

## Hold and rollout decision

The original OilPriceAPI listing remains the policy canary. Do not submit these
five listings in parallel. Once the original clears OAuth and Marketplace
review:

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
