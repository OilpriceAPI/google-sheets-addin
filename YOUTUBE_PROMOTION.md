# Public YouTube promotion plan

## Recommendation

Create a separate public promotional video for OilPriceAPI.

The business case is YouTube search discovery, branded visibility, qualified
referral traffic, and measurable signup or workbook activation. Do not model
the description URL as a guaranteed high-authority SEO backlink. Treat any
indexing or link-equity benefit as secondary.

Keep the OAuth reviewer video Unlisted. It contains review-specific screens and
an account identifier and is not the polished public asset.

YouTube's own guidance says unique, keyword-relevant titles and descriptions
help viewers find videos through search. Long-form descriptions can contain
clickable external links when the channel has advanced features enabled:

- `https://support.google.com/youtube/answer/12948449?hl=en`
- `https://support.google.com/youtube/answer/13748639?hl=en`

## Release sequence

The original add-on is publicly available in Google Workspace Marketplace.
Keep the two video concepts distinct so workflow discovery and installation
intent remain measurable.

### Video 1: API workflow tutorial

Focus on the underlying workflow and OilPriceAPI:

> How to Pull WTI and Brent Prices into Google Sheets™ | OilPriceAPI Tutorial

Show a clean spreadsheet, API setup, live values, units, source timestamps, and
freshness. Use a signup or integration-page CTA so this video measures the
underlying data workflow rather than add-on installation.

### Video 2: Marketplace installation tutorial

Focus on discovery and installation:

> OilPriceAPI Google Sheets™ Add-on: Install, Configure, and Build a Live Oil
> Price Sheet

Show the public Marketplace listing, installation, **Use in this document**,
configuration, formulas, batch fetch, and a finished workbook.

## Trademark and policy guardrails

- Use `OilPriceAPI` as the brand and publisher identity.
- Use `Google Sheets™` only to describe compatibility or the demonstrated
  workflow.
- Do not use Google branding in a way that implies sponsorship, certification,
  endorsement, or ownership.
- Do not place `Google`, `Google Sheets`, or another Google product trademark
  in the Marketplace product name for future portfolio products.
- Add this description footer:

> Google Sheets™ is a trademark of Google LLC. OilPriceAPI is not affiliated
> with or endorsed by Google LLC.

- Do not claim the add-on is available from Marketplace before Google publishes
  the listing.

## Public recording standard

Record a new public asset; do not republish the OAuth reviewer video.

- Use a dedicated demo account and a blank, professionally formatted workbook.
- Show no email address, OAuth warning page, API key, password manager,
  notifications, unrelated tabs, or capture controls.
- Use narration or concise on-screen callouts.
- Record at 1080p, 30 fps, and target 4–7 minutes.
- Lead with the finished outcome in the first 10 seconds.
- Demonstrate WTI, Brent, and natural gas, then source timestamp and unit
  metadata.
- Include one recovery state, such as a missing key, followed by the successful
  next action.
- End with one CTA.

## Search packaging

Primary topic:

`oil price data in Google Sheets`

Supporting phrases:

- WTI price Google Sheets;
- Brent crude price spreadsheet;
- oil price API tutorial;
- energy market data spreadsheet;
- commodity price formulas;
- live oil prices in a spreadsheet.

Use one primary phrase naturally in the title and first two description lines.
Use YouTube Analytics Research to validate the final wording before recording.
Do not stuff tags or repeat exact-match phrases unnaturally.

Suggested thumbnail:

- real spreadsheet crop with WTI and Brent values;
- OilPriceAPI droplet mark;
- 3–5 words: `LIVE OIL DATA → SHEETS`;
- no Google logo and no endorsement-style badge.

Suggested chapters:

```text
00:00 Live oil-price workbook
00:15 What OilPriceAPI provides
00:40 Configure the spreadsheet
01:30 WTI and Brent formulas
02:30 Units, sources, and timestamps
03:30 Fetch a market-data table
04:30 Common errors and recovery
05:15 Next step
```

## Description template

```text
Pull WTI, Brent, natural-gas, and other energy-market data into Google Sheets™
with OilPriceAPI. This tutorial shows live price formulas, units, source
timestamps, freshness metadata, and a multi-commodity table.

Start here:
https://www.oilpriceapi.com/integrations/google-sheets?utm_source=youtube&utm_medium=organic_video&utm_campaign=google_sheets_addon&utm_content=<video-id>_overview_demo

Create an OilPriceAPI account:
https://www.oilpriceapi.com/auth/signup?utm_source=youtube&utm_medium=organic_video&utm_campaign=google_sheets_addon&utm_content=<video-id>_signup_cta

Documentation:
https://docs.oilpriceapi.com

Google Sheets™ is a trademark of Google LLC. OilPriceAPI is not affiliated
with or endorsed by Google LLC.
```

Confirm that both destination URLs preserve UTM parameters through redirects
before publishing.

## Measurement contract

Google Analytics recommends consistent `utm_source`, `utm_medium`, and
`utm_campaign` values so referral sessions appear in Traffic acquisition:

`https://support.google.com/analytics/answer/10917952?hl=en`

Use:

```text
utm_source=youtube
utm_medium=organic_video
utm_campaign=google_sheets_addon
utm_content=<video-and-cta>
```

Track weekly for 90 days:

| Funnel stage | Metric | Source |
| --- | --- | --- |
| Discovery | Impressions and YouTube Search traffic | YouTube Analytics |
| Packaging | Impression click-through rate | YouTube Analytics |
| Engagement | Average percentage viewed and 30-second retention | YouTube Analytics |
| Intent | Description-link clicks or GA4 sessions | YouTube/GA4 |
| Acquisition | Signups attributed to the UTM campaign | First-party attribution |
| Activation | First successful add-on/API request with product client header | OilPriceAPI logs |
| Revenue | Paid conversions attributed to the campaign | Billing attribution |

North-star metrics:

1. activated OilPriceAPI accounts per 1,000 video views;
2. activated workbooks per 100 YouTube-referred landing-page sessions.

Use a unique `utm_content` for each CTA and video, such as:

- `overview_demo_description`;
- `install_tutorial_description`;
- `pinned_comment`;
- `channel_profile`.

Do not put email addresses, Google account IDs, spreadsheet contents, API keys,
or other user data into analytics events.

## 90-day launch test

1. Publish one outcome-led tutorial.
2. Add the tracked integration-page link to the first two description lines.
3. Add a tracked signup link below it.
4. Add the integration-page link to the channel profile.
5. Publish one short excerpt that points to the long-form related video; Shorts
   description URLs are not clickable.
6. Review YouTube search terms and audience retention after 7 days.
7. Treat Day 30 as the interim review and Day 90 as the final success decision.
8. Compare YouTube-referred signup and activation rates with Marketplace and
   organic-search traffic.
9. Produce the next portfolio-product video only if the first video produces
   qualified visits or activation signal, not merely views.

Success threshold for the first experiment:

- at least 100 qualified landing-page sessions, or
- at least 10 attributed signups, or
- at least 3 first API/add-on activations

within 90 days. If none occur, revise the topic, thumbnail, CTA, or landing-page
match before scaling the series.
