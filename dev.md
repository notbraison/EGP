# Dev Notes

_Last updated: 2026-09-10_

venv/Scripts/activate

py main.py

Budget - [docs.google.com/spreadsheets/d/18PZby5HOyY1EWue6rIZL6kHKv0h1mN-2oa3EV-aF9kw/edit?usp=sharing](https://docs.google.com/spreadsheets/d/18PZby5HOyY1EWue6rIZL6kHKv0h1mN-2oa3EV-aF9kw/edit?usp=sharing)

Addresses - [docs.google.com/spreadsheets/d/1tCpLpMCHzEGwZzbk63q2j-KmXoG9SCUIoRPWl22MiIc/edit?usp=sharing](https://docs.google.com/spreadsheets/d/1tCpLpMCHzEGwZzbk63q2j-KmXoG9SCUIoRPWl22MiIc/edit?usp=sharing)

## What this project does

Automated pipeline that pulls Kenya e-GP public procurement plan data
(https://egpkenya.go.ke/public-app), builds a local master workbook
(`Procuring_Entities.xlsx`), generates one formatted workbook per procuring
entity (`./entity_plans/`), and syncs the budget + address-tracking sheets to
Google Sheets via Apps Script Web Apps.

## Architecture (current — API-based, post-rewrite)

The original implementation scraped the portal's HTML table and tried to
resolve `javascript:void(0)` anchor links via regex on `onclick` attributes.
That approach never worked: this is a production Angular build, and click
handlers are bound via `(click)="viewApp(item)"` templates, which compile
to JS event listeners with **no `onclick` attribute in the DOM at all**.
Confirmed via Playwright request/response logging.

Investigation also found the site sits behind bot protection (plain
`requests` gets `403` with no cookies set, even on a simple `GET` — the
`XSRF-TOKEN` cookie is only issued to a real browser session). So the
pipeline now uses a hybrid: **Playwright to mint a real, trusted browser
session once, then plain fast JSON POSTs through that same session**
(`context.request.post(...)`) for everything else. No more DOM table
walking, no more clicking, no more `networkidle` waits (which never
resolved anyway — see "Known site quirks" below).

### Key API endpoints (discovered via network interception)

All require header `X-XSRF-TOKEN: <value from XSRF-TOKEN cookie>`, obtained
by loading `https://egpkenya.go.ke/public-app` once in a real browser
context.

| Endpoint                                     | Purpose                                                                                                                                              | Body                                                                                                                        |
| -------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------- |
| `POST /api/app/public-app-detail`          | Entity/APP listing, paginated                                                                                                                        | `{appNumber:"", finYear:0, procuringEntity:null, pageSize, page}` (0-indexed page)                                        |
| `POST /api/app/view-app-summary`           | Per-entity procurement segments, paginated                                                                                                           | `{page, searchTerm: json.dumps({appDetailsId, isSearch:false}), pageSize, appDetailsId, isSearch:false}` (1-indexed page) |
| `POST /api/app/get-appdetail-summary/{id}` | Ministry name + APP metadata                                                                                                                         | none                                                                                                                        |
| `POST /api/app/get-app-details/{id}`       | **AVOID** — leaks real user PII (login ID, email, phone, bcrypt hash of the officer who created the plan). Not used anywhere in the pipeline. |                                                                                                                             |

Every listing/segment response includes a `totalCount`/`totalcount` field,
so pagination loops are exact (no trial-and-error probing).

`view-app-summary`'s raw segment row also includes an `unspscsegment` field
(e.g. `"44000000"`) that isn't currently extracted by `fetch_entity_segments()`
— confirmed present via direct inspection, not yet wired in. This is the
join key the Item Details feature (below) uses to map items back to segments.

### Known site quirks

- A Deskpro live-chat widget (`support.egpkenya.go.ke/deskpro-messenger/...`)
  keeps a persistent connection open, which means `wait_until="networkidle"`
  **never resolves**. Always use `domcontentloaded` + an explicit
  `wait_for_selector`, and block the widget's requests via
  `page.route("**/deskpro-messenger/**", lambda route: route.abort())`.
- **The hash-less `public-view-app/{id}/{mode}/` URL does not work for
  navigation.** Previously assumed merely "not clickable" / unneeded; now
  confirmed it actively **silently redirects to the site homepage** when
  navigated to directly, with no error thrown. It's fine to keep storing
  it in `latest_entities` as a human-readable reference, but any future
  automation that needs to actually load an entity's real page (not just
  call its API by `appdetailid`) must reach it by clicking through the
  search UI — see "Item Details / OPEN-AGPO feature" below. The real
  hashed URL (`.../{id}/{mode}/{hash}`) is generated client-side by the
  Angular router only at click time; it cannot be constructed from any
  data the API returns.

## File responsibilities

- **`main.py`** — CLI menu / orchestrator. Steps 1a→1b→2→3→4, full pipeline,
  skip-addresses variant, or standalone reconcile (option `[8]`).
- **`scrape_entities.py`** — light scrape (`scrape_to_latest_entities`),
  deep scrape (`run_deep_scrape`), entity workbook generation
  (`create_entity_template`), blacklist filtering (`is_blacklisted`), and
  `reconcile_budget_totals_from_workbooks` (now wired into `main.py`).
  **Item Details / OPEN-AGPO functions not yet added — see below.**
- **`budget_tracker.py`** — re-indexes `Budget Totals` against
  `latest_entities`, preserves computed Total Budget, and pulls
  collaborator-edited Status/Prequal from the live Google Sheet before
  recomputing (see "Google Sheets round trip" below).
- **`address_tracker.py`** — Bing-based contact enrichment, filtered
  against the blacklist and run concurrently across multiple browser
  instances (see "Address tracker rewrite" below).
- **`sync_gsuite.py`** — pushes `Budget Totals` (5 cols), `blacklisted`
  (2 cols, same spreadsheet as Budget Totals) and `addresses` (9 cols) to
  their respective Google Apps Script Web App endpoints, with a
  `sheetName` field in the payload so `code.gs` routes to the right tab.
- **`code.gs`** (Apps Script project `egp-scraper-updater`) — receives the
  POST payload, routes by `sheetName`, clears + rewrites the target sheet,
  reapplies formatting/status highlighting for `Budget Totals` only.
  `doGet` is a proper top-level function, accepting an optional `?sheet=`
  param.

## Data schema

### `latest_entities` (6 columns)

`Sr. No. | Financial Year | Procuring Entity | APP Number | APP URL | APP Detail ID`

`APP Detail ID` is the numeric `appdetailid` from the API. `APP Number`
(e.g. `TENP/864/APP/2026-27/3`) turns out to double as the search key for
the Item Details feature's click-through navigation — see below.

### `Budget Totals` (5 columns)

`Sr. No. | Procuring Entity | Total Budget | Status | Prequalification done`

- **Total Budget**: formatted as `"KES" #,##0.00` everywhere it's written.
- **Status**: `Pending` → `Partial` (amber, set automatically by deep scrape)
  → `Done` (green, reserved for once OPEN/AGPO + Q1–Q4 exist — nothing
  currently sets this automatically; still true even once the Item Details
  feature is built, since a segment can legitimately aggregate to all
  blank/`0` if it has no qualifying RFQ/Open Tender items — that's not the
  same thing as "not yet computed").
- **Prequalification done**: manual, round-trips from the live Google
  Sheet, never overwritten locally.

### `blacklisted` (2 columns)

`Sr. No. | Procuring Entity` — visibility sheet for `is_blacklisted()`
exclusions. Currently one-way (sourced from the hardcoded set); inverting
this is still a planned follow-up, not yet built.

### Entity workbooks (`./entity_plans/{sr_no} {entity}.xlsx`)

`Sr. No. | Segment | Total Cost | OPEN/AGPO | Q1 | Q2 | Q3 | Q4`

Only Sr No/Segment/Total Cost are populated by the current deep scrape.
**Populating OPEN/AGPO and Q1–Q4 is actively in progress — see below for
current status.** They remain blank strings (not `0.0`) until this ships.

### `addresses` (9 columns)

`Sr. No. | Procuring Entity | Physical Address | Town/Area | County | Phone | Email | ASSIGNED | VISIT`

Blacklisted entities excluded before being added to the search list at all.

## Item Details / OPEN-AGPO feature (in progress)

Full spec lives in `item_details_feature_spec.md`. Summary of where things
actually stand:

**The business rule is confirmed and validated at full scale** (3,246
items / 55 segments from one entity, zero edge cases broken it): filter
items to `Request for Quotation` or `Open Tender` procurement methods,
tag each by method + AGPO reservation (`OPEN` / `RFQ` / `RFQ WOMEN` /
`RFQ YOUTH` / `RFQ PWD`), union per quarter per segment, `0` if empty.

**The data source turned out simpler than planned.** Originally assumed
this would need a reverse-engineered paginated API call per segment. It
doesn't — one "Export to Excel" click on an entity's Item Details tab
downloads **every item across every segment for that entity in one file**,
and each item's segment can be derived for free from its UNSPSC code
(first 2 digits = segment, matches the `unspscsegment` field already
sitting in `view-app-summary`). So this is one browser download per
entity, not one API call per segment.

**The navigation chain to reach that download is now solved and tested
headless, for one entity (Eldoret)**: use `APP Number` to search via a
collapsed accordion panel (`#first-toggle` to expand, `input[formcontrolname='appNumber']` to search) → click the resulting row → land on the real
hashed URL → Item Details tab → Export to Excel → catch the download →
parse. Zero manual steps required once built correctly.

**Not yet done:**

- Only tested against one entity — needs at least 2–3 more to trust it
  generalizes (missing segments, missing buttons, different page timing
  are all unverified failure modes).
- No error handling for any of those edge cases yet — currently throws
  and halts rather than skipping/logging.
- No rate limiting between entities — this is heavier browser traffic
  per entity than the existing pure-API `fetch_entity_segments()` call,
  and the site's bot protection is a known real constraint.
- None of this is wired into `scrape_entities.py` yet — everything so far
  is standalone `test.py` scripts run manually.

## Google Sheets round trip

Unchanged. Design: the Google Sheet is authoritative for **Status +
Prequalification done only**.

1. `pull_status_from_gsheet()` — reads columns D/E from the live sheet,
   before any local recompute.
2. Local recompute proceeds as normal.
3. `sync_gsuite.py` pushes the full sheet back up after 1+2.

## Address tracker rewrite (performance)

Unchanged from previous notes. 108 non-blacklisted entities in ~88
seconds via per-thread Playwright lifecycles + batched saves, down from
45+ minutes sequential. Playwright's sync API is pinned to the OS thread
that created it — never share a Playwright object across threads, even
via a queue.

## Deferred / not yet built

- Item Details / OPEN-AGPO feature integration into `scrape_entities.py`
  (see above — mechanism proven, not wired in).
- Multi-entity validation of the Item Details navigation chain.
- Rate limiting for the Item Details download step.
- Once OPEN/AGPO + Q1–Q4 are actually populated: decide the exact rule
  for promoting `Partial → Done`, being careful not to conflate "computed
  and found nothing" with "not yet computed."
- Inverting the blacklist sheet so `is_blacklisted()` reads from
  `blacklisted` instead of the hardcoded set.

## Testing checklist for the next full run

1. Run 1a → confirm `latest_entities` has `APP Detail ID` populated for
   every row, no blacklisted entities present, `blacklisted` sheet
   populated correctly.
2. Run 1b → confirm one `.xlsx` per entity, Sr No/Segment/Total Cost
   filled, OPEN/AGPO–Q4 still blank (until the new feature ships), TOTAL
   row sums correctly, `Budget Totals` shows `Partial` with correct KES
   totals.
3. Run Step 2 → confirm Status/Prequal pulled correctly, no row
   misalignment.
4. Run Step 3 → confirm it completes quickly, no `greenlet.error`,
   blacklisted entities never appear.
5. Run Step 4 → confirm all three Sheets tabs receive matching data,
   collaborator edits survive the round trip.
6. Try option `[8]` (reconcile) standalone.
7. **Once Item Details feature ships**: confirm OPEN/AGPO and Q1–Q4 are
   populated with the correct tag strings, `'0'` appears (not blank) for
   segments with no qualifying items, and spot-check at least one segment
   against its known-correct hand-computed result (segment `44000000`,
   Eldoret, should read `RFQ, RFQ WOMEN, RFQ YOUTH` across all four
   quarters).
