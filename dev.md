# Dev Notes

_Last updated: 2026-09-09_

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

### Known site quirks

- A Deskpro live-chat widget (`support.egpkenya.go.ke/deskpro-messenger/...`)
  keeps a persistent connection open, which means `wait_until="networkidle"`
  **never resolves**. Always use `domcontentloaded` + an explicit
  `wait_for_selector`, and block the widget's requests via
  `page.route("**/deskpro-messenger/**", lambda route: route.abort())`.
- The hash segment of `public-view-app/{id}/{mode}/{hash}` URLs is not
  needed for anything — every useful endpoint is keyed by `appdetailid`
  alone. The pipeline stores a hash-less reference URL in `latest_entities`
  for human cross-checking only; it isn't clickable.

## File responsibilities

- **`main.py`** — CLI menu / orchestrator. Steps 1a→1b→2→3→4, or full
  pipeline, or skip-addresses variant.
- **`scrape_entities.py`** — light scrape (`scrape_to_latest_entities`),
  deep scrape (`run_deep_scrape`), entity workbook generation
  (`create_entity_template`), blacklist filtering, and the new
  `reconcile_budget_totals_from_workbooks` auditor (not yet wired into
  `main.py`).
- **`budget_tracker.py`** — re-indexes `Budget Totals` against
  `latest_entities`, preserves computed Total Budget, and now also pulls
  collaborator-edited Status/Prequal from the live Google Sheet before
  recomputing (see "Google Sheets round trip" below).
- Google sheets to this above - [docs.google.com/spreadsheets/d/18PZby5HOyY1EWue6rIZL6kHKv0h1mN-2oa3EV-aF9kw/edit?usp=sharing](https://docs.google.com/spreadsheets/d/18PZby5HOyY1EWue6rIZL6kHKv0h1mN-2oa3EV-aF9kw/edit?usp=sharing)
- **`address_tracker.py`** — unchanged from the original design. Bing-based
  contact enrichment via Playwright, independent of the eGP API rewrite.
- Google sheets to this above - [docs.google.com/spreadsheets/d/1tCpLpMCHzEGwZzbk63q2j-KmXoG9SCUIoRPWl22MiIc/edit?usp=sharing](https://docs.google.com/spreadsheets/d/1tCpLpMCHzEGwZzbk63q2j-KmXoG9SCUIoRPWl22MiIc/edit?usp=sharing)
- **`sync_gsuite.py`** — pushes `Budget Totals` (5 cols) and `addresses`
  (9 cols) to their respective Google Apps Script Web App endpoints.
- **`code.gs`** (Apps Script project `egp-scraper-updater`) — receives the
  POST payload, clears + rewrites the sheet, reapplies formatting/status
  highlighting. Also intended to serve `doGet` for the pull-down direction
  (see bugs below).

## Data schema

### `latest_entities` (6 columns — added APP Detail ID)

`Sr. No. | Financial Year | Procuring Entity | APP Number | APP URL | APP Detail ID`

`APP Detail ID` is the numeric `appdetailid` from the API — this is what
lets deep scrape fetch segments directly, with no name-matching or
re-navigating the listing page required.

### `Budget Totals` (5 columns — reordered)

`Sr. No. | Procuring Entity | Total Budget | Status | Prequalification done`

- **Total Budget**: formatted as `"KES" #,##0.00` everywhere it's written
  (`scrape_to_latest_entities`, `update_master_budget_totals`,
  `budget_tracker.py`, `reconcile_budget_totals_from_workbooks`).
- **Status**: `Pending` (no formatting) → `Partial` (amber `FFF2CC` /
  `9C6500`, bold) → `Done` (green `E2EFDA` / `375623`, bold). `Done` is
  reserved for when OPEN/AGPO + Q1–Q4 exist — nothing in the current
  pipeline sets it automatically. Deep scrape always writes `Partial`.
- **Prequalification done**: intentionally left blank by the pipeline —
  filled manually by your collaborator on the Google Sheet side, then
  pulled back down locally by `pull_status_from_gsheet()`.

### Entity workbooks (`./entity_plans/{sr_no} {entity}.xlsx`)

`Sr. No. | Segment | Total Cost | OPEN/AGPO | Q1 | Q2 | Q3 | Q4`

Only Sr No/Segment/Total Cost are populated by the current deep scrape.
OPEN/AGPO and Q1–Q4 are written as blank strings (not `0.0`) so it's
visually obvious in Excel that they're unpopulated rather than genuinely
zero. This is intentional — second implementation territory (see below).

### `addresses` (9 columns — unchanged)

`Sr. No. | Procuring Entity | Physical Address | Town/Area | County | Phone | Email | ASSIGNED | VISIT`

## Google Sheets round trip (new)

Design: the Google Sheet is authoritative for **Status + Prequalification
done only**. Your collaborator edits those two columns directly in
Sheets. The local pipeline must never clobber that.

Flow, in order:

1. `budget_tracker.py` → `pull_status_from_gsheet()` — `GET`s the live
   sheet, reads columns D/E (Status/Prequal) per entity, writes them into
   the local workbook. Runs **before** any local recompute.
2. Local recompute proceeds as normal (Total Budget from deep scrape /
   reconcile, Status defaults for any entity the collaborator hasn't
   touched yet).
3. `sync_gsuite.py` pushes the full sheet back up, **after** step 1+2 — so
   the collaborator's edits round-trip back to Sheets unchanged, alongside
   the freshly computed totals.

This depends on `code.gs` correctly serving `doGet` for step 1 to work at
all — see bugs below.

## Deferred to "second implementation"

- Populating OPEN/AGPO and Q1–Q4 in entity workbooks — likely from the
  portal's "Item Details" tab (seen in the UI, not yet reverse-engineered
  via network capture the way APP Summary was).
