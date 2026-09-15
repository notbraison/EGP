# Dev Notes

_Last updated: 2026-09-10_

venv/Scripts/activate

py main.py

Budget - [docs.google.com/spreadsheets/d/18PZby5HOyY1EWue6rIZL6kHKv0h1mN-2oa3EV-aF9kw/edit?usp=sharing](https://docs.google.com/spreadsheets/d/18PZby5HOyY1EWue6rIZL6kHKv0h1mN-2oa3EV-aF9kw/edit?usp=sharing)

Addresses - [docs.google.com/spreadsheets/d/1tCpLpMCHzEGwZzbk63q2j-KmXoG9SCUIoRPWl22MiIc/edit?usp=sharing](https://docs.google.com/spreadsheets/d/1tCpLpMCHzEGwZzbk63q2j-KmXoG9SCUIoRPWl22MiIc/edit?usp=sharing)


## What this project does

Automated pipeline that pulls Kenya e-GP public procurement plan data,
builds a local master workbook (`Procuring_Entities.xlsx`), generates one
formatted workbook per procuring entity (`./entity_plans/`), and syncs the
budget + address-tracking sheets to Google Sheets via Apps Script Web Apps.

## Architecture (API-based for listing/segments, browser-automated for Item Details)

Entity listing and segment data go through the portal's internal JSON API,
authenticated via a real Playwright session's XSRF cookie — see previous
notes for the full API endpoint table and the reasoning for why DOM
scraping/`onclick` parsing never worked on this Angular SPA.

 **New this round** : the OPEN/AGPO + Q1–Q4 feature (section below) is a
deliberate, scoped exception to "no clicking" — it drives real browser
navigation, because the data it needs has no equivalent JSON API found
(or expected to exist). Everything else stays API-only.

## File responsibilities

* **`main.py`** — CLI menu. Steps 1a→1b→2→3→4, full pipeline, skip-addresses
  variant, reconcile (`[8]`). **Item Details enrichment is not yet wired
  into the menu** — currently invoked directly in code for testing.
* **`scrape_entities.py`** — light scrape, deep scrape, entity workbook
  generation, blacklist filtering, reconcile, **and the new Item Details /
  OPEN-AGPO feature** (see below).
* **`budget_tracker.py`** ,  **`address_tracker.py`** ,  **`sync_gsuite.py`** ,
  **`code.gs`** — unchanged from previous notes.

## Item Details / OPEN-AGPO feature — status: built, tested on one entity

### Business rule (confirmed, validated at scale on one entity)

For each item in an entity's Item Details export:

1. Filter to `Procurement Method` in `{"Request for Quotation", "Open Tender"}` — drop everything else.
2. Tag: `Open Tender` → always `OPEN`. `Request for Quotation` + blank `Reservation Group` → `RFQ`. `Request for Quotation` + `Reservation Group = "AGPO"` → `RFQ` plus whichever of Women/Youth/PWD has nonzero quantity (exactly one has ever been nonzero, across every row checked so far).
3. Per quarter, per segment: union of tags among items with nonzero quantity that quarter, fixed order `OPEN, RFQ, RFQ WOMEN, RFQ YOUTH, RFQ PWD`, comma-separated, or `'0'` if empty.
4. `OPEN/AGPO` summary column = same union across all four quarters.

Validated against 3,246 items / 55 segments from one entity (Eldoret
National Polytechnic) — zero broken edge cases in that run. **Not**
exhaustively verified across other entities.

### Data source — simpler than originally planned

Originally assumed this needed a reverse-engineered paginated per-segment
API call (`fetch_segment_items()`). Dropped that plan entirely: an
entity's Item Details tab has an "Export to Excel" button that downloads
**every item across every segment for that entity in one file** —
confirmed via a real download (3,246 rows, matching the entity's full
segment count). Each item's segment is derived for free from its
`UNSPSC/Item Code` (first 2 digits = segment code, by UNSPSC definition),
matching the previously-unused `unspscsegment` field `fetch_entity_segments()`
now also captures as `segment_code`.

**Net effect: one browser download per entity, not one API call per
segment.** No `capture_item_details_api.py` / network-capture work needed.

### Navigation chain — solved, headless, one entity, including a real bug found

1. **Never use the hash-less `latest_entities` APP URL** — confirmed it
   silently redirects to the homepage. No exception thrown. Dangerous
   failure mode inside a loop.
2. **Real hashed URL only obtainable via the portal's own search UI** :
   expand a collapsed accordion (`#first-toggle`, verified via
   `aria-expanded` flipping to `"true"`), fill
   `input[formcontrolname='appNumber']` with the entity's APP Number
   (already in `latest_entities`), click Search, click the resulting row.
3. **Bug found and fixed** : the accordion toggle's own visible text is
   also "Search." A selector of `button:has-text('Search')` matched
   *both* the accordion toggle and the real submit button, and `.first`
   silently grabbed the wrong one — no error, no crash, just a table that
   quietly kept showing its default unfiltered first page no matter what
   was searched. This produced a  **false-positive test** : an earlier
   check against Eldoret "passed" only because Eldoret happens to be row
   1 of the default unfiltered listing, not because search actually
   filtered anything.
   Diagnosed properly, not guessed: ran the same script against two
   *different* APP Numbers and printed the full row list each time — both
   searches returned the identical unfiltered 10-row listing, proving the
   filter never fired at all (rather than firing and coincidentally
   matching wrong).
   Root cause confirmed via DevTools inspection of the real submit
   button vs. the toggle: the real button has `type="submit"`, the
   accordion toggle does not.
   **Fix** : `button:has-text('Search')` → `button[type='submit']:has-text('Search')`.
   Verified: searching `AWWDA/704/APP/2026-27/2` now returns exactly 1
   row (Athi Water Works Development Agency), not the default 10.
4. From the real URL: click "Item Details" tab → click "Export to Excel"
   → catch download via `page.expect_download()` → parse → delete temp
   file.

### Functions added to `scrape_entities.py`

* `derive_segment_code(item_code)` — UNSPSC first-2-digits → segment code.
* `tag_for_item(item)` — per-item categorization.
* `categorize_and_aggregate_items(items)` — pure, no API/browser
  dependency, independently testable.
* `parse_item_details_export(filepath)` — reads a downloaded `.xlsx` into
  raw item dicts.
* `search_and_open_entity(page, app_number, entity_name)` — the fixed
  navigation chain. Raises clear `RuntimeError`s (accordion missing,
  search input missing, submit button missing, no matching row, row has
  no link) instead of failing silently.
* `download_and_parse_item_details(page)` — tab click → export → download
  → parse → delete.
* `entity_needs_item_details(filepath)` — true if workbook missing or any
  segment row has a blank OPEN/AGPO cell.
* `run_item_details_enrichment(file_path, overwrite, delay_seconds, limit)`
  — orchestrating loop. Per entity: fetch segments via API, run the
  navigation+download+parse chain, merge by `segment_code`, regenerate
  the workbook via `create_entity_template(..., item_data=...)`, update
  Budget Totals status to `"Partial"`. `delay_seconds` (default 2.0)
  between entities — bot-protection caution, not yet stress-tested.
  `limit` param for small test batches. **Deliberately separate from
  `run_deep_scrape()`** — that function's skip logic keys off "has a
  Total Budget," unrelated to OPEN/AGPO population, so folding this in
  would mean it almost never runs once most entities already have a
  budget.

`create_entity_template()` updated to accept an optional `item_data` dict
(`{segment_code: {open_agpo, q1, q2, q3, q4}}`); falls back to blank
strings for any segment not found in it — unchanged behavior from before
this feature existed, so nothing breaks for entities not yet enriched.

### Site quirk discovered this round: don't block Deskpro beyond initial load

Blocking the Deskpro chat widget's requests (`page.route(...).abort()`)
is necessary during `get_authenticated_context()`'s one-time initial page
load (otherwise `networkidle` never resolves — see earlier notes). But
doing the same on the Item Details page caused a real problem: an
uncaught `TypeError` inside the widget's own bundle propagates through
Angular's shared `zone.js` task queue and can silently disrupt  *other* ,
unrelated in-flight requests on the same page. Confirmed: this caused the
Item Details table to show "No Records Found" with zero errors surfaced
anywhere in Playwright — looked exactly like a dead endpoint, wasn't.
Fix: only block Deskpro during the initial authenticated-session load;
leave it unblocked on any page where other requests need to complete
reliably.

### Not yet done (open items before this is production-ready)

1. Tested end-to-end on exactly one entity (Eldoret). Athi Water Works
   only had its raw export inspected manually — full automated chain not
   run against it yet.
2. Generalization unverified: entities with zero segments, a missing
   Export button, different page timing, etc.
3. No rate-limiting stress test beyond the default 2s delay — this site
   has known bot protection and this feature's per-entity browser+download
   footprint is heavier than the pure-API calls elsewhere in the pipeline.
4. Not wired into `main.py`'s menu.
5. Cross-entity validation still pending for: "exactly one of
   Women/Youth/PWD ever nonzero" and "no non-RFQ/Open-Tender method ever
   has nonzero W/Y/PWD" — both currently resting on one entity's data.
6. General lesson worth carrying forward: **ambiguous text-based
   selectors are a real, repeated risk on this site.** The Search button
   bug is the second time a `text=`/`has-text()` match has needed
   hardening (the first, minor one, was `text=Search` also matching a
   "...RESEARCH" div during early testing, caught before it mattered).
   Prefer attribute-based selectors (`type=`, `id`, `aria-*`) over
   text-only matching wherever more than one element could plausibly
   share visible text.

## Resolved issues (cumulative, previous rounds)

1–8: see earlier notes (missing save call, nested `doGet`, off-by-one
formatting, unwired reconcile tool, address tracker performance +
concurrency crash + dead code cleanup) — all fixed and verified.

9. ~Item Details search returning unfiltered results with no error~ —
   root-caused to an ambiguous `has-text('Search')` selector matching the
   accordion toggle instead of the real submit button; fixed via
   `type='submit']` qualifier; verified against a live search.

## Testing checklist for the next full run

Unchanged items 1–6 from previous notes, plus, for the new feature:

7. Run `run_item_details_enrichment(limit=1)` against a single entity you
   haven't tested before (not Eldoret) — confirm the workbook gets real
   OPEN/AGPO + Q1–Q4 values, not blanks, and that the Budget Totals status
   still reads correctly.
8. Run it with `limit=5` or so against a small batch — watch for any
   entity that throws a `RuntimeError` from `search_and_open_entity()`
   (missing accordion/input/button/row/link) and note which step failed;
   that's the next generalization gap to handle.
9. Once a handful of entities pass cleanly, decide on: wiring this into
   `main.py`'s menu, and whether `delay_seconds` needs to be longer for a
   full unattended run over the whole entity list.
