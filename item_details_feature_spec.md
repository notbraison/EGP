# Feature Spec: Populating OPEN/AGPO and Q1–Q4 in Entity Workbooks

_Status: business rule confirmed AND validated at scale. Navigation chain
proven headless end-to-end for one entity. Not yet wired into
`scrape_entities.py`. Not yet tested across multiple entities._

## 1. What this feature is

Every entity workbook in `./entity_plans/` currently has 8 columns:

```
Sr. No. | Segment | Total Cost | OPEN/AGPO | Q1 | Q2 | Q3 | Q4
```

Only the first three are populated today, by `create_entity_template()` in
`scrape_entities.py`, using data from the `view-app-summary` API endpoint
(segment name + total cost only). `OPEN/AGPO` and `Q1`–`Q4` are written as
blank strings intentionally — this feature is what fills them in.

## 2. The business rule (confirmed AND validated at scale)

### 2.1 Per-item categorization

For each row in a segment's Item Details table:

1. **Filter**: only rows where `Procurement Method` is one of:

   - `"Request for Quotation"` → tender type `RFQ`
   - `"Open Tender"` → tender type `OPEN`
     Any other procurement method (Direct Procurement, Restricted tender,
     KEMSA, etc.) is dropped entirely.
2. **Tag**, based on tender type + `Preference & Reservation Group`:

   - Tender type `OPEN` → tag is always just `OPEN`. Confirmed directly
     from a real Open Tender row (Athi Water Works, Chlorine Cl item):
     `Procurement Method = "Open Tender"`, `Reservation Group` blank,
     `Women/Youth/PWD = "-"`. Open Tender items never carry an AGPO
     reservation in the source data.
   - Tender type `RFQ` and `Reservation Group` blank → tag is `RFQ`.
   - Tender type `RFQ` and `Reservation Group = "AGPO"` → tag is `RFQ`
     plus whichever of `Women`/`Youth`/`PWD` has a nonzero quantity.

   Full tag set: `OPEN`, `RFQ`, `RFQ WOMEN`, `RFQ YOUTH`, `RFQ PWD`.
3. **Per-quarter presence**: an item is "in" quarter `Qn` if its `Qn`
   quantity column is nonzero.

### 2.2 Segment-level aggregation

Per quarter, per segment: union of item-tags over every qualifying item
with nonzero quantity that quarter, written in fixed order
`OPEN, RFQ, RFQ WOMEN, RFQ YOUTH, RFQ PWD`, comma-separated. Empty set →
write `0`.

The `OPEN/AGPO` summary column is the same union taken across all four
quarters combined for that segment.

### 2.3 Validated against

- Item Details export, "Office Equipment & Accessories & Supplies" segment,
  Eldoret National Polytechnic (91 rows, all RFQ).
- Item Details export, "Furniture & Furnishings" segment, same entity
  (79 rows, all RFQ) — confirmed AGPO items can be present in a single
  quarter only.
- Item Details export, one Open Tender row, Athi Water Works Development
  Agency — confirmed Open Tender's blank reservation group, matched
  exactly against that entity's own hand-filled workbook.
- **NEW — full-scale validation**: ran `categorize_and_aggregate_items()`
  against Eldoret's *complete* Item Details export — 3,246 items across
  55 distinct segments (not just the 91-item sample). Segment `44000000`
  reproduced exactly the same result as the hand-checked sample
  (`RFQ, RFQ WOMEN, RFQ YOUTH` across all four quarters). Zero items,
  across all 3,246, ever had more than one of Women/Youth/PWD nonzero
  simultaneously — the "pick one" assumption in `tag_for_item()` never
  had to make a real choice. Segments with genuinely zero qualifying
  items (e.g. `93000000`, `94000000`) correctly aggregated to `'0'`
  rather than erroring or defaulting to something misleading.
- ~18 of your own hand-filled entity workbooks (files 5–9, 13–24, 26) —
  confirmed the general shape even though none individually used fully
  consistent formatting.

**Still not exhaustively verified**: whether a procurement method beyond
RFQ/Open Tender ever shows a nonzero Women/Youth/PWD value, and whether
the "exactly one of Women/Youth/PWD" pattern holds for every entity, not
just Eldoret. One entity's full data is a strong signal, not a proof.

## 3. Data source: resolved, and simpler than expected

**The original plan (reverse-engineer a paginated `fetch_segment_items()`
API call, one call per segment) is dropped.** What actually works, found
and confirmed this session:

- Each segment's public-view-app page has an **"Item Details" tab** with
  an **"Export to Excel" button**.
- Clicking Export to Excel downloads **every item across every segment
  for that entity in one file** — not scoped to the currently-viewed
  segment. Confirmed: one download for Eldoret returned 3,246 rows / 55
  segments in a single `.xlsx`, matching the entity's full segment count
  from `view-app-summary`.
- **No segment ID needs to be passed anywhere.** Each item's own
  `UNSPSC/Item Code` column already encodes its segment: UNSPSC is a
  standardized hierarchy where the first 2 digits *are* the segment code
  by definition (e.g. item code `10101510` → segment `10000000`). This
  matches the `unspscsegment` field already present (and previously
  unused) in `view-app-summary`'s raw JSON response for each segment
  (`{"unspscsegment": "44000000", "description": "...", ...}`).
- **This means one browser download per entity, not one per segment** —
  a huge simplification versus the original spec's "50+ extra API calls
  per entity" concern. `capture_item_details_api.py` (network-capture
  approach) is no longer needed; this is a browser-automation problem
  (navigate → click → catch download), not an API reverse-engineering one.

### 3.1 The navigation chain (fully solved, headless, one entity proven)

This took several rounds to get right — worth recording exactly what
failed and why, since the failure modes are non-obvious:

1. **The hash-less URL doesn't work.** `latest_entities`' stored `APP URL`
   (`https://egpkenya.go.ke/public-view-app/{appdetailid}/1/`, no hash) —
   confirmed by the dev notes as "not clickable" — was tested directly:
   it **silently redirects to the site homepage**, no error thrown. This
   is a dangerous failure mode for automation since it produces no
   exception, just an empty/wrong result.
2. **The real hashed URL** (`.../28506/1/5116DFB0510F...`) is only
   obtainable by clicking the entity's link from the **search listing** —
   the hash is generated client-side by the Angular router at click time,
   it's not derivable from any data the API already returns.
3. **The search UI itself is behind a collapsed accordion**, not a
   visible input by default. Found via browser DevTools inspection (not
   guessing):
   - Toggle: `<button id="first-toggle" ngbaccordionbutton aria-expanded="false">` — expand it, and confirm via `aria-expanded` flipping to `"true"` rather than a blind `wait_for_timeout` (more reliable, throws a clear error if the click didn't register).
   - Search field: `<input formcontrolname="appNumber">` — filters by
     APP Number specifically (already stored in `latest_entities`, no
     extra data needed).
   - Submit: a `<button>` containing the text "Search" (careful: Playwright's
     `text=` locator does substring matching — "Search" also matched a
     `<div>` reading "...RESEARCH" during testing. Not currently a bug
     since `.first` was used correctly, but worth hardening later with a
     more specific selector if this expands to more pages).
4. **From the real hashed URL**: click "Item Details" tab, click "Export
   to Excel", catch the download via `page.expect_download()`, parse.

**Full chain tested headless, zero manual intervention, one entity
(Eldoret)**: search by APP Number → expand accordion → click result row
→ land on real URL → Item Details → Export → download → parse. Result:
3,246 items, 55 segments, matches the manual-click baseline from the
file-based test exactly.

## 4. How this maps onto the existing pipeline

### 4.1 Functions to build (none exist in `scrape_entities.py` yet)

- **`search_and_open_entity(page, app_number, entity_name)`** — the
  navigation chain from §3.1, steps 1–3. Tested standalone; not yet in
  `scrape_entities.py`.
- **`download_and_parse_item_details(page)`** — Item Details tab → Export
  → download → `parse_item_details_export()`. Tested standalone (§3.1
  step 4); not yet in `scrape_entities.py`.
- **`parse_item_details_export(filepath)`** — reads the downloaded
  `.xlsx` into raw item dicts, deriving `segment_code` from the UNSPSC
  item code. Tested and working against real downloaded files.
- **`categorize_and_aggregate_items(items)`** — pure function implementing
  §2. Tested and working at full scale (3,246 items). No API/browser
  dependency — this part is done and just needs to be dropped into the
  module as-is.
- **`tag_for_item(item)`** — the per-item categorization helper used by
  the above. Tested.

### 4.2 Where it plugs into existing functions

- **`run_deep_scrape()`** — currently calls `fetch_entity_segments()` then
  `create_entity_template()`, using pure API calls, no browser navigation
  at all. This feature requires an *additional* browser-driven step per
  entity: after getting segments via the existing API call, also run
  `search_and_open_entity()` + `download_and_parse_item_details()` once
  per entity (not once per segment — confirmed this is entity-scoped, see
  §3), then group the resulting items by `segment_code` and merge each
  group's `categorize_and_aggregate_items()` result into the corresponding
  segment dict before it reaches `create_entity_template()`.
- **`create_entity_template()`** — currently writes `"", "", "", "", ""`
  for columns D–H. Change to write the four aggregated strings instead.
- **`Status` column semantics** — unchanged reasoning from before: a
  segment aggregating to all-`0` is a legitimate "no qualifying items"
  result, not an "unfinished" one. Don't conflate the two when deciding
  any future Partial→Done promotion logic.

### 4.3 What this feature does not touch

Unchanged from before: `budget_tracker.py`, `sync_gsuite.py`/`code.gs`,
`address_tracker.py` are all unaffected.

## 5. Open items before this is production-ready

Rewritten to reflect what's actually been tested vs. assumed:

1. **Only tested on one entity (Eldoret).** Every step — search, accordion,
   Export button, download — worked for this one entity's page. Whether
   this generalizes (different entities might have zero segments, a
   missing Export button, a slower-loading tab, etc.) is unverified.
   Athi Water Works was intended as a second test case but never actually
   completed the full navigation chain (only its raw item-detail export
   was inspected manually, separately, earlier).
2. **No error handling for entity-specific edge cases** — a missing
   accordion, missing search input, missing Export button, or zero search
   results would currently throw and halt rather than skip/log and
   continue. Needs handling before running unattended over hundreds of
   entities.
3. **No rate limiting / delay between entities.** The existing dev notes
   already flag this site's bot protection as a real constraint (plain
   `requests` calls get blocked). This feature's per-entity browser
   navigation + download is a heavier footprint than the existing
   pure-API calls in `fetch_entity_segments()` — worth deliberately
   throttling before looping over the full entity list, not just hoping
   it's fine.
4. **Not yet wired into `scrape_entities.py` at all.** Everything so far
   lives in standalone `test.py` scripts run manually via Notepad/`py test.py`. None of it has been integrated into `run_deep_scrape()` or
   `create_entity_template()` yet.
5. **The `text=Item Details"` / `text=Export to Excel"` selectors** work
   but are substring-text matches — the same class of fragility that
   produced one confirmed false-positive already (`text=Search` matching
   a "RESEARCH" div). Not yet hardened to more specific selectors (e.g.
   role/aria-based) the way the search accordion and input were.
6. **Whether a procurement method beyond RFQ/Open Tender ever carries a
   nonzero Women/Youth/PWD value**, and whether "exactly one of
   Women/Youth/PWD nonzero" holds across entities generally — both still
   resting on one entity's data (Eldoret), not a cross-entity sampl

# Feature Spec: Populating OPEN/AGPO and Q1–Q4 in Entity Workbooks

_Status: business rule confirmed, endpoint not yet reverse-engineered. Not implemented._

## 1. What this feature is

Every entity workbook in `./entity_plans/` currently has 8 columns:

```
Sr. No. | Segment | Total Cost | OPEN/AGPO | Q1 | Q2 | Q3 | Q4
```

Only the first three are populated today, by `create_entity_template()` in
`scrape_entities.py`, using data from the `view-app-summary` API endpoint
(segment name + total cost only). `OPEN/AGPO` and `Q1`–`Q4` are written as
blank strings intentionally — this feature is what fills them in.

The source data for these four columns lives one level deeper than anything
the pipeline currently fetches: the **"Item Details" tab** on each segment's
public-view-app page (see Image 2/3/4 from earlier — click the eye/Action
icon next to a segment row in APP Summary to get there). This is a paginated
table of individual line items (91 items for one sample segment, ~10/page),
not currently reachable via any endpoint the scraper calls.

## 2. The business rule (confirmed)

This took several rounds of cross-checking against real item-level exports
(docs 3 and 6) and ~18 of your own hand-filled entity workbooks to pin down,
because your own historical files used at least four different manual
conventions (comma vs slash-delimited, "open" meaning two different things
depending on the file, procurement method sometimes folded into the tag and
sometimes not). The rule below is the one that's actually evidence-backed
against raw data, not just the most common historical convention.

### 2.1 Per-item categorization

For each row in a segment's Item Details table:

1. **Filter**: only rows where `Procurement Method` is one of:

   - `"Request for Quotation"` → tender type `RFQ`
   - `"Open Tender"` → tender type `OPEN`
     Any other procurement method (Direct Procurement, Restricted tender,
     KEMSA, etc.) is dropped entirely — confirmed by your own tendering notes:
     *"Our focus (as a supplier): Open tenders and RFQs."*
2. **Tag**, based on tender type + `Preference & Reservation Group`:

   - Tender type `OPEN` → tag is always just `OPEN`.
     Confirmed directly from a real Open Tender row (Athi Water Works,
     Chlorine Cl item): `Procurement Method = "Open Tender"`,
     `Reservation Group` blank, `Women/Youth/PWD = "-"`. Open Tender items
     never carry an AGPO reservation in the source data — it's not a rule
     we're imposing, it's how the portal's own data is structured.
   - Tender type `RFQ` and `Reservation Group` blank → tag is `RFQ`.
   - Tender type `RFQ` and `Reservation Group = "AGPO"` → tag is `RFQ` plus
     whichever of `Women` / `Youth` / `PWD` has a nonzero quantity for that
     row (exactly one of the three was ever nonzero in every sample seen).

   So the full set of possible per-item tags is:
   `OPEN`, `RFQ`, `RFQ WOMEN`, `RFQ YOUTH`, `RFQ PWD`.
3. **Per-quarter presence**: an item is "in" quarter `Qn` if its `Qn`
   quantity column is nonzero. (Confirmed against Tables row 26 in doc 6 —
   an AGPO/Youth item with quantity only in Q3, tag correctly appears only
   in Q3, not all four quarters — this rules out any "one tag for the whole
   segment/year" theory.)

### 2.2 Segment-level aggregation

For a given segment, for each quarter `Qn` independently:

```
tags_in_quarter(segment, Qn) = union of item-tags, over every qualifying
                                item in the segment with nonzero Qn quantity
```

Write that as a delimited string in a **fixed, consistent order**:
`OPEN, RFQ, RFQ WOMEN, RFQ YOUTH, RFQ PWD` (whichever subset applies,
comma-separated). If the set is empty, write `0` (matches your existing
convention in the OL KALAU workbook — file 5 — the cleanest of your
historical examples; blank-cell conventions varied elsewhere, worth locking
this down as the one true convention going forward rather than mixing).

The `OPEN/AGPO` summary column (column D) is the same union but taken across
**all four quarters combined** for that segment — i.e. does this segment
contain any Open Tender items anywhere in its year, and/or any RFQ items,
and/or any AGPO-reserved RFQ items. This is a coarser roll-up, not a
per-quarter value, and it matches why it reads e.g. `"OPEN,AGPO"` even
when no single quarter shows every category.

### 2.3 Validated against

- Item Details export, "Office Equipment & Accessories & Supplies" segment,
  Eldoret National Polytechnic (91 rows, all RFQ) — confirmed multi-category
  presence-per-quarter, ruled out "single dominant category" theories.
- Item Details export, "Furniture & Furnishings" segment, same entity
  (79 rows, all RFQ) — confirmed AGPO items can be present in a single
  quarter only (row 26, row 65).
- Item Details export, one Open Tender row, Athi Water Works Development
  Agency (Chlorine Cl, KES 84,672,000, `Procurement Method = "Open Tender"`)
  — confirmed Open Tender's blank reservation group and matched exactly
  against that entity's own workbook (segment 1: `OPEN/AGPO = "open"`,
  `Q1`–`Q4 = "open"` throughout, since it's a single item present every
  quarter).
- ~18 of your own hand-filled entity workbooks (files 5–9, 13–24, 26) —
  used to confirm the general shape (comma/slash-delimited unions per
  quarter) even though none of them individually used fully consistent
  formatting.

## 3. How this maps onto the existing pipeline

### 3.1 What's missing: the Item Details fetch

`scrape_entities.py` has two API-calling functions today:

```python
fetch_all_entity_records(context, xsrf, page_size=100)   # -> entities
fetch_entity_segments(context, xsrf, appdetailid, ...)   # -> segments
```

This feature needs a third, one level deeper:

```python
fetch_segment_items(context, xsrf, appdetailid, segment_ref, page_size=?)
```

`segment_ref` is the unknown piece — we don't yet know what value the
Angular app passes to identify *which segment's* items you want (UNSPSC
code like `44000000`? a segment id not currently captured by
`fetch_entity_segments`? something else?). This has to come from network
capture (see `capture_item_details_api.py`, already provided) since it's
not observable from the exported columns alone.

**Action item worth checking now, cheaply, without any network capture:**
look at the raw JSON `fetch_entity_segments()` already receives from
`view-app-summary` — it currently only extracts `description` and
`totalCost` into the segment dict:

```python
all_segments.append({
    "sr_no": i,
    "segment": row.get("description", ""),
    "total_cost": row.get("totalCost", 0.0),
})
```

`row` (the raw API response per segment) may already contain an id field
that's silently being discarded here — worth printing `row.keys()` once to
check before assuming a full network capture is necessary.

### 3.2 Where the new logic plugs in

- **`fetch_segment_items()`** — new function, sibling to
  `fetch_entity_segments()`, same pagination pattern (loop on `page`,
  stop when `len(all_items) >= total`). Returns raw item rows: procurement
  method, reservation group, women/youth/pwd quantities, Q1–Q4 quantities.
- **New pure function, e.g. `categorize_and_aggregate_items(items)`** —
  no API calls, just the rule from §2. Takes the raw item list for one
  segment, returns `{"open_agpo": "...", "q1": "...", "q2": "...", "q3": "...", "q4": "..."}`. Keeping this separate from the fetch function
  means it's independently testable against the sample data already in
  this conversation, before any live API access is needed.
- **`run_deep_scrape()`** — currently calls `fetch_entity_segments()` then
  `create_entity_template()`. Needs an additional loop per segment calling
  `fetch_segment_items()` + `categorize_and_aggregate_items()`, feeding
  results into the segment dict before it reaches `create_entity_template()`.
  This is the expensive part: one extra paginated API call *per segment*,
  not per entity — for an entity with 50+ segments (see the Eldoret
  example), that's 50+ additional calls where today there's one
  (`fetch_entity_segments` itself). Worth deciding whether to rate-limit /
  add delay here given the existing bot-protection concerns already
  documented for this site.
- **`create_entity_template()`** — currently writes `"", "", "", "", ""`
  for columns D–H. Change to write the four aggregated strings instead,
  keeping the same styling (thin borders, currency format on column C only,
  centered headers) already in place.
- **`Status` column semantics** (`Budget Totals` sheet) — the dev notes say
  `"Done"` is *"reserved for once OPEN/AGPO and Q1–Q4 data exists"* but
  *"nothing in the current pipeline sets this automatically."* Once this
  feature exists, `should_skip_entity()` / `reconcile_budget_totals_from_workbooks()`
  could reasonably promote an entity from `"Partial"` to `"Done"` once every
  segment in its workbook has a non-blank OPEN/AGPO value — but note the
  existing reconcile function's docstring explicitly warns against this:
  *"many legitimate segments have blank OPEN/AGPO..Q4 cells because those
  fields don't apply to them, not because the row is unfinished."* That's
  still true even with this feature built — a segment with zero qualifying
  (RFQ/Open Tender) items legitimately aggregates to all-`0` columns. Don't
  conflate "computed and found nothing" with "not yet computed."

### 3.3 What this feature does *not* touch

- `budget_tracker.py` (Status/Prequal round-trip with Google Sheets) —
  unaffected, since Total Budget is still segment-total-cost-based, not
  item-count-based.
- `sync_gsuite.py` / `code.gs` — the Google Sheets sync only pushes
  `Budget Totals`, `addresses`, and `blacklisted` sheets (5/9/2 columns
  respectively). Entity workbooks themselves are never synced to Sheets,
  so this feature is purely local-file-facing.
- `address_tracker.py` — fully unrelated.

## 4. Open items before implementation

1. **The endpoint itself.** Use `capture_item_details_api.py` (already
   generated) or check `view-app-summary`'s raw response for an unused
   segment-id field first — see §3.1.
2. **Pagination page size** for Item Details — Image 4 showed 10 items/page
   in the UI; unknown whether the API defaults to 10 or accepts a larger
   `pageSize` like `fetch_entity_segments` does (10) or
   `fetch_all_entity_records` does (100).
3. **Exact JSON field names** for procurement method / reservation group /
   quantities — the export column headers (`"Procurement Method"`,
   `"Preference & Reservation Group"`, etc.) are almost certainly not the
   literal JSON keys (compare: `view-app-summary`'s `"Segment"` export
   column is `description` in the raw JSON).
4. **Whether a procurement method beyond RFQ/Open Tender ever shows up**
   with a nonzero Women/Youth/PWD value — everything checked so far assumes
   it can't, but this hasn't been exhaustively verified across entities.
5. **Rate limiting / bot-protection risk** from the per-segment fetch
   volume — worth a small delay between calls, matching the caution already
   baked into `address_tracker.py`'s "batch delays."
