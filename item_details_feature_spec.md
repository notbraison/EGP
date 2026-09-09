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
