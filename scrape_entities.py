## scrape_entities.py   is being used for tests at the moment 
import os
import re
import json
import time
import openpyxl
import pandas as pd
import urllib.request
import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright
from sync_gsuite import BUDGET_SHEET_URL

EXCEL_FILE = "Procuring_Entities.xlsx"
TARGET_SHEET = "latest_entities"
SIMPLE_SHEET = "latest_entities_only"
BUDGET_SHEET = "Budget Totals"
PLANS_DIR = "./entity_plans"
API_BASE = "https://egpkenya.go.ke/api/app"
CURRENCY_FORMAT = '"KES" #,##0.00'
BLACKLISTED_SHEET = "blacklisted"
ITEM_DOWNLOAD_DIR = "./temp_downloads"
BIG_TICKET_THRESHOLD = 200000.0

os.makedirs(PLANS_DIR, exist_ok=True)
os.makedirs(ITEM_DOWNLOAD_DIR, exist_ok=True)

BLACKLIST_ENTITIES = {
    "OL KALOU TECHNICAL AND VOCATIONAL COLLEGE",
    "KAPSABET NANDI WATER AND SANITATION COMPANY LTD",
    "TECHNICAL AND VOCATIONAL EDUCATION AND TRAINING AUTHORITY",
    "KARATINA UNIVERSITY",
    "KIRINYAGA UNIVERSITY",
    "NATIONAL MUSEUMS OF KENYA",
    "KIBABII UNIVERSITY",
    "ALUPE UNIVERSITY",
    "OLLESSOS NATIONAL POLYTECHNIC",
    "MACHAKOS UNIVERSITY",
    "RONGO UNIVERSITY",
    "MATILI TECHNICAL TRAINING INSTITUTE",
    "LAIKIPIA UNIVERSITY",
    "GARISSA UNIVERSITY",
    "4815 KAKAMEGA - TRANSPORT, INFRASTRUCTURE, AND PUBLIC WORKS",
    "4172 WEST POKOT-COUNTY ASSEMBLY",
    "4772 BOMET - COUNTY ASSEMBLY PE",
    "4465-BARINGO TRANSPORT AND INFRASTRUCTURE",
    "4425 NANDI - OFFICE OF THE COUNTY ATTORNEY",
    "4811 COUNTY ASSEMBLY OF KAKAMEGA",
    "4432 NANDI - TRANSPORT, PUBLIC WORKS AND INFRASTRUCTURE DEVELOPMENT",
    "4866 VIHIGA EDUCATION SCIENCE AND TECHNOLOGY",
    "4517 LAIKIPIA TRADE,TOURISM & ENTERPRISE DEVELOPEMENT",
    "4427 NANDI - ADMINISTRATION PUBLIC SERVICE AND ICT",
    "4871 VIHIGA PUBLIC SERVICE AND ADMINSTRATION",
    "3913 NYERI-FINANCE,ECONOMIC PLANNING AND ICT",
    "3921 NYERI-WATER,ENVIRONMENT AND CLIMATE CHANGE",
    "4012 MURANGA - MINISTRY COUNTY COORDINATION AND ADMINSTRATION",
    "3914 NYERI-LANDS,PHYSICAL PLANNING AND URBAN DEVELOPMENT",
    "3918 NYERI-AGRICULTURE,LIVESTOCK AND AQUACULTURE DEVELOPMENT",
    "3919 NYERI-TRADE,CO-OPERATIVES,CULTURE AND TOURISM",
    "3211 COUNTY ASSEMBLY OF LAMU",
    "4018 MURANGA - MINISTRY OF HEALTH AND SANITATION",
    "4431 NANDI - LANDS,PHYSICAL PLANNING, HOUSING ENVI,WATER,NATURAL RESOURCES/CLI...",
    "4430 NANDI - EDUCATION VOCATIONAL TRAINING AND SCHOLARSHIP...",
    "4429 NANDI-SPORTS,YOUTH AFFAIRS,GENDER AND SOCIAL WELFARE ARTS...",
    "4869 VIHIGA COUNTY TREASURY",
    "4414 NANDI - HEALTH",
    "4469 BARINGO LANDS, HOUSING AND URBAN DEVELOPMENT",
    "4423 NANDI - COUNTY ASSEMBLY",
}


# In-process cache: the current run's authoritative blacklist, as a set of
# normalize_key()'d entity names. Populated either by scrape_to_latest_entities()
# (which pulls fresh from Google Sheets) or lazily from the local 'blacklisted'
# sheet when is_blacklisted() is called standalone (e.g. from address_tracker.py)
# without a light scrape having run first in this process.
_blacklist_cache = None


def pull_blacklist_from_gsheet(sheet_url=None):
    """Fetches the current blacklist entity list directly from the
    'blacklisted' tab in Google Sheets. This is the authoritative list —
    removing an entity here permanently un-blacklists it (even if it
    still matches the COUNTY pattern or the legacy hardcoded set), and
    adding one here blacklists it, going forward. Returns a set of
    normalize_key()'d names, or None if the fetch failed (caller should
    fall back to the last known local copy rather than treat this as
    "empty list").
    """
    base_url = sheet_url or BUDGET_SHEET_URL
    sep = "&" if "?" in base_url else "?"
    url = f"{base_url}{sep}sheet=blacklisted"
    try:
        with urllib.request.urlopen(url, timeout=30) as resp:
            payload = json.loads(resp.read().decode())
    except Exception as e:
        print(f"[WARNING] Could not fetch blacklist from Google Sheet: {e}")
        return None

    rows = payload.get("data", [])[1:]  # skip header row
    names = set()
    for row in rows:
        if len(row) >= 2 and row[1]:
            names.add(normalize_key(row[1]))
    return names


def _auto_detect_blacklist(entity_name: str) -> bool:
    """Heuristic used ONLY to seed the blacklist the first time an entity
    is ever seen. Once an entity has been classified in any prior run
    (blacklisted or not), the Google Sheet decides — this is never
    consulted again for it, so removing it from the sheet permanently
    un-blacklists it even if it still matches these patterns.
    """
    clean_name = str(entity_name).strip()
    upper_name = clean_name.upper()

    if clean_name in BLACKLIST_ENTITIES or upper_name in BLACKLIST_ENTITIES:
        return True
    if "COUNTY" in upper_name:
        return True
    if clean_name and clean_name[0].isdigit() and ("-" in clean_name[:6] or " " in clean_name[:6]):
        return True
    return False


def _set_blacklist_cache(entity_names):
    global _blacklist_cache
    _blacklist_cache = {normalize_key(n) for n in entity_names}


def _load_blacklist_cache_from_local_file(file_path=EXCEL_FILE):
    """Fallback loader: reads whatever 'blacklisted' sheet is already in
    the local workbook (last known good list from the previous run's
    Google Sheets pull). Used when is_blacklisted() is called without a
    fresh light scrape having populated the cache in this process yet.
    """
    global _blacklist_cache
    if _blacklist_cache is not None:
        return

    names = set()
    if os.path.exists(file_path):
        try:
            wb = load_workbook(file_path, data_only=True)
            if BLACKLISTED_SHEET in wb.sheetnames:
                ws = wb[BLACKLISTED_SHEET]
                for r in range(2, ws.max_row + 1):
                    entity = ws.cell(row=r, column=2).value
                    if entity:
                        names.add(normalize_key(entity))
        except Exception:
            pass

    if not names:
        # Last-resort fallback: the legacy hardcoded set, so is_blacklisted()
        # never returns wrong answers just because no file exists yet.
        names = {normalize_key(n) for n in BLACKLIST_ENTITIES}

    _blacklist_cache = names


def is_blacklisted(entity_name: str) -> bool:
    """Checks the current run's blacklist cache. Google Sheets is the
    source of truth (see scrape_to_latest_entities() for how the cache
    gets populated from there); the hardcoded set and pattern matching
    are only ever used to auto-seed a BRAND NEW entity's initial
    classification, never to override an explicit removal from the sheet.
    """
    _load_blacklist_cache_from_local_file()
    return normalize_key(entity_name) in _blacklist_cache


def normalize_key(name: str) -> str:
    s = str(name).strip()
    s = re.sub(r"\b(PP|APP|PROCUREMENT PLAN)\b$", "", s, flags=re.IGNORECASE).strip()
    return "".join(c for c in s.upper() if c.isalnum())


def check_and_warn_locked_files() -> bool:
    active_locks = []
    dirs_to_check = [".", PLANS_DIR] if os.path.exists(PLANS_DIR) else ["."]

    for d in dirs_to_check:
        for f in os.listdir(d):
            if f.startswith("~$") and f.endswith(".xlsx"):
                full_path = os.path.join(d, f)
                try:
                    os.remove(full_path)
                    print(f"[CLEANUP] Deleted orphan lock file: {full_path}")
                except OSError:
                    active_locks.append(full_path)

    if active_locks:
        print("\n" + "=" * 65)
        print("⚠️  ACTIVE EXCEL FILE LOCK DETECTED ⚠️")
        print("A process currently has these files open:")
        for tf in active_locks:
            print(f"  - {tf}")
        print("Please close Microsoft Excel or end the task in Task Manager.")
        print("=" * 65 + "\n")
        return False

    return True


def safe_save_workbook(wb: Workbook, filepath: str):
    try:
        wb.save(filepath)
    except PermissionError:
        print(f"\n[PERMISSION ERROR] Cannot save to '{filepath}'.")
        print("--> Please CLOSE this file in Microsoft Excel and re-run.\n")
        raise


def sync_and_rename_workbooks(master_df: pd.DataFrame):
    existing_files = [
        f for f in os.listdir(PLANS_DIR)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]

    valid_keys = {normalize_key(row["Procuring Entity"]) for _, row in master_df.iterrows()}

    file_map = {}
    for fname in existing_files:
        match = re.match(r"^(?:\d+\s+)?(.*)\.xlsx$", fname, re.IGNORECASE)
        if match:
            entity_part = match.group(1)
            norm_key = normalize_key(entity_part)

            if norm_key not in valid_keys:
                try:
                    os.remove(os.path.join(PLANS_DIR, fname))
                    print(f"[PURGE] Removed blacklisted workbook: {fname}")
                except Exception as e:
                    print(f"[WARNING] Could not delete {fname}: {e}")
            else:
                file_map[norm_key] = fname

    for _, row in master_df.iterrows():
        sr_no = int(row["Sr. No."])
        entity_name = str(row["Procuring Entity"]).strip()
        norm_key = normalize_key(entity_name)

        new_filename = f"{sr_no} {entity_name}.xlsx"
        new_filepath = os.path.join(PLANS_DIR, new_filename)

        if norm_key in file_map:
            old_filename = file_map[norm_key]
            old_filepath = os.path.join(PLANS_DIR, old_filename)
            if old_filename != new_filename:
                try:
                    os.replace(old_filepath, new_filepath)
                    print(f"[RENAME] {os.path.basename(old_filepath)} -> {os.path.basename(new_filepath)}")
                except PermissionError:
                    print(f"[WARNING] Could not rename {os.path.basename(old_filepath)}: File locked.")


# ---------------------------------------------------------------------------
# API-based session + fetch helpers (replaces DOM scraping / clicking)
# ---------------------------------------------------------------------------

def get_authenticated_context(playwright):
    """Loads the portal once to establish a valid session + XSRF cookie."""
    browser = playwright.chromium.launch(headless=True)
    context = browser.new_context()
    page = context.new_page()
    page.route("**/deskpro-messenger/**", lambda route: route.abort())
    page.goto("https://egpkenya.go.ke/public-app", wait_until="domcontentloaded", timeout=60000)
    page.wait_for_selector("table tbody tr", timeout=30000)

    cookies = context.cookies()
    xsrf = next((c["value"] for c in cookies if c["name"] == "XSRF-TOKEN"), None)
    if not xsrf:
        raise RuntimeError("Could not obtain XSRF-TOKEN — site may have changed its auth flow.")

    page.close()
    return browser, context, xsrf


def api_post(context, xsrf, path, body=None):
    headers = {"Content-Type": "application/json", "X-XSRF-TOKEN": xsrf}
    resp = context.request.post(
        f"{API_BASE}/{path}",
        headers=headers,
        data=json.dumps(body) if body is not None else None,
    )
    if resp.status != 200:
        raise RuntimeError(f"API call failed [{resp.status}]: {path}")
    return resp.json()


def fetch_all_entity_records(context, xsrf, page_size=100):
    """Pulls every listed entity/APP record via public-app-detail.
    Loop is safe even if the server ignores page_size and caps it lower,
    since 'page' always matches the number of successful iterations so far.
    """
    all_records = []
    page = 0
    total = None

    while True:
        body = {"appNumber": "", "finYear": 0, "procuringEntity": None, "pageSize": page_size, "page": page}
        result = api_post(context, xsrf, "public-app-detail", body)
        records = result.get("data", [])
        if not records:
            break

        if total is None:
            total = records[0].get("totalCount", len(records))

        all_records.extend(records)
        page += 1

        if len(all_records) >= total:
            break

    return all_records

def fetch_entity_created_date(context, xsrf, appdetailid):
    """Returns the entity's real creation date on the portal (YYYY-MM-DD),
    from get-appdetail-summary — NOT get-app-details, which leaks PII and
    is never called anywhere in this pipeline. Returns "" if unavailable.
    """
    try:
        result = api_post(context, xsrf, f"get-appdetail-summary/{appdetailid}")
    except Exception:
        return ""
    rows = result.get("reportdata", [])
    if not rows:
        return ""
    return str(rows[0].get("createdon", "")).strip()


def fetch_entity_segments(context, xsrf, appdetailid, page_size=10):
    """Pulls every procurement segment for an entity via view-app-summary.

    Captures 'segment_code' (the raw 'unspscsegment' field, e.g.
    "44000000") alongside segment/total_cost — this is the join key used
    by categorize_and_aggregate_items() to map Item Details rows back to
    the correct segment.
    """
    all_segments = []
    page = 1
    total = None

    while True:
        body = {
            "page": page,
            "searchTerm": json.dumps({"appDetailsId": appdetailid, "isSearch": False}),
            "pageSize": page_size,
            "appDetailsId": appdetailid,
            "isSearch": False,
        }
        result = api_post(context, xsrf, "view-app-summary", body)
        resp_data = result.get("respData", {})
        rows = resp_data.get("reportdata", [])
        if total is None:
            total = resp_data.get("totalcount", len(rows))

        for i, row in enumerate(rows, start=len(all_segments) + 1):
            all_segments.append({
                "sr_no": i,
                "segment": row.get("description", ""),
                "total_cost": row.get("totalCost", 0.0),
                "segment_code": str(row.get("unspscsegment", "")).strip(),
            })

        if not rows or len(all_segments) >= total:
            break
        page += 1

    return all_segments


# ---------------------------------------------------------------------------
# Stage 1a: Light Scrape (API-based)
# ---------------------------------------------------------------------------

def scrape_to_latest_entities(file_path=EXCEL_FILE):
    """Light Scrape: Pulls all entity/APP records via direct API calls.

    Blacklist source of truth: Google Sheets' 'blacklisted' tab, pulled
    fresh at the start of every run. An entity is blacklisted this run
    if and only if:
      - it's in the freshly-pulled sheet list, OR
      - it has NEVER been classified before (not present in the previous
        run's latest_entities or blacklisted sheet) AND matches the
        legacy auto-detection heuristic (COUNTY pattern / hardcoded set).
    Any entity previously seen and NOT currently in the sheet is treated
    as explicitly un-blacklisted, even if it would still match the
    heuristic — this is what makes removing an entity from the sheet
    actually stick permanently.
    """
    if not check_and_warn_locked_files():
        input("Press Enter after closing Excel to continue...")

    print("Launching browser to establish authenticated session...")
    with sync_playwright() as p:
        browser, context, xsrf = get_authenticated_context(p)
        try:
            print("Fetching full entity listing via API...")
            records = fetch_all_entity_records(context, xsrf)
            print(f"Retrieved {len(records)} raw records from API.")
        finally:
            browser.close()

    if not records:
        print("No valid records scraped.")
        return

    # 0. Read prior-run state before anything gets overwritten:
    #    - existing Date Added values, to preserve them across runs
    #    - the set of entities already classified before (whether they
    #      ended up blacklisted or not), so brand-new entities can be
    #      told apart from ones the sheet has already decided on
    today_str = datetime.date.today().isoformat()
    existing_dates = {}
    previously_classified = set()

    if os.path.exists(file_path):
        try:
            old_wb = load_workbook(file_path, data_only=True)
        except Exception:
            old_wb = None

        if old_wb is not None:
            if TARGET_SHEET in old_wb.sheetnames:
                old_sheet = old_wb[TARGET_SHEET]
                old_header = [old_sheet.cell(row=1, column=c).value for c in range(1, old_sheet.max_column + 1)]
                if "Procuring Entity" in old_header:
                    entity_col = old_header.index("Procuring Entity") + 1
                    appnum_col = old_header.index("APP Number") + 1 if "APP Number" in old_header else None
                    date_col = old_header.index("Date Added") + 1 if "Date Added" in old_header else None
                    for r in range(2, old_sheet.max_row + 1):
                        old_entity = old_sheet.cell(row=r, column=entity_col).value
                        if not old_entity:
                            continue
                        previously_classified.add(normalize_key(old_entity))
                        if date_col and appnum_col:
                            old_appnum = old_sheet.cell(row=r, column=appnum_col).value
                            old_date = old_sheet.cell(row=r, column=date_col).value
                            if old_date:
                                existing_dates[(str(old_entity).upper(), str(old_appnum or "").upper())] = old_date

            if BLACKLISTED_SHEET in old_wb.sheetnames:
                old_bl = old_wb[BLACKLISTED_SHEET]
                for r in range(2, old_bl.max_row + 1):
                    old_entity = old_bl.cell(row=r, column=2).value
                    if old_entity:
                        previously_classified.add(normalize_key(old_entity))

    # 1. Pull the authoritative blacklist from Google Sheets. If the fetch
    # fails (offline, Apps Script down, etc.), fall back to whatever's
    # already in the local file rather than treating it as an empty list.
    print("Pulling current blacklist from Google Sheet...")
    sheet_blacklist = pull_blacklist_from_gsheet()
    if sheet_blacklist is None:
        print("[WARNING] Using last local blacklist copy — could not reach Google Sheets.")
        _load_blacklist_cache_from_local_file(file_path)
        sheet_blacklist = set(_blacklist_cache) if _blacklist_cache else set()

    # 2. Classify every raw record.
    scraped_data = []
    blacklisted_seen = []
    for rec in records:
        entity = str(rec.get("procuringentity", "")).strip()
        if not entity:
            continue

        key = normalize_key(entity)
        if key in sheet_blacklist:
            is_bl = True
        elif key in previously_classified:
            is_bl = False  # previously seen, not currently in sheet -> stays un-blacklisted
        else:
            is_bl = _auto_detect_blacklist(entity)  # brand new -> seed via heuristic

        if is_bl:
            blacklisted_seen.append(entity)
            continue

        fin_year = rec.get("financialyear", "")
        app_num = rec.get("apprefno", "")
        appdetailid = rec.get("appdetailid")
        app_url = f"https://egpkenya.go.ke/public-view-app/{appdetailid}/1/"

        scraped_data.append((fin_year, entity, app_num, app_url, appdetailid))

    blacklisted_unique = list(dict.fromkeys(blacklisted_seen))

    seen = set()
    unique_records = []
    for fin_year, entity, app_num, app_url, appdetailid in scraped_data:
        key = (entity.upper(), str(app_num).upper())
        if key not in seen:
            seen.add(key)
            unique_records.append((fin_year, entity, app_num, app_url, appdetailid))

    if not unique_records:
        print("No valid records scraped.")
        return

    # 3. Fetch each entity's real creation date from the portal (skipped
    # for entities we already have a stored date for, to save calls).
    print(f"Fetching creation dates for {len(unique_records)} entities...")
    with sync_playwright() as p:
        browser2, context2, xsrf2 = get_authenticated_context(p)
        try:
            dated_records = []
            for fin_year, entity, app_num, app_url, appdetailid in unique_records:
                key = (entity.upper(), str(app_num).upper())
                if key in existing_dates:
                    created_date = existing_dates[key]
                else:
                    created_date = fetch_entity_created_date(context2, xsrf2, appdetailid) or today_str
                dated_records.append((fin_year, entity, app_num, app_url, appdetailid, created_date))
        finally:
            browser2.close()

    try:
        wb = (
            openpyxl.load_workbook(file_path)
            if os.path.exists(file_path)
            else openpyxl.Workbook()
        )
    except PermissionError:
        print(f"[PERMISSION ERROR] '{file_path}' is open in Excel. Close it and try again.")
        return

    # 4. Populate 'latest_entities' (7 columns, incl. APP Detail ID + Date Added).
    if TARGET_SHEET in wb.sheetnames:
        del wb[TARGET_SHEET]
    s_full = wb.create_sheet(TARGET_SHEET)
    s_full.append(
        ["Sr. No.", "Financial Year", "Procuring Entity", "APP Number", "APP URL", "APP Detail ID", "Date Added"]
    )
    for i, (fin_year, entity, app_num, app_url, appdetailid, created_date) in enumerate(dated_records, start=1):
        s_full.append([i, fin_year, entity, app_num, app_url, appdetailid, created_date])

    # 5. Populate 'latest_entities_only' — sorted by Date Added, newest first.
    if SIMPLE_SHEET in wb.sheetnames:
        del wb[SIMPLE_SHEET]
    s_simple = wb.create_sheet(SIMPLE_SHEET)
    s_simple.append(["Sr. No.", "Procuring Entity", "Date Added"])

    sorted_by_date = sorted(dated_records, key=lambda r: r[5] or "0000-00-00", reverse=True)
    for i, (_, entity, _, _, _, created_date) in enumerate(sorted_by_date, start=1):
        s_simple.append([i, entity, created_date])

    # 6. Preserve/Initialize 'Budget Totals'
    if BUDGET_SHEET not in wb.sheetnames:
        s_budget = wb.create_sheet(BUDGET_SHEET)
        s_budget.append(
            ["Sr. No.", "Procuring Entity", "Total Budget", "Status", "Prequalification done"]
        )
        for i, (_, entity, _, _, _, _) in enumerate(dated_records, start=1):
            s_budget.append([i, entity, 0.0, "Pending", ""])
            s_budget.cell(row=s_budget.max_row, column=3).number_format = CURRENCY_FORMAT
    else:
        s_budget = wb[BUDGET_SHEET]
        existing_data = {}
        for r in range(2, s_budget.max_row + 1):
            p_entity = s_budget.cell(row=r, column=2).value
            if p_entity:
                budget = s_budget.cell(row=r, column=3).value or 0.0
                status = s_budget.cell(row=r, column=4).value or "Pending"
                prequal = s_budget.cell(row=r, column=5).value or ""
                existing_data[normalize_key(p_entity)] = (budget, status, prequal)

        wb.remove(s_budget)
        s_budget = wb.create_sheet(BUDGET_SHEET)
        s_budget.append(
            ["Sr. No.", "Procuring Entity", "Total Budget", "Status", "Prequalification done"]
        )
        for i, (_, entity, _, _, _, _) in enumerate(dated_records, start=1):
            key = normalize_key(entity)
            budget, status, prequal = existing_data.get(key, (0.0, "Pending", ""))
            s_budget.append([i, entity, budget, status, prequal])
            s_budget.cell(row=s_budget.max_row, column=3).number_format = CURRENCY_FORMAT

    # 7. Populate 'blacklisted' sheet with the final decided list, and
    # update the in-process cache so later is_blacklisted() calls in the
    # same run (deep scrape, enrichment, address tracker) see it too.
    if BLACKLISTED_SHEET in wb.sheetnames:
        del wb[BLACKLISTED_SHEET]
    s_blacklist = wb.create_sheet(BLACKLISTED_SHEET)
    s_blacklist.append(["Sr. No.", "Procuring Entity"])
    for i, entity in enumerate(blacklisted_unique, start=1):
        s_blacklist.append([i, entity])

    _set_blacklist_cache(blacklisted_unique)

    # 8. Save
    safe_save_workbook(wb, file_path)
    print(f"Success! Written {len(dated_records)} clean records to '{EXCEL_FILE}'.")
    print(f"Logged {len(blacklisted_unique)} blacklisted entities to '{BLACKLISTED_SHEET}'.")

# ---------------------------------------------------------------------------
# Entity workbook generation
# ---------------------------------------------------------------------------

# Full tag -> short display label used in Q1-Q4 cells (RFQ prefix dropped,
# since the OPEN/AGPO summary column already states "RFQ AGPO" once).

def categorize_and_aggregate_items(items: list) -> dict:
    """One segment's items -> {"open_agpo": "...", "q1": "...", ...}

    Q1-Q4 show the full tag per item present that quarter (OPEN, RFQ,
    RFQ WOMEN, RFQ YOUTH, RFQ PWD), joined with "/" when more than one
    applies — e.g. "OPEN/RFQ WOMEN". "0" if nothing qualifies that quarter.

    OPEN/AGPO summary column only ever shows "OPEN" (any Open Tender item
    anywhere in the segment's year) and/or "RFQ AGPO" (any GENUINELY
    AGPO-reserved RFQ item — i.e. RFQ WOMEN/YOUTH/PWD, never plain
    unreserved RFQ). A segment with only plain "RFQ" tags and no OPEN
    items summarizes to "0" — unreserved RFQ carries no AGPO significance
    and should not be flagged as if it did.
    """
    quarter_tags = {"q1": set(), "q2": set(), "q3": set(), "q4": set()}
    all_tags = set()

    for item in items:
        tag = tag_for_item(item)
        if tag is None:
            continue
        all_tags.add(tag)
        for q in ("q1", "q2", "q3", "q4"):
            if item[q] > 0:
                quarter_tags[q].add(tag)

    QUARTER_ORDER = ["OPEN", "RFQ", "RFQ WOMEN", "RFQ YOUTH", "RFQ PWD"]
    AGPO_TAGS = {"RFQ WOMEN", "RFQ YOUTH", "RFQ PWD"}

    def quarter_join(tag_set):
        ordered = [t for t in QUARTER_ORDER if t in tag_set]
        return "/".join(ordered) if ordered else "0"

    def summary_join(tag_set):
        parts = []
        if "OPEN" in tag_set:
            parts.append("OPEN")
        if tag_set & AGPO_TAGS:
            parts.append("RFQ AGPO")
        return "/".join(parts) if parts else "0"

    return {
        "open_agpo": summary_join(all_tags),
        "q1": quarter_join(quarter_tags["q1"]),
        "q2": quarter_join(quarter_tags["q2"]),
        "q3": quarter_join(quarter_tags["q3"]),
        "q4": quarter_join(quarter_tags["q4"]),
    }
    
BIG_TICKET_HEADERS = [
    "Sr. No.", "UNSPSC/Item Code", "Item Description", "Procurement Type",
    "Quantity", "Unit of Issue", "Estimated Unit Cost", "Total Cost",
    "Procurement Method", "Reservation Group", "Tag",
]


def create_big_ticket_sheets(wb, segments: list, by_segment: dict, item_data: dict,
                              threshold: float = BIG_TICKET_THRESHOLD):
    """Adds one sheet per qualifying segment to `wb`. A segment qualifies
    when its overall budget exceeds `threshold` AND its aggregated
    OPEN/AGPO summary is non-'0' (i.e. contains OPEN and/or RFQ AGPO
    anywhere in the year). Within a qualifying segment, only individual
    items that (a) pass the same RFQ/Open Tender qualification rule as
    the rest of this feature, and (b) have their own Total Cost >=
    threshold, are listed. Sheet name is the segment's Sr. No. (matching
    the main sheet's numbering), since segment descriptions are often too
    long or contain characters Excel won't allow in a sheet name.
    """
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    align_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for seg in segments:
        seg_code = seg.get("segment_code", "")
        seg_total = seg.get("total_cost", 0.0)
        seg_result = item_data.get(seg_code, {})
        open_agpo = seg_result.get("open_agpo", "0")

        if seg_total <= threshold or open_agpo == "0":
            continue

        raw_items = by_segment.get(seg_code, [])
        qualifying_items = [it for it in raw_items if tag_for_item(it) is not None]
        big_items = [it for it in qualifying_items if it.get("total_cost", 0.0) >= threshold]
        if not big_items:
            continue  # segment qualifies overall, but no single item meets the per-item threshold

        sheet_name = str(seg.get("sr_no", seg_code)).strip()[:31] or seg_code[:31]
        if sheet_name in wb.sheetnames:
            del wb[sheet_name]
        ws = wb.create_sheet(sheet_name)

        ncols = len(BIG_TICKET_HEADERS)
        ws.merge_cells(f"A1:{get_column_letter(ncols)}1")
        title = ws["A1"]
        title.value = f"{seg.get('segment', '')} — Items ≥ KES {threshold:,.0f}"
        title.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        title.fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
        title.alignment = align_center
        ws.row_dimensions[1].height = 28

        ws.append(BIG_TICKET_HEADERS)
        for c in range(1, ncols + 1):
            cell = ws.cell(row=2, column=c)
            cell.font = font_header
            cell.fill = fill_header
            cell.alignment = align_center

        row_idx = 3
        for it in sorted(big_items, key=lambda x: x.get("total_cost", 0.0), reverse=True):
            tag = tag_for_item(it) or ""
            ws.append([
                it.get("sr_no", ""),
                it.get("item_code", ""),
                it.get("item_description", ""),
                it.get("procurement_type", ""),
                it.get("quantity", ""),
                it.get("unit_of_issue", ""),
                it.get("unit_cost", 0.0),
                it.get("total_cost", 0.0),
                it.get("procurement_method", ""),
                it.get("reservation_group", ""),
                tag,
            ])
            ws.cell(row=row_idx, column=7).number_format = CURRENCY_FORMAT
            ws.cell(row=row_idx, column=8).number_format = CURRENCY_FORMAT
            for c in range(1, ncols + 1):
                ws.cell(row=row_idx, column=c).alignment = align_wrap
            row_idx += 1

        for c in range(1, ncols + 1):
            col_letter = get_column_letter(c)
            max_len = max(
                (len(str(ws.cell(row=r, column=c).value or "")) for r in range(2, row_idx)),
                default=0,
            )
            ws.column_dimensions[col_letter].width = max(max_len + 3, 12)


def apply_big_ticket_sheets(filepath: str, segments: list, by_segment: dict, item_data: dict,
                             threshold: float = BIG_TICKET_THRESHOLD):
    """Reopens an already-saved entity workbook and adds/refreshes its
    big-ticket item sheets. Separate from create_entity_template() so
    the main sheet's save logic doesn't need to know about this feature.
    """
    try:
        wb = load_workbook(filepath)
    except Exception as e:
        print(f"[WARNING] Could not open {filepath} for big-ticket sheets: {e}")
        return
    create_big_ticket_sheets(wb, segments, by_segment, item_data, threshold=threshold)
    safe_save_workbook(wb, filepath)


def create_entity_template(sr_no: int, entity_name: str, app_no: str,
                            scraped_segments: list, item_data: dict = None) -> float:
    """Generates formatted entity Excel workbook and returns total budget."""
    item_data = item_data or {}

    filepath = os.path.join(PLANS_DIR, f"{sr_no} {entity_name}.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Procurement Plan"

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_bold = Font(name="Calibri", size=11, bold=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_wrap = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin_border = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    ws.merge_cells("A1:H1")
    banner = ws["A1"]
    banner.value = f"{entity_name.upper()} ({app_no}) TOTAL COST"
    banner.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
    banner.fill = PatternFill(start_color="002060", end_color="002060", fill_type="solid")
    banner.alignment = align_center
    ws.row_dimensions[1].height = 28

    headers = ["Sr. No.", "Segment", "Total Cost", "OPEN/AGPO", "Q1", "Q2", "Q3", "Q4"]
    ws.append(headers)
    ws.row_dimensions[2].height = 24

    for col_num in range(1, 9):
        cell = ws.cell(row=2, column=col_num)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center

    start_row = 3
    current_row = start_row

    for item in scraped_segments:
        seg_result = item_data.get(item.get("segment_code", ""), {})

        ws.append([
            item.get("sr_no", current_row - start_row + 1),
            item.get("segment", ""),
            item.get("total_cost", 0.0),
            seg_result.get("open_agpo", ""),
            seg_result.get("q1", ""),
            seg_result.get("q2", ""),
            seg_result.get("q3", ""),
            seg_result.get("q4", ""),
        ])

        ws.cell(row=current_row, column=3).number_format = "#,##0.00"
        for col in range(1, 9):
            cell = ws.cell(row=current_row, column=col)
            cell.border = thin_border
            cell.alignment = align_center if col in (1, 3, 4, 5, 6, 7, 8) else align_wrap
        current_row += 1

    total_row = current_row
    ws.cell(row=total_row, column=1, value="TOTAL").font = font_bold
    ws.cell(row=total_row, column=3, value=f"=SUM(C{start_row}:C{total_row-1})").font = font_bold
    ws.cell(row=total_row, column=3).number_format = "#,##0.00"

    for col in range(1, 9):
        ws.cell(row=total_row, column=col).border = Border(
            top=Side(style="thin", color="000000"),
            bottom=Side(style="double", color="000000"),
        )
        ws.cell(row=total_row, column=col).alignment = align_center

    # Width: base it on data rows only (start_row..total_row), never row 1 —
    # row 1 is the merged banner, and A1 (its top-left cell) holds the FULL
    # banner text, which would otherwise inflate column A's computed width
    # since merged cells store their value only in the top-left cell.
    for col_num in range(1, 9):
        col_letter = get_column_letter(col_num)
        max_len = max(
            (len(str(ws.cell(row=r, column=col_num).value or "")) for r in range(2, total_row + 1)),
            default=0,
        )
        ws.column_dimensions[col_letter].width = max(max_len + 3, 12)

    # Sr. No. is just small integers — force it thin regardless of the
    # above calculation (header text "Sr. No." would otherwise set it to ~12).
    ws.column_dimensions["A"].width = 8

    safe_save_workbook(wb, filepath)
    return sum(item.get("total_cost", 0.0) for item in scraped_segments)

def update_master_budget_totals(entity_name: str, calculated_total: float, status: str = "Partial"):
    """Updates entity budget total and status in 'Budget Totals' sheet.
    Column order: Sr. No. | Procuring Entity | Total Budget | Status | Prequalification done
    """
    try:
        wb = load_workbook(EXCEL_FILE)
    except PermissionError:
        print(f"[PERMISSION ERROR] '{EXCEL_FILE}' is open in Excel. Close it to update Budget Totals.")
        return

    ws = wb[BUDGET_SHEET] if BUDGET_SHEET in wb.sheetnames else wb.active
    norm_target = normalize_key(entity_name)

    for r in range(2, ws.max_row + 1):
        cell_val = str(ws.cell(row=r, column=2).value or "")
        if normalize_key(cell_val) == norm_target:
            budget_cell = ws.cell(row=r, column=3, value=calculated_total)
            budget_cell.number_format = CURRENCY_FORMAT
            ws.cell(row=r, column=4, value=status)
            break

    safe_save_workbook(wb, EXCEL_FILE)


def reconcile_budget_totals_from_workbooks(plans_dir=PLANS_DIR, file_path=EXCEL_FILE):
    """
    Walks every entity workbook in `plans_dir` and sums all genuine segment
    rows in column C (Total Cost). Rows are excluded from the sum only if:
      (a) column A or B literally reads "TOTAL" (legacy summary rows), or
      (b) column C holds a formula rather than a literal number — since
          data_only=False returns formula cells as their formula TEXT
          (e.g. "=SUM(C2:C56)"), which fails float conversion and is
          naturally skipped without needing special-case detection.
    This means unlabeled trailing SUM-formula rows (no "TOTAL" text) are
    already safe and won't be double-counted.

    Status is intentionally NOT auto-promoted to "Done" here: many
    legitimate segments have blank OPEN/AGPO..Q4 cells because those
    fields don't apply to them, not because the row is unfinished — and
    there's no reliable way to tell the two apart from cell contents
    alone. This only ever moves Pending -> Partial when a nonzero total
    is found, and never touches an existing "Done" or downgrades anything.
    Mark "Done" manually once you're satisfied an entity is complete.
    """
    if not os.path.exists(plans_dir):
        print(f"[ERROR] '{plans_dir}' does not exist.")
        return

    workbook_files = [f for f in os.listdir(plans_dir) if f.endswith(".xlsx") and not f.startswith("~$")]
    if not workbook_files:
        print("No entity workbooks found to reconcile.")
        return

    try:
        master_wb = load_workbook(file_path)
    except PermissionError:
        print(f"[PERMISSION ERROR] '{file_path}' is open in Excel. Close it and re-run.")
        return

    if BUDGET_SHEET not in master_wb.sheetnames:
        print(f"[ERROR] '{BUDGET_SHEET}' sheet not found in {file_path}.")
        return

    budget_ws = master_wb[BUDGET_SHEET]

    row_lookup = {}
    for r in range(2, budget_ws.max_row + 1):
        entity_val = budget_ws.cell(row=r, column=2).value
        if entity_val:
            row_lookup[normalize_key(entity_val)] = r

    updated, skipped, unmatched = 0, 0, []

    def to_number(val):
        if isinstance(val, (int, float)):
            return float(val)
        if isinstance(val, str):
            try:
                return float(val.replace(",", "").replace("KES", "").strip())
            except ValueError:
                return None  # covers formula text like "=SUM(...)" too
        return None

    def is_total_label(val) -> bool:
        return str(val or "").strip().upper() == "TOTAL"

    for fname in workbook_files:
        filepath = os.path.join(plans_dir, fname)
        match = re.match(r"^(?:\d+\s+)?(.*)\.xlsx$", fname, re.IGNORECASE)
        entity_from_filename = match.group(1) if match else fname.replace(".xlsx", "")
        norm_key = normalize_key(entity_from_filename)

        try:
            wb = load_workbook(filepath, data_only=False)
        except Exception as e:
            print(f"[ERROR] Could not open {fname}: {e}")
            continue

        ws = wb.active

        total = 0.0
        row_count = 0
        for r in range(3, ws.max_row + 1):
            col_a = ws.cell(row=r, column=1).value
            col_b = ws.cell(row=r, column=2).value
            if is_total_label(col_a) or is_total_label(col_b):
                continue

            num = to_number(ws.cell(row=r, column=3).value)
            if num is not None:
                total += num
                row_count += 1

        if row_count == 0:
            unmatched.append(f"{fname} (no numeric values found in column C)")
            continue

        target_row = row_lookup.get(norm_key)
        if target_row is None:
            unmatched.append(fname)
            continue

        current_status = str(budget_ws.cell(row=target_row, column=4).value or "").strip().lower()
        if current_status == "done":
            new_status = "Done"
        elif total > 0:
            new_status = "Partial"
        else:
            new_status = "Pending"

        current_val = budget_ws.cell(row=target_row, column=3).value or 0.0
        values_match = isinstance(current_val, (int, float)) and abs(current_val - total) < 0.01
        status_match = current_status == new_status.lower()

        if values_match and status_match:
            skipped += 1
            continue

        budget_cell = budget_ws.cell(row=target_row, column=3, value=total)
        budget_cell.number_format = CURRENCY_FORMAT
        budget_ws.cell(row=target_row, column=4, value=new_status)

        updated += 1
        print(f"[RECONCILE] {entity_from_filename}: KES {total:,.2f} ({row_count} segments -> {new_status})")

    safe_save_workbook(master_wb, file_path)

    print(f"\nDone. Updated: {updated}, already correct: {skipped}, unmatched: {len(unmatched)}")
    if unmatched:
        print("Unmatched:")
        for f in unmatched:
            print(f"  - {f}")


def should_skip_entity(entity_name: str, budget_df: pd.DataFrame) -> bool:
    if budget_df is None or budget_df.empty:
        return False

    norm_target = normalize_key(entity_name)
    for _, row in budget_df.iterrows():
        ent_name = str(row.get("Procuring Entity", ""))
        if normalize_key(ent_name) == norm_target:
            status = str(row.get("Status", "")).strip().lower()
            try:
                budget = float(row.get("Total Budget", 0.0))
            except (ValueError, TypeError):
                budget = 0.0

            if status == "done" or budget > 0:
                return True
    return False


# ---------------------------------------------------------------------------
# Stage 1b: Deep Scrape (API-based)
# ---------------------------------------------------------------------------

def run_deep_scrape(file_path=EXCEL_FILE, overwrite=False):
    """Deep Scrape via direct API calls keyed on APP Detail ID — no clicking, no DOM waits."""
    if not check_and_warn_locked_files():
        input("Press Enter after closing Excel to continue...")

    try:
        df = pd.read_excel(file_path, sheet_name=TARGET_SHEET)
        try:
            budget_df = pd.read_excel(file_path, sheet_name=BUDGET_SHEET)
        except Exception:
            budget_df = pd.DataFrame()
    except PermissionError:
        print(f"[PERMISSION ERROR] Cannot read '{file_path}'. Please close Microsoft Excel.")
        return

    if "APP Detail ID" not in df.columns:
        print("[ERROR] 'latest_entities' has no 'APP Detail ID' column. Re-run Step 1a (Light Scrape) first.")
        return

    sync_and_rename_workbooks(df)

    print("Launching browser to establish authenticated session...")
    with sync_playwright() as p:
        browser, context, xsrf = get_authenticated_context(p)

        try:
            for _, row in df.iterrows():
                sr_no = int(row["Sr. No."])
                entity = str(row["Procuring Entity"]).strip()
                app_no = str(row["APP Number"]).strip() if pd.notna(row.get("APP Number")) else ""
                appdetailid = row.get("APP Detail ID")

                if pd.isna(appdetailid):
                    print(f"[WARNING] Missing APP Detail ID for [{sr_no}] {entity}. Skipping.")
                    continue
                appdetailid = int(appdetailid)

                target_filepath = os.path.join(PLANS_DIR, f"{sr_no} {entity}.xlsx")

                if not overwrite and should_skip_entity(entity, budget_df):
                    print(f"[SKIP] Entity completed/has budget: [{sr_no}] {entity}")
                    continue

                if os.path.exists(target_filepath) and not overwrite:
                    print(f"[SKIP] Workbook exists: {target_filepath}")
                    continue

                print(f"Deep Scraping [{sr_no}] {entity} (id={appdetailid})...")
                try:
                    segments = fetch_entity_segments(context, xsrf, appdetailid)
                    tot = create_entity_template(sr_no, entity, app_no, segments)
                    update_master_budget_totals(entity, tot, status="Partial")
                    print(f"  -> {len(segments)} segments, total KES {tot:,.2f}")
                except Exception as e:
                    print(f"[ERROR] Failed deep scrape for {entity}: {e}")
        finally:
            browser.close()


# ---------------------------------------------------------------------------
# Item Details / OPEN-AGPO feature
# ---------------------------------------------------------------------------

def derive_segment_code(item_code: str) -> str:
    """UNSPSC item codes are hierarchical: first 2 digits are the segment.
    E.g. '10101510' -> '10000000'. Not a heuristic — this is what the
    first 2 digits of a UNSPSC code mean by definition.
    """
    digits = re.sub(r"\D", "", str(item_code))
    return digits[:2] + "000000" if len(digits) >= 2 else ""


def _to_float(val) -> float:
    try:
        return float(str(val).replace(",", "").strip())
    except (ValueError, TypeError):
        return 0.0


def tag_for_item(item: dict):
    """Returns the tag for a qualifying item, or None if it doesn't qualify
    (not Request for Quotation / Open Tender)."""
    method = item["procurement_method"]

    if method == "Open Tender":
        return "OPEN"

    if method == "Request for Quotation":
        if item["reservation_group"].upper() == "AGPO":
            if item["women"] > 0:
                return "RFQ WOMEN"
            if item["youth"] > 0:
                return "RFQ YOUTH"
            if item["pwd"] > 0:
                return "RFQ PWD"
            return "RFQ"
        return "RFQ"

    return None


def parse_item_details_export(filepath: str) -> list:
    """Reads a downloaded Item Details export into a list of raw item dicts.
    Keeps both the categorization fields (procurement_method, reservation_group,
    women/youth/pwd, q1-q4 quantities) and the display fields needed for the
    big-ticket item sheets (item description, quantity, costs, etc.).
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    header = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    col_idx = {name: i for i, name in enumerate(header)}

    def get(row_vals, name, default=""):
        idx = col_idx.get(name)
        return row_vals[idx] if idx is not None and idx < len(row_vals) else default

    items = []
    for r in range(2, ws.max_row + 1):
        row_vals = [ws.cell(row=r, column=c).value for c in range(1, ws.max_column + 1)]
        if not any(v not in (None, "") for v in row_vals):
            continue

        item_code = get(row_vals, "UNSPSC/Item Code")
        items.append({
            "segment_code": derive_segment_code(item_code),
            "sr_no": get(row_vals, "Sr.No.", ""),
            "item_code": item_code,
            "item_description": get(row_vals, "UNSPSC/Item Description", ""),
            "procurement_type": get(row_vals, "Procurement Type", ""),
            "quantity": get(row_vals, "Quantity", ""),
            "unit_of_issue": get(row_vals, "Unit of Issue", ""),
            "unit_cost": _to_float(get(row_vals, "Estimated Unit Cost", 0)),
            "total_cost": _to_float(get(row_vals, "Total Cost", 0)),
            "procurement_method": str(get(row_vals, "Procurement Method", "")).strip(),
            "reservation_group": str(get(row_vals, "Preference & Reservation Group", "")).strip(),
            "women": _to_float(get(row_vals, "Women", 0)),
            "youth": _to_float(get(row_vals, "Youth", 0)),
            "pwd": _to_float(get(row_vals, "PWD", 0)),
            "q1": _to_float(get(row_vals, "Q1", 0)),
            "q2": _to_float(get(row_vals, "Q2", 0)),
            "q3": _to_float(get(row_vals, "Q3", 0)),
            "q4": _to_float(get(row_vals, "Q4", 0)),
        })

    return items


def search_and_open_entity(page, app_number: str, entity_name: str) -> str:
    """Navigates the listing, searches by APP Number, clicks the matching
    row, and lands on the entity's real hashed public-view-app URL."""
    page.goto("https://egpkenya.go.ke/public-app", wait_until="domcontentloaded", timeout=60000)
    page.wait_for_selector("table tbody tr", timeout=30000)

    toggle = page.locator("#first-toggle")
    if toggle.count() == 0:
        raise RuntimeError("Search accordion toggle (#first-toggle) not found.")

    if toggle.get_attribute("aria-expanded") != "true":
        toggle.click()
        page.wait_for_function(
            "document.querySelector('#first-toggle')?.getAttribute('aria-expanded') === 'true'",
            timeout=5000,
        )

    search_input = page.locator("input[formcontrolname='appNumber']")
    if search_input.count() == 0:
        raise RuntimeError("Search input (formcontrolname='appNumber') not found after expanding accordion.")
    search_input.first.wait_for(state="visible", timeout=5000)
    search_input.first.fill(app_number)
    page.wait_for_timeout(300)

    search_button = page.locator("button[type='submit']:has-text('Search')")
    if search_button.count() == 0:
        raise RuntimeError("Search submit button not found.")
    search_button.first.click()
    page.wait_for_timeout(2000)

    row_match = page.locator(f"text={entity_name}")
    if row_match.count() == 0:
        raise RuntimeError(f"No row found matching entity name '{entity_name}' after search.")

    row = page.locator("tr", has=row_match.first)
    link = row.locator("a").first
    if link.count() == 0:
        raise RuntimeError("Matching row found but it has no clickable link.")

    with page.expect_navigation(timeout=15000):
        link.click()

    return page.url


def download_and_parse_item_details(page) -> list:
    """Assumes `page` is already on an entity's real public-view-app URL
    (i.e. right after search_and_open_entity()). Clicks Item Details,
    clicks Export to Excel, catches the download, parses it, deletes the
    temp file, and returns the raw item list.
    """
    page.click("text=Item Details")
    page.wait_for_timeout(2000)

    with page.expect_download(timeout=30000) as download_info:
        page.click("text=Export to Excel")
    download = download_info.value

    temp_path = os.path.join(ITEM_DOWNLOAD_DIR, f"export_{os.getpid()}_{id(page)}.xlsx")
    download.save_as(temp_path)

    try:
        items = parse_item_details_export(temp_path)
    finally:
        try:
            os.remove(temp_path)
        except OSError:
            pass

    return items


def entity_needs_item_details(filepath: str) -> bool:
    """True if the entity workbook doesn't exist yet, or any of its
    segment rows has a blank OPEN/AGPO (column D) value."""
    if not os.path.exists(filepath):
        return True
    try:
        wb = load_workbook(filepath, data_only=True)
    except Exception:
        return True

    ws = wb.active
    for r in range(3, ws.max_row + 1):
        col_a = ws.cell(row=r, column=1).value
        if str(col_a or "").strip().upper() == "TOTAL":
            continue
        segment_val = ws.cell(row=r, column=2).value
        if not segment_val:
            continue
        open_agpo_val = ws.cell(row=r, column=4).value
        if open_agpo_val in (None, ""):
            return True

    return False


def run_item_details_enrichment(file_path=EXCEL_FILE, overwrite=False, delay_seconds=2.0, limit=None):
    """
    Standalone step: populates OPEN/AGPO + Q1-Q4 for every entity that
    needs it. Deliberately separate from run_deep_scrape() — that
    function's skip logic is keyed on "has a Total Budget", which is
    unrelated to whether OPEN/AGPO has been populated; almost every
    entity already has a budget by the time this feature matters, so
    folding this into run_deep_scrape() would mean it almost never runs.

    `limit`: if set, only processes the first N entities that need
    enrichment — use this for a small test run before running against
    everything.
    """
    if not check_and_warn_locked_files():
        input("Press Enter after closing Excel to continue...")

    try:
        df = pd.read_excel(file_path, sheet_name=TARGET_SHEET)
    except PermissionError:
        print(f"[PERMISSION ERROR] Cannot read '{file_path}'. Please close Microsoft Excel.")
        return

    if "APP Number" not in df.columns or "APP Detail ID" not in df.columns:
        print("[ERROR] 'latest_entities' missing required columns. Re-run Step 1a first.")
        return

    processed = 0
    print("Launching browser to establish authenticated session...")
    with sync_playwright() as p:
        browser, context, xsrf = get_authenticated_context(p)
        page = context.new_page()
        page.route("**/deskpro-messenger/**", lambda route: route.abort())

        try:
            for _, row in df.iterrows():
                if limit is not None and processed >= limit:
                    print(f"[LIMIT] Reached test limit of {limit} entities. Stopping.")
                    break

                sr_no = int(row["Sr. No."])
                entity = str(row["Procuring Entity"]).strip()
                app_no = str(row["APP Number"]).strip() if pd.notna(row.get("APP Number")) else ""
                appdetailid = row.get("APP Detail ID")

                if pd.isna(appdetailid) or not app_no:
                    print(f"[WARNING] Missing APP Detail ID or APP Number for [{sr_no}] {entity}. Skipping.")
                    continue
                appdetailid = int(appdetailid)

                target_filepath = os.path.join(PLANS_DIR, f"{sr_no} {entity}.xlsx")

                if not overwrite and not entity_needs_item_details(target_filepath):
                    print(f"[SKIP] Already enriched: [{sr_no}] {entity}")
                    continue

                print(f"Enriching [{sr_no}] {entity} (id={appdetailid})...")
                try:
                    segments = fetch_entity_segments(context, xsrf, appdetailid)

                    search_and_open_entity(page, app_no, entity)
                    items = download_and_parse_item_details(page)

                    by_segment = {}
                    for item in items:
                        by_segment.setdefault(item["segment_code"], []).append(item)

                    item_data = {
                        seg: categorize_and_aggregate_items(rows)
                        for seg, rows in by_segment.items()
                    }

                    tot = create_entity_template(sr_no, entity, app_no, segments, item_data=item_data)
                    apply_big_ticket_sheets(target_filepath, segments, by_segment, item_data)
                    update_master_budget_totals(entity, tot, status="Done")
                    print(f"  -> {len(segments)} segments, {len(items)} items parsed, total KES {tot:,.2f}")
                    processed += 1

                except Exception as e:
                    print(f"[ERROR] Failed item-details enrichment for [{sr_no}] {entity}: {e}")

                time.sleep(delay_seconds)
        finally:
            browser.close()

    print(f"\nDone. Enriched {processed} entities.")