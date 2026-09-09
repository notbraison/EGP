## scrape_entities.py
import os
import re
import json
import openpyxl
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from playwright.sync_api import sync_playwright

EXCEL_FILE = "Procuring_Entities.xlsx"
TARGET_SHEET = "latest_entities"
SIMPLE_SHEET = "latest_entities_only"
BUDGET_SHEET = "Budget Totals"
PLANS_DIR = "./entity_plans"
API_BASE = "https://egpkenya.go.ke/api/app"
CURRENCY_FORMAT = '"KES" #,##0.00'
BLACKLISTED_SHEET = "blacklisted"

os.makedirs(PLANS_DIR, exist_ok=True)

BLACKLIST_ENTITIES = {
    "OL KALOU TECHNICAL AND VOCATIONAL COLLEGE",
    "THIKA WATER AND SEWERAGE COMPANY LTD",
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


def is_blacklisted(entity_name: str) -> bool:
    clean_name = str(entity_name).strip()
    upper_name = clean_name.upper()

    if clean_name in BLACKLIST_ENTITIES or upper_name in BLACKLIST_ENTITIES:
        return True
    if "COUNTY" in upper_name:
        return True
    if clean_name and clean_name[0].isdigit() and ("-" in clean_name[:6] or " " in clean_name[:6]):
        return True
    return False


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


def fetch_entity_segments(context, xsrf, appdetailid, page_size=10):
    """Pulls every procurement segment for an entity via view-app-summary."""
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
            })

        if not rows or len(all_segments) >= total:
            break
        page += 1

    return all_segments


# ---------------------------------------------------------------------------
# Stage 1a: Light Scrape (API-based)
# ---------------------------------------------------------------------------

def scrape_to_latest_entities(file_path=EXCEL_FILE):
    """Light Scrape: Pulls all entity/APP records via direct API calls, filters blacklist."""
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

    scraped_data = []
    blacklisted_seen = []
    for rec in records:
        entity = str(rec.get("procuringentity", "")).strip()
        if not entity:
            continue
        if is_blacklisted(entity):
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

    try:
        wb = (
            openpyxl.load_workbook(file_path)
            if os.path.exists(file_path)
            else openpyxl.Workbook()
        )
    except PermissionError:
        print(f"[PERMISSION ERROR] '{file_path}' is open in Excel. Close it and try again.")
        return

    # 1. Populate 'latest_entities' (6 columns, incl. APP Detail ID)
    if TARGET_SHEET in wb.sheetnames:
        del wb[TARGET_SHEET]
    s_full = wb.create_sheet(TARGET_SHEET)
    s_full.append(
        ["Sr. No.", "Financial Year", "Procuring Entity", "APP Number", "APP URL", "APP Detail ID"]
    )
    for i, (fin_year, entity, app_num, app_url, appdetailid) in enumerate(unique_records, start=1):
        s_full.append([i, fin_year, entity, app_num, app_url, appdetailid])

    # 2. Populate 'latest_entities_only'
    if SIMPLE_SHEET in wb.sheetnames:
        del wb[SIMPLE_SHEET]
    s_simple = wb.create_sheet(SIMPLE_SHEET)
    s_simple.append(["Sr. No.", "Procuring Entity"])
    for i, (_, entity, _, _, _) in enumerate(unique_records, start=1):
        s_simple.append([i, entity])

    # 3. Preserve/Initialize 'Budget Totals'
    if BUDGET_SHEET not in wb.sheetnames:
        s_budget = wb.create_sheet(BUDGET_SHEET)
        s_budget.append(
            ["Sr. No.", "Procuring Entity", "Total Budget", "Status", "Prequalification done"]
        )
        for i, (_, entity, _, _, _) in enumerate(unique_records, start=1):
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
        for i, (_, entity, _, _, _) in enumerate(unique_records, start=1):
            key = normalize_key(entity)
            budget, status, prequal = existing_data.get(key, (0.0, "Pending", ""))
            s_budget.append([i, entity, budget, status, prequal])
            s_budget.cell(row=s_budget.max_row, column=3).number_format = CURRENCY_FORMAT

    # 4. Populate 'blacklisted' sheet
    if BLACKLISTED_SHEET in wb.sheetnames:
        del wb[BLACKLISTED_SHEET]
    s_blacklist = wb.create_sheet(BLACKLISTED_SHEET)
    s_blacklist.append(["Sr. No.", "Procuring Entity"])
    for i, entity in enumerate(blacklisted_unique, start=1):
        s_blacklist.append([i, entity])

    # 5. Save
    safe_save_workbook(wb, file_path)
    print(f"Success! Written {len(unique_records)} clean records to '{EXCEL_FILE}'.")
    print(f"Logged {len(blacklisted_unique)} blacklisted entities to '{BLACKLISTED_SHEET}'.")

# ---------------------------------------------------------------------------
# Entity workbook generation (as you edited — unchanged)
# ---------------------------------------------------------------------------

def create_entity_template(sr_no: int, entity_name: str, app_no: str, scraped_segments: list) -> float:
    """Generates formatted entity Excel workbook and returns total budget.

    scraped_segments items only need: sr_no, segment, total_cost.
    OPEN/AGPO and Q1-Q4 columns are left blank pending a future deep-scrape stage.
    """
    filepath = os.path.join(PLANS_DIR, f"{sr_no} {entity_name}.xlsx")
    wb = Workbook()
    ws = wb.active
    ws.title = "Procurement Plan"

    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    font_bold = Font(name="Calibri", size=11, bold=True)
    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
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
        ws.append([
            item.get("sr_no", current_row - start_row + 1),
            item.get("segment", ""),
            item.get("total_cost", 0.0),
            "",
            "",
            "",
            "",
            "",
        ])

        ws.cell(row=current_row, column=3).number_format = "#,##0.00"
        for col in range(1, 9):
            ws.cell(row=current_row, column=col).border = thin_border
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

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[get_column_letter(col[0].column)].width = max(max_len + 3, 12)

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

if __name__ == "__main__":
    reconcile_budget_totals_from_workbooks()


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