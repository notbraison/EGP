import os
import re
import sys
import time
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

os.makedirs(PLANS_DIR, exist_ok=True)

# Explicit list of 42 trimmed/county entities to drop automatically
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
    """Checks if entity matches explicit blacklist or county naming patterns."""
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
    """Normalizes strings for strict matching."""
    s = str(name).strip()
    s = re.sub(r"\b(PP|APP|PROCUREMENT PLAN)\b$", "", s, flags=re.IGNORECASE).strip()
    return "".join(c for c in s.upper() if c.isalnum())


def check_and_warn_locked_files() -> bool:
    """Checks for Excel lock files ('~$*.xlsx') and clears orphaned locks."""
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
    """Saves workbook with explicit PermissionError handling."""
    try:
        wb.save(filepath)
    except PermissionError:
        print(f"\n[PERMISSION ERROR] Cannot save to '{filepath}'.")
        print("--> Please CLOSE this file in Microsoft Excel and re-run.\n")
        raise


def sync_and_rename_workbooks(master_df: pd.DataFrame):
    """Renames existing entity workbooks and deletes blacklisted/orphan files."""
    existing_files = [
        f for f in os.listdir(PLANS_DIR)
        if f.endswith(".xlsx") and not f.startswith("~$")
    ]
    
    # Build set of valid normalized keys from master_df
    valid_keys = {normalize_key(row["Procuring Entity"]) for _, row in master_df.iterrows()}
    
    file_map = {}
    for fname in existing_files:
        match = re.match(r"^(?:\d+\s+)?(.*)\.xlsx$", fname, re.IGNORECASE)
        if match:
            entity_part = match.group(1)
            norm_key = normalize_key(entity_part)
            
            # Delete file if entity is blacklisted/removed from latest_entities
            if norm_key not in valid_keys:
                try:
                    os.remove(os.path.join(PLANS_DIR, fname))
                    print(f"[PURGE] Removed blacklisted workbook: {fname}")
                except Exception as e:
                    print(f"[WARNING] Could not delete {fname}: {e}")
            else:
                file_map[norm_key] = fname

    # Rename active files to match new Sr. No. indexes
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
                    

def get_real_app_url(elem):
    """Extracts viewApp parameters from onclick or href attributes on a link element,

    handling both quoted ('28506') and unquoted (28506) argument formats.
    """
    if not elem:
        return ""

    # Combine onclick and href to ensure we check both
    onclick_attr = elem.get_attribute("onclick") or ""
    href_attr = elem.get_attribute("href") or ""
    combined_attr = f"{onclick_attr} {href_attr}"

    # Target viewApp(...) function call
    match = re.search(r"viewApp\((.*?)\)", combined_attr)
    if match:
        raw_args = match.group(1)
        # Strip quotes and spaces from arguments
        args = [
            arg.strip(" '\"") for arg in raw_args.split(",") if arg.strip()
        ]

        if len(args) >= 3:
            app_id, mode, app_hash = args[0], args[1], args[2]
            return f"https://egpkenya.go.ke/public-view-app/{app_id}/{mode}/{app_hash}"

    # Fallback: Direct relative/absolute path in href
    if href_attr.startswith("/public-view-app"):
        return f"https://egpkenya.go.ke{href_attr}"
    elif href_attr.startswith("http"):
        return href_attr

    return ""


def scrape_to_latest_entities(file_path=EXCEL_FILE):
    """Light Scrape: Extracts portal listing, converts void(0) JavaScript links into direct URLs, and filters out blacklisted entities."""
    if not check_and_warn_locked_files():
        input("Press Enter after closing Excel to continue...")

    scraped_data = []

    with sync_playwright() as p:
        print("Launching browser for Light Scrape...")
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1280, "height": 800})
        page = context.new_page()

        page.route(
            "**/*.{png,jpg,jpeg,svg,gif,woff,woff2,ttf}",
            lambda route: route.abort(),
        )

        print("Navigating to e-GP Public APP page...")
        page.goto(
            "https://egpkenya.go.ke/public-app", wait_until="domcontentloaded"
        )

        current_page = 1

        while True:
            print(f"Scraping Page {current_page}...")
            page.wait_for_selector("table tbody tr", timeout=10000)
            time.sleep(0.5)

            rows = page.query_selector_all("table tbody tr")
            page_has_data = False

            for row in rows:
                cols = row.query_selector_all("td")
                if len(cols) >= 4:
                    fin_year = cols[1].inner_text().strip()
                    entity = cols[2].inner_text().strip()

                    # Apply early blacklist filter
                    if is_blacklisted(entity):
                        continue

                    app_num_elem = cols[3].query_selector("a")
                    if app_num_elem:
                        app_num = app_num_elem.inner_text().strip()
                        # Pass app_num_elem directly (NOT row)
                        app_url = get_real_app_url(app_num_elem)
                    else:
                        app_num = cols[3].inner_text().strip()
                        app_url = ""

                    if entity:
                        scraped_data.append((fin_year, entity, app_num, app_url))
                        page_has_data = True

            next_btn = page.query_selector(
                "ul.pagination li:last-child a, .pagination button:last-child"
            )

            is_disabled = False
            if next_btn:
                parent_class = (
                    next_btn.evaluate("el => el.parentElement.className") or ""
                )
                btn_class = next_btn.evaluate("el => el.className") or ""
                if (
                    "disabled" in parent_class.lower()
                    or "disabled" in btn_class.lower()
                ):
                    is_disabled = True

            if next_btn and not is_disabled and page_has_data:
                current_page += 1
                next_btn.click()
                time.sleep(1.5)
            else:
                print("Reached final page.")
                break

        browser.close()

    if not scraped_data:
        print("No valid records scraped.")
        return

    seen = set()
    unique_records = []
    for fin_year, entity, app_num, app_url in scraped_data:
        key = (entity.upper(), app_num.upper())
        if key not in seen:
            seen.add(key)
            unique_records.append((fin_year, entity, app_num, app_url))

    try:
        wb = (
            openpyxl.load_workbook(file_path)
            if os.path.exists(file_path)
            else openpyxl.Workbook()
        )
    except PermissionError:
        print(
            f"[PERMISSION ERROR] '{file_path}' is open in Excel. Close it and"
            " try again."
        )
        return

    # 1. Populate 'latest_entities'
    if TARGET_SHEET in wb.sheetnames:
        del wb[TARGET_SHEET]
    s_full = wb.create_sheet(TARGET_SHEET)
    s_full.append(
        ["Sr. No.", "Financial Year", "Procuring Entity", "APP Number", "APP URL"]
    )

    for i, (fin_year, entity, app_num, app_url) in enumerate(
        unique_records, start=1
    ):
        s_full.append([i, fin_year, entity, app_num, app_url])

    # 2. Populate 'latest_entities_only'
    if SIMPLE_SHEET in wb.sheetnames:
        del wb[SIMPLE_SHEET]
    s_simple = wb.create_sheet(SIMPLE_SHEET)
    s_simple.append(["Sr. No.", "Procuring Entity"])

    for i, (_, entity, _, _) in enumerate(unique_records, start=1):
        s_simple.append([i, entity])

    # 3. Preserve/Initialize 'Budget Totals'
    if BUDGET_SHEET not in wb.sheetnames:
        s_budget = wb.create_sheet(BUDGET_SHEET)
        s_budget.append(
            [
                "Sr. No.",
                "Procuring Entity",
                "Prequalification done",
                "Total Budget",
                "Status",
            ]
        )
        for i, (_, entity, _, _) in enumerate(unique_records, start=1):
            s_budget.append([i, entity, "", 0.0, "Pending"])
    else:
        s_budget = wb[BUDGET_SHEET]
        existing_data = {}
        for r in range(2, s_budget.max_row + 1):
            p_entity = s_budget.cell(row=r, column=2).value
            if p_entity:
                prequal = s_budget.cell(row=r, column=3).value or ""
                budget = s_budget.cell(row=r, column=4).value or 0.0
                status = s_budget.cell(row=r, column=5).value or "Pending"
                existing_data[normalize_key(p_entity)] = (
                    prequal,
                    budget,
                    status,
                )

        # Re-write sheet with clean active entities
        wb.remove(s_budget)
        s_budget = wb.create_sheet(BUDGET_SHEET)
        s_budget.append(
            [
                "Sr. No.",
                "Procuring Entity",
                "Prequalification done",
                "Total Budget",
                "Status",
            ]
        )

        for i, (_, entity, _, _) in enumerate(unique_records, start=1):
            key = normalize_key(entity)
            prequal, budget, status = existing_data.get(
                key, ("", 0.0, "Pending")
            )
            s_budget.append([i, entity, prequal, budget, status])

    safe_save_workbook(wb, file_path)
    print(
        f"Success! Written {len(unique_records)} clean records with resolved"
        f" URLs to '{EXCEL_FILE}'."
    )
    
def create_entity_template(sr_no: int, entity_name: str, app_no: str, scraped_segments: list) -> float:
    """Generates formatted entity Excel workbook and returns total budget."""
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
            item.get("open_agpo", ""),
            item.get("q1", 0.0),
            item.get("q2", 0.0),
            item.get("q3", 0.0),
            item.get("q4", 0.0),
        ])

        ws.cell(row=current_row, column=3).number_format = "#,##0.00"
        for col in range(5, 9):
            ws.cell(row=current_row, column=col).number_format = "#,##0.00"
        for col in range(1, 9):
            ws.cell(row=current_row, column=col).border = thin_border
        current_row += 1

    total_row = current_row
    ws.cell(row=total_row, column=1, value="TOTAL").font = font_bold
    ws.cell(row=total_row, column=3, value=f"=SUM(C{start_row}:C{total_row-1})").font = font_bold
    ws.cell(row=total_row, column=3).number_format = "#,##0.00"

    for col in range(5, 9):
        col_letter = get_column_letter(col)
        c = ws.cell(row=total_row, column=col, value=f"=SUM({col_letter}{start_row}:{col_letter}{total_row-1})")
        c.font = font_bold
        c.number_format = "#,##0.00"

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


def update_master_budget_totals(entity_name: str, calculated_total: float, status: str = "Done"):
    """Updates entity budget total and status in 'Budget Totals' sheet."""
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
            ws.cell(row=r, column=4, value=calculated_total).number_format = "#,##0.00"
            ws.cell(row=r, column=5, value=status)
            break

    safe_save_workbook(wb, EXCEL_FILE)


def should_skip_entity(entity_name: str, budget_df: pd.DataFrame) -> bool:
    """Checks if entity already has a calculated budget or is marked 'Done'."""
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


def run_deep_scrape(file_path=EXCEL_FILE, overwrite=False):
    """Deep Scrape with smart skip filters and resilient DOM rendering waits."""
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

    sync_and_rename_workbooks(df)

    with sync_playwright() as p:
        print("Launching browser for Deep Scrape...")
        browser = p.chromium.launch(headless=True)
        context = browser.new_context()
        page = context.new_page()

        for _, row in df.iterrows():
            sr_no = int(row["Sr. No."])
            entity = str(row["Procuring Entity"]).strip()
            app_no = str(row["APP Number"]).strip() if pd.notna(row.get("APP Number")) else ""
            app_url = str(row.get("APP URL", "")).strip()

            target_filepath = os.path.join(PLANS_DIR, f"{sr_no} {entity}.xlsx")

            # 1. Skip check based on local Excel budget status
            if not overwrite and should_skip_entity(entity, budget_df):
                print(f"[SKIP] Entity completed/has budget: [{sr_no}] {entity}")
                continue

            if os.path.exists(target_filepath) and not overwrite:
                print(f"[SKIP] Workbook exists: {target_filepath}")
                continue

            print(f"Deep Scraping [{sr_no}] {entity}...")
            try:
                # Navigate to Portal APP directory page
                page.goto("https://egpkenya.go.ke/public-app", wait_until="domcontentloaded")
                
                # Locate specific entity row
                link_locator = page.locator(
                    f"xpath=//tr[contains(., '{entity}') or contains(., '{app_no}')]//a"
                ).first

                if link_locator.count() == 0:
                    print(f"[WARNING] Could not locate link for [{sr_no}] {entity}")
                    continue

                # Click link safely & wait for network response
                with page.expect_navigation(wait_until="networkidle", timeout=12000):
                    link_locator.click()

                # Resilient DOM check for loaded table cells
                page.wait_for_selector("table tbody tr td", timeout=10000)

                segments = []
                s_rows = page.query_selector_all("table tbody tr")
                for s_row in s_rows:
                    cols = s_row.query_selector_all("td")
                    if len(cols) >= 3:
                        def parse_num(val_str):
                            clean = re.sub(r"[^\d.]", "", val_str)
                            return float(clean) if clean else 0.0

                        segments.append({
                            "sr_no": cols[0].inner_text().strip(),
                            "segment": cols[1].inner_text().strip(),
                            "total_cost": parse_num(cols[2].inner_text()),
                            "open_agpo": cols[3].inner_text().strip() if len(cols) > 3 else "",
                            "q1": parse_num(cols[4].inner_text()) if len(cols) > 4 else 0.0,
                            "q2": parse_num(cols[5].inner_text()) if len(cols) > 5 else 0.0,
                            "q3": parse_num(cols[6].inner_text()) if len(cols) > 6 else 0.0,
                            "q4": parse_num(cols[7].inner_text()) if len(cols) > 7 else 0.0,
                        })

                tot = create_entity_template(sr_no, entity, app_no, segments)
                update_master_budget_totals(entity, tot, status="Done")

            except Exception as e:
                print(f"[ERROR] Failed deep scrape for {entity}: {e}")

        browser.close()