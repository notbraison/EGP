from copy import copy
import re
import time
import random
import urllib.parse
import openpyxl
from openpyxl.styles import Border, PatternFill
from playwright.sync_api import sync_playwright
from scrape_entities import is_blacklisted

EXCEL_FILE = "Procuring_Entities.xlsx"
ENTITIES_SHEET = "latest_entities"
ADDRESSES_SHEET = "addresses"

HEADERS = [
    "Sr. No.", "Procuring Entity", "Physical Address (Building/Street)", 
    "Town/Area", "County", "Phone", "Email", "ASSIGNED", "VISIT"
]

KENYA_COUNTIES = [
    "Nairobi", "Uasin Gishu", "Nyandarua", "Kiambu", "Nandi", "Nyeri",
    "Kirinyaga", "Busia", "Kisumu", "Bungoma", "Machakos", "Migori",
    "Bomet", "West Pokot", "Baringo", "Kakamega", "Homa Bay", "Mombasa",
    "Kajiado", "Nakuru", "Murang'a", "Trans Nzoia", "Kilifi", "Kwale",
    "Taita Taveta", "Garissa", "Wajir", "Mandera", "Marsabit", "Isiolo",
    "Meru", "Tharaka-Nithi", "Embu", "Kitui", "Makueni", "Turkana",
    "Samburu", "Elgeyo Marakwet", "Kericho", "Laikipia", "Narok", "Vihiga",
    "Siaya", "Kisii", "Nyamira", "Tana River", "Lamu",
]



def clean_addresses_sheet(file_path=EXCEL_FILE):
    wb = openpyxl.load_workbook(file_path)
    if "addresses" not in wb.sheetnames:
        return
        
    ws = wb["addresses"]
    rows_to_keep = []
    
    # Preserve header row
    headers = [cell.value for cell in ws[1]]
    rows_to_keep.append(headers)
    
    # Filter out blacklisted entity rows
    for row in ws.iter_rows(min_row=2, values_only=True):
        entity_name = row[1]  # Assuming Procuring Entity is Column 2
        if entity_name and not is_blacklisted(entity_name):
            rows_to_keep.append(row)
            
    # Clear and rewrite sheet
    ws.delete_rows(1, ws.max_row)
    for r in rows_to_keep:
        ws.append(r)
        
    wb.save(file_path)

def normalize_key(text):
    if not text:
        return ""
    clean = str(text).replace("\xa0", " ").strip().upper()
    return "".join(c for c in clean if c.isalnum())


def extract_contact_info(search_text):
    info = {"address": "", "town": "", "county": "", "phone": "", "email": ""}
    if not search_text:
        return info

    email_match = re.search(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", search_text)
    if email_match:
        info["email"] = email_match.group(0).lower()

    phone_match = re.search(r"(?:\+?254|0)(?:\s?\d{2,3}){1,3}\s?\d{3,4}\b", search_text)
    if phone_match:
        raw_phone = phone_match.group(0).strip()
        info["phone"] = f"'{raw_phone}" if not raw_phone.startswith("'") else raw_phone

    for county in KENYA_COUNTIES:
        if re.search(r"\b" + re.escape(county) + r"\b", search_text, re.IGNORECASE):
            info["county"] = county
            break

    road_match = re.search(
        r"([A-Za-z0-9\s,-]+(?:Road|Street|Avenue|Way|Building|House|Towers|Plaza|Campus))",
        search_text, re.IGNORECASE,
    )
    if road_match:
        info["address"] = road_match.group(0).strip()

    return info


def perform_single_search(entity, page):
    """Executes a single web search using Bing to avoid DuckDuckGo IP blocks."""
    clean_query = re.sub(r'^\d+\s*-?\s*', '', entity)
    query = f'"{clean_query}" physical address contact phone email county Kenya'
    search_url = f"https://www.bing.com/search?q={urllib.parse.quote(query)}"

    max_retries = 2
    for attempt in range(max_retries):
        try:
            page.goto(search_url, wait_until="domcontentloaded", timeout=20000)
            
            try:
                page.wait_for_selector("#b_results", timeout=5000)
            except Exception:
                pass  # Fallback to evaluating whole body if specific ID isn't found
            
            combined_text = page.evaluate("document.body.innerText")
            return extract_contact_info(combined_text)

        except Exception as e:
            print(f"    [Attempt {attempt + 1}/{max_retries}] Failed search: {e}")
            if attempt < max_retries - 1:
                print("    Sleeping for 5 seconds before retrying...")
                time.sleep(5)
            
    return {"address": "", "town": "", "county": "", "phone": "", "email": ""}


def update_address_tracker(file_path=EXCEL_FILE):
    wb = openpyxl.load_workbook(file_path)

    if ENTITIES_SHEET not in wb.sheetnames:
        raise ValueError(f"'{ENTITIES_SHEET}' must exist with scraped entity data.")
    s2 = wb[ENTITIES_SHEET]

    if ADDRESSES_SHEET not in wb.sheetnames:
        s3 = wb.create_sheet(ADDRESSES_SHEET)
        s3.append(HEADERS)
    else:
        s3 = wb[ADDRESSES_SHEET]

    # 1. Index existing address records
    existing_data = {}
    if s3.max_row >= 2:
        for r in range(2, s3.max_row + 1):
            entity = s3.cell(row=r, column=2).value
            if entity:
                key = normalize_key(entity)
                existing_data[key] = {
                    "address": s3.cell(row=r, column=3).value or "",
                    "town": s3.cell(row=r, column=4).value or "",
                    "county": s3.cell(row=r, column=5).value or "",
                    "phone": s3.cell(row=r, column=6).value or "",
                    "email": s3.cell(row=r, column=7).value or "",
                    "assigned": s3.cell(row=r, column=8).value or "",
                    "visit": s3.cell(row=r, column=9).value or "",
                }

    scraped_entities = []
    for r in range(2, s2.max_row + 1):
        entity = s2.cell(row=r, column=3).value
        if entity:
            scraped_entities.append(str(entity).replace("\xa0", " ").strip())

    template_cells = [s3.cell(row=2, column=c) for c in range(1, 10)] if s3.max_row >= 2 else None

    # 2. Iterate, Search, and Save Incrementally
    print("Syncing entities and checking for missing contact info...")
    
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = context.new_page()
        page.route("**/*.{png,jpg,jpeg,svg,gif,woff,css}", lambda route: route.abort())

        searches_performed = 0

        for i, entity_name in enumerate(scraped_entities, start=1):
            row_idx = i + 1
            key = normalize_key(entity_name)
            record = existing_data.get(key, {})

            address = record.get("address", "")
            town = record.get("town", "")
            county = record.get("county", "")
            phone = record.get("phone", "")
            email = record.get("email", "")

            needs_search = not all([address, phone, email, county])

            # If data is missing, perform search and update variables immediately
            if needs_search:
                searches_performed += 1
                print(f"  [{i}/{len(scraped_entities)}] Searching web for: {entity_name}...")
                
                searched_data = perform_single_search(entity_name, page)
                address = address or searched_data.get("address", "")
                town = town or searched_data.get("town", "")
                county = county or searched_data.get("county", "")
                phone = phone or searched_data.get("phone", "")
                email = email or searched_data.get("email", "")

            if phone and not str(phone).startswith("'"):
                phone = f"'{phone}"

            row_values = [
                i, entity_name, address, town, county, phone, email, 
                record.get("assigned", ""), record.get("visit", "")
            ]

            # Write row to memory
            for col_idx, val in enumerate(row_values, start=1):
                cell = s3.cell(row=row_idx, column=col_idx)
                cell.value = val
                if template_cells:
                    tmpl = template_cells[col_idx - 1]
                    cell.border = copy(tmpl.border)
                    cell.alignment = copy(tmpl.alignment)
                    cell.font = copy(tmpl.font)

            # SAVE IMMEDIATELY if a search was performed so no data is lost on interrupt
            if needs_search:
                try:
                    wb.save(file_path)
                except Exception as e:
                    print(f"    [!] Could not save file (is it open in Excel?): {e}")

                # Rate limiting delay
                if searches_performed % 10 == 0:
                    print("  [Rate Limit Guard] Cooling down for 10 seconds...")
                    time.sleep(10)
                else:
                    time.sleep(random.uniform(3.0, 6.0))

        browser.close()

    # 3. Clear dead rows at the bottom
    last_valid_row = len(scraped_entities) + 1
    if s3.max_row > last_valid_row:
        for r in range(last_valid_row + 1, s3.max_row + 1):
            for c in range(1, 10):
                cell = s3.cell(row=r, column=c)
                cell.value = None
                cell.border = Border()
                cell.fill = PatternFill(fill_type=None)
        
        wb.save(file_path)

    print(f"\nUpdated '{ADDRESSES_SHEET}' with {len(scraped_entities)} entities.")


if __name__ == "__main__":
    update_address_tracker()