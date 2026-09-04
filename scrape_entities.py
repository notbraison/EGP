import time
import openpyxl
from playwright.sync_api import sync_playwright

EXCEL_FILE = "Procuring_Entities_Workflow.xlsx"


def scrape_to_sheet2():
    scraped_data = []

    with sync_playwright() as p:
        print("Launching browser...")
        browser = p.chromium.launch(headless=True)
        context = browser.new_context(viewport={"width": 1280, "height": 800})
        page = context.new_page()

        print("Navigating to e-GP Public APP page...")
        page.goto("https://egpkenya.go.ke/public-app", wait_until="networkidle")

        current_page = 1

        while True:
            print(f"Scraping Page {current_page}...")
            page.wait_for_selector("table tbody tr", timeout=10000)
            time.sleep(1)

            rows = page.query_selector_all("table tbody tr")
            page_has_data = False

            for row in rows:
                cols = row.query_selector_all("td")
                if len(cols) >= 4:
                    sr_no = cols[0].inner_text().strip()
                    fin_year = cols[1].inner_text().strip()
                    entity = cols[2].inner_text().strip()
                    app_num = cols[3].inner_text().strip()

                    if entity:
                        scraped_data.append((fin_year, entity, app_num))
                        page_has_data = True

            # Check next page button state
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
                time.sleep(2)
            else:
                print("Reached the final page.")
                break

        browser.close()

    if not scraped_data:
        print("No records scraped. Sheet2 was not updated.")
        return

    # Deduplicate while preserving original order
    seen = set()
    unique_records = []
    for fin_year, entity, app_num in scraped_data:
        key = (entity.upper(), app_num.upper())
        if key not in seen:
            seen.add(key)
            unique_records.append((fin_year, entity, app_num))

    # Open workbook and overwrite Sheet2
    wb = openpyxl.load_workbook(EXCEL_FILE)

    if "Sheet2" in wb.sheetnames:
        del wb["Sheet2"]
    s2 = wb.create_sheet("Sheet2")

    # Add standard header row
    s2.append(["Sr. No.", "Financial Year", "Procuring Entity", "APP Number"])

    # Write scraped rows with fresh serial numbering
    for i, (fin_year, entity, app_num) in enumerate(unique_records, start=1):
        s2.append([i, fin_year, entity, app_num])

    wb.save(EXCEL_FILE)
    print(
        f"\nSuccess! Written {len(unique_records)} records directly into Sheet2"
        f" of '{EXCEL_FILE}'."
    )


if __name__ == "__main__":
    scrape_to_sheet2()