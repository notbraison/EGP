# DEV.md

# Initialize & Activate Virtual Environment
python -m venv venv
venv\Scripts\activate

# Install Core Dependencies
pip install openpyxl playwright pandas
playwright install chromium

# Run Pipeline
python main.py

# Deactivate Environment
deactivate


# File Architecture

Procuring_Entities.xlsx: Local master workbook containing four sheets:
  - Budget Totals: Main budget tracking sheet structured as: `Sr. No.` | `Procuring Entity` | `Prequalification done` | `Total Budget` | `Status`.
  - latest_entities: Master scraped dataset containing Sr. No., Financial Year, Procuring Entity, APP Number, and target APP URL for deep scraping.
  - latest_entities_only: Streamlined light-scrape dataset listing only Sr. No. and Procuring Entity.
  - addresses: Master database for physical addresses, contact info, assigned officers, and visit tracking.

./entity_plans/: Dedicated directory holding individual formatted Excel workbooks per approved procuring entity (e.g., `./entity_plans/20 PEST CONTROL PRODUCTS BOARD PP.xlsx`).

main.py: Pipeline orchestrator with CLI menu supporting Full Pipeline, Light Scrape (Step 1a), Deep Scrape (Step 1b), selective execution, timing, and error handling.

scrape_entities.py: 
  - Light Scrape & Filtering: Extracts e-GP Kenya portal records, immediately filters out blacklisted county/non-paying entities, and writes approved entities (~93 master targets) to `latest_entities`, `latest_entities_only`, and `Budget Totals`.
  - Dynamic Workbook Sync: Matches existing entity files strictly by alphanumeric name keys and automatically re-indexes filename prefixes (`Sr. No.`) when portal order changes.
  - Deep Scrape: Checks local budget statuses to skip pre-completed or manual entries, handles dynamic JavaScript click/navigation events on the e-GP portal, parses procurement segment tables, creates formatted individual `.xlsx` plan files, and updates `Budget Totals`.

budget_tracker.py: Syncs `latest_entities` with `Budget Totals`, preserving manual entries in `Prequalification done`, `Total Budget`, and `Status` columns.

address_tracker.py: Syncs `latest_entities` with `addresses`, preserving manual contact and visit details while appending new entities.

sync_gsuite.py: Pushes local workbook datasets live to Google Sheets via Web App endpoints, preserving the expanded schema with `Prequalification done`.


# Key Techniques Used

Headless Browser Scraping (Playwright): Bypasses WAF protections, CSRF token checks, and 403 payload rejections. Handles client-side JavaScript rendering and inline navigation (`href="javascript:void(0)"`) by waiting for `networkidle` state and loaded table cell elements (`table tbody tr td`).

Dual-Layer Blacklist Filtering: Automatically drops county-level entities, county assemblies, and non-paying authorities during Light Scrape using:
  1. An explicit `BLACKLIST_ENTITIES` set containing 42 target entities.
  2. Pattern matching rules that catch `"COUNTY"` keywords or leading county numeric codes (e.g., `4815 KAKAMEGA`).

Smart Execution & Delta Scraping: `should_skip_entity()` checks `Budget Totals` before triggering browser runs. If an entity is marked `Status == "Done"` or has a calculated `Total Budget > 0`, deep scraping is skipped to save time and reduce server hits.

Index-Resilient File Renaming: Normalizes entity names using `"".join(c for c in name.upper() if c.isalnum())` to re-index local entity filenames in `./entity_plans/` when portal ordering shifts, preventing duplicate file creation and preserving sheet content.

Excel Template Engineering & Formula Injection: Programmatically constructs individual entity workbooks with merged top-level banners, standard headers, `#,#0.00` currency cell formatting, thin borders, auto-adjusted column widths, and live `=SUM()` dynamic formulas via `openpyxl`.

Lightweight Web App Bridge: Uses standard Python `urllib.request` to POST JSON to a Google Apps Script `doPost(e)` endpoint, eliminating complex OAuth2 credentials and third-party login prompts.

Selective Canvas Clearing: Uses `sheet.clearContents()` in Google Apps Script to update dataset values without destroying title banners, column widths, or header fills.


# Crucial Considerations for Future Builds

Excel File Lock: `Procuring_Entities.xlsx` and any individual `.xlsx` files inside `./entity_plans/` must be closed in Microsoft Excel before running `main.py` or `scrape_entities.py`; otherwise, `openpyxl` or `os.rename()` will throw a `PermissionError`.

Blacklist Adjustments: If an entity needs to be restored or removed from scraping, update `BLACKLIST_ENTITIES` and `is_blacklisted()` inside `scrape_entities.py`.

Dynamic AJAX Loading: The e-GP portal loads plan segments asynchronously via client-side JavaScript. Deep scrape logic must wait for network responses (`wait_until="networkidle"`) and cell elements (`table tbody tr td`) rather than generic table containers.

Apps Script Deployment Updates: Whenever modifying code in Google Apps Script, you must publish a New Version under Deploy → Manage deployments for changes to take effect on the Web App URL.