import sys
import time
from egp import update_workflow_with_formatting
from scrape_entities import scrape_to_sheet2
from sync_gsuite import sync_sheet1_to_google_sheet

EXCEL_FILE = "Procuring_Entities_Workflow.xlsx"


def main():
    start_time = time.time()
    print("==========================================")
    print(" STARTING WORKFLOW PIPELINE")
    print("==========================================\n")

    # Step 1: Run Scraper -> Update Sheet2
    print("[STEP 1/2] Scraping e-GP portal to Sheet2...")
    try:
        scrape_to_sheet2()
    except Exception as e:
        print(f"\n[ERROR] Scraping failed: {e}")
        sys.exit(1)

    print("\n------------------------------------------\n")

    # Step 2: Run Updater -> Format Sheet1
    print("[STEP 2/2] Updating Sheet1 with new data & formatting...")
    try:
        update_workflow_with_formatting(EXCEL_FILE)
    except Exception as e:
        print(f"\n[ERROR] Updating Excel workflow failed: {e}")
        sys.exit(1)

    elapsed = round(time.time() - start_time, 2)
    print("\n==========================================")
    print(f" PIPELINE COMPLETED SUCCESSFULLY IN {elapsed}s")
    print("==========================================")

    print("[STEP 3/3] Syncing Sheet1 to Google Sheet...")
    sync_sheet1_to_google_sheet()


if __name__ == "__main__":
    main()