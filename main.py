## main.py
import sys
import time
from address_tracker import update_address_tracker
from budget_tracker import update_budget_totals
from scrape_entities import (
    check_and_warn_locked_files,
    reconcile_budget_totals_from_workbooks,
    run_deep_scrape,
    run_item_details_enrichment,
    scrape_to_latest_entities,
)
from sync_gsuite import sync_all_google_sheets

EXCEL_FILE = "Procuring_Entities.xlsx"


def ensure_no_file_locks() -> bool:
    """
    Checks for open Excel lock files (~$*.xlsx).
    Prompts the user to close Excel and press Enter to re-check, or 'q' to abort.
    """
    while not check_and_warn_locked_files():
        choice = (
            input("Press [Enter] to re-check after closing Excel, or type 'q' to cancel: ")
            .strip()
            .lower()
        )
        if choice == "q":
            print("[!] Operation canceled due to active Excel file locks.\n")
            return False
    return True


def run_step_1a():
    """Step 1: Light Scrape e-GP Portal list into 'latest_entities' & 'latest_entities_only'"""
    if not ensure_no_file_locks():
        return False

    print("\n[STEP 1] Running Light Scrape (e-GP portal list)...")
    start = time.time()
    try:
        scrape_to_latest_entities(EXCEL_FILE)
        print(f"-> Completed Step 1 in {round(time.time() - start, 2)}s\n")
        return True
    except PermissionError:
        print("\n[ERROR] PermissionError: Please close 'Procuring_Entities.xlsx' in Excel and retry.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Light Scraping failed: {e}\n")
        return False


def run_step_1b():
    """Step 2: Deep Scrape individual entity plans, then enrich each with
    OPEN/AGPO + Q1-Q4 via Item Details. The enrichment is treated as a
    second pass of the same step rather than a separate pipeline stage,
    since its only job is to fill in columns Deep Scrape deliberately
    leaves blank on the first pass.
    """
    if not ensure_no_file_locks():
        return False

    print("\n[STEP 2] Running Deep Scrape (Entity plans & segment details)...")
    start = time.time()
    try:
        run_deep_scrape(EXCEL_FILE)
        print(f"-> Deep Scrape completed in {round(time.time() - start, 2)}s")
    except PermissionError:
        print("\n[ERROR] PermissionError: Open files in Excel prevented file updates.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Deep Scraping failed: {e}\n")
        return False

    print("\n[STEP 2b] Enriching entity plans (OPEN/AGPO + Q1-Q4)...")
    start = time.time()
    try:
        run_item_details_enrichment(EXCEL_FILE)
        print(f"-> Completed Step 2 (Deep Scrape + Enrichment) in {round(time.time() - start, 2)}s\n")
        return True
    except PermissionError:
        print("\n[ERROR] PermissionError: Please close 'Procuring_Entities.xlsx' in Excel and retry.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Item Details enrichment failed: {e}\n")
        return False


def run_step_2_budget():
    """Step 3: Update 'Budget Totals' sheet"""
    if not ensure_no_file_locks():
        return False

    print("\n[STEP 3] Updating 'Budget Totals'...")
    start = time.time()
    try:
        update_budget_totals(EXCEL_FILE)
        print(f"-> Completed Step 3 in {round(time.time() - start, 2)}s\n")
        return True
    except PermissionError:
        print("\n[ERROR] PermissionError: Please close 'Procuring_Entities.xlsx' in Excel and retry.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Budget Totals update failed: {e}\n")
        return False


def run_step_3_addresses():
    """Step 4: Update 'addresses' sheet (Web lookup & Visit tracker)"""
    if not ensure_no_file_locks():
        return False

    print("\n[STEP 4] Updating 'addresses' (Visits Tracker)...")
    start = time.time()
    try:
        update_address_tracker(EXCEL_FILE)
        print(f"-> Completed Step 4 in {round(time.time() - start, 2)}s\n")
        return True
    except PermissionError:
        print("\n[ERROR] PermissionError: Please close 'Procuring_Entities.xlsx' in Excel and retry.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Addresses update failed: {e}\n")
        return False


def run_step_4_sync():
    """Step 5: Push datasets to Google Workspace"""
    if not ensure_no_file_locks():
        return False

    print("\n[STEP 5] Synchronizing with Google Sheets...")
    start = time.time()
    try:
        sync_all_google_sheets()
        print(f"-> Completed Step 5 in {round(time.time() - start, 2)}s\n")
        return True
    except PermissionError:
        print("\n[ERROR] PermissionError: Please close 'Procuring_Entities.xlsx' in Excel and retry.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Google Sheets sync failed: {e}\n")
        return False


def run_step_reconcile():
    """Reconciles 'Budget Totals' by re-reading every entity workbook directly."""
    if not ensure_no_file_locks():
        return False

    print("\n[RECONCILE] Recomputing Budget Totals from entity workbooks...")
    start = time.time()
    try:
        reconcile_budget_totals_from_workbooks()
        print(f"-> Completed reconcile in {round(time.time() - start, 2)}s\n")
        return True
    except PermissionError:
        print("\n[ERROR] PermissionError: Please close 'Procuring_Entities.xlsx' in Excel and retry.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Reconcile failed: {e}\n")
        return False


def run_full_pipeline():
    """Executes full sequence: Light Scrape -> Deep Scrape+Enrichment -> Budget -> Addresses -> Sync"""
    if not ensure_no_file_locks():
        return

    start_time = time.time()
    print("\n==========================================")
    print(" RUNNING FULL WORKFLOW PIPELINE")
    print("==========================================")

    if not run_step_1a():
        return
    if not run_step_1b():
        return
    if not run_step_2_budget():
        return
    if not run_step_3_addresses():
        return
    if not run_step_4_sync():
        return

    elapsed = round(time.time() - start_time, 2)
    print("==========================================")
    print(f" FULL PIPELINE COMPLETED IN {elapsed}s")
    print("==========================================\n")


def print_menu():
    print("==========================================")
    print("        PROCURING ENTITIES WORKFLOW       ")
    print("==========================================")
    print(" [1] Run FULL Pipeline (all steps below, in order)")
    print(" ----------------------------------------")
    print(" [2] Step 1: Light Scrape ('latest_entities' & 'latest_entities_only')")
    print(" [3] Step 2: Deep Scrape + Enrichment (Entity Plans, OPEN/AGPO, Q1-Q4)")
    print(" [4] Step 3: Update 'Budget Totals'")
    print(" [5] Step 4: Update 'addresses' (Search & Fill)")
    print(" [6] Step 5: Sync to Google Workspace")
    print(" ----------------------------------------")
    print(" [7] Reconcile 'Budget Totals' from entity workbooks (audit tool)")
    print(" ----------------------------------------")
    print(" [0] Exit")
    print("==========================================")


def main():
    while True:
        print_menu()
        choice = input("Select an option [0-7]: ").strip()

        if choice == "1":
            run_full_pipeline()
        elif choice == "2":
            run_step_1a()
        elif choice == "3":
            run_step_1b()
        elif choice == "4":
            run_step_2_budget()
        elif choice == "5":
            run_step_3_addresses()
        elif choice == "6":
            run_step_4_sync()
        elif choice == "7":
            run_step_reconcile()
        elif choice == "0":
            print("\nExiting workflow runner. Goodbye!")
            sys.exit(0)
        else:
            print("\n[!] Invalid selection. Please enter a number from 0 to 7.\n")


if __name__ == "__main__":
    main()