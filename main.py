## main.py
import sys
import time
from address_tracker import update_address_tracker
from budget_tracker import update_budget_totals
from scrape_entities import (
    check_and_warn_locked_files,
    reconcile_budget_totals_from_workbooks,
    run_deep_scrape,
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


def run_step_1a():
    """Step 1a: Light Scrape e-GP Portal list into 'latest_entities' & 'latest_entities_only'"""
    if not ensure_no_file_locks():
        return False

    print("\n[STEP 1a] Running Light Scrape (e-GP portal list)...")
    start = time.time()
    try:
        scrape_to_latest_entities(EXCEL_FILE)
        print(f"-> Completed Step 1a in {round(time.time() - start, 2)}s\n")
        return True
    except PermissionError:
        print("\n[ERROR] PermissionError: Please close 'Procuring_Entities.xlsx' in Excel and retry.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Light Scraping failed: {e}\n")
        return False


def run_step_1b():
    """Step 1b: Deep Scrape individual entity plans & update workbook templates"""
    if not ensure_no_file_locks():
        return False

    print("\n[STEP 1b] Running Deep Scrape (Entity plans & segment details)...")
    start = time.time()
    try:
        run_deep_scrape(EXCEL_FILE)
        print(f"-> Completed Step 1b in {round(time.time() - start, 2)}s\n")
        return True
    except PermissionError:
        print("\n[ERROR] PermissionError: Open files in Excel prevented file updates.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Deep Scraping failed: {e}\n")
        return False


def run_step_2():
    """Step 2: Update 'Budget Totals' sheet"""
    if not ensure_no_file_locks():
        return False

    print("\n[STEP 2/4] Updating 'Budget Totals'...")
    start = time.time()
    try:
        update_budget_totals(EXCEL_FILE)
        print(f"-> Completed Step 2 in {round(time.time() - start, 2)}s\n")
        return True
    except PermissionError:
        print("\n[ERROR] PermissionError: Please close 'Procuring_Entities.xlsx' in Excel and retry.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Budget Totals update failed: {e}\n")
        return False


def run_step_3():
    """Step 3: Update 'addresses' sheet (Web lookup & Visit tracker)"""
    if not ensure_no_file_locks():
        return False

    print("\n[STEP 3/4] Updating 'addresses' (Visits Tracker)...")
    start = time.time()
    try:
        update_address_tracker(EXCEL_FILE)
        print(f"-> Completed Step 3 in {round(time.time() - start, 2)}s\n")
        return True
    except PermissionError:
        print("\n[ERROR] PermissionError: Please close 'Procuring_Entities.xlsx' in Excel and retry.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Addresses update failed: {e}\n")
        return False


def run_step_4():
    """Step 4: Push datasets to Google Workspace"""
    if not ensure_no_file_locks():
        return False

    print("\n[STEP 4/4] Synchronizing with Google Sheets...")
    start = time.time()
    try:
        sync_all_google_sheets()
        print(f"-> Completed Step 4 in {round(time.time() - start, 2)}s\n")
        return True
    except PermissionError:
        print("\n[ERROR] PermissionError: Please close 'Procuring_Entities.xlsx' in Excel and retry.\n")
        return False
    except Exception as e:
        print(f"\n[ERROR] Google Sheets sync failed: {e}\n")
        return False


def run_full_pipeline():
    """Executes full sequence: Light Scrape -> Deep Scrape -> Budget -> Addresses -> Sync"""
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
    if not run_step_2():
        return
    if not run_step_3():
        return
    if not run_step_4():
        return

    elapsed = round(time.time() - start_time, 2)
    print("==========================================")
    print(f" FULL PIPELINE COMPLETED IN {elapsed}s")
    print("==========================================\n")


def run_pipeline_skip_step_3():
    """Executes pipeline skipping addresses tracker"""
    if not ensure_no_file_locks():
        return

    start_time = time.time()
    print("\n==========================================")
    print(" RUNNING PIPELINE (SKIPPING STEP 3)")
    print("==========================================")

    if not run_step_1a():
        return
    if not run_step_1b():
        return
    if not run_step_2():
        return

    print("\n[STEP 3] SKIPPED: Updating 'addresses' (Visits Tracker)...\n")

    if not run_step_4():
        return

    elapsed = round(time.time() - start_time, 2)
    print("==========================================")
    print(f" PIPELINE COMPLETED IN {elapsed}s")
    print("==========================================\n")


def print_menu():
    print("==========================================")
    print("        PROCURING ENTITIES WORKFLOW       ")
    print("==========================================")
    print(" [1] Run FULL Pipeline (Light+Deep -> Budget -> Addresses -> Sync)")
    print(" [7] Run Pipeline (SKIP Step 3: Addresses)")
    print(" ----------------------------------------")
    print(" [2] Step 1a: Light Scrape ('latest_entities' & 'latest_entities_only')")
    print(" [3] Step 1b: Deep Scrape (Generate Entity Plans & Budget Totals)")
    print(" [4] Step 2: Update 'Budget Totals'")
    print(" [5] Step 3: Update 'addresses' (Search & Fill)")
    print(" [6] Step 4: Sync to Google Workspace")
    print(" [8] Reconcile 'Budget Totals' from entity workbooks (audit tool)")
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
            run_step_2()
        elif choice == "5":
            run_step_3()
        elif choice == "6":
            run_step_4()
        elif choice == "7":
            run_pipeline_skip_step_3()
        elif choice == "8":
            run_step_reconcile()
        elif choice == "0":
            print("\nExiting workflow runner. Goodbye!")
            sys.exit(0)
        else:
            print("\n[!] Invalid selection. Please enter a number from 0 to 7.\n")


if __name__ == "__main__":
    main()