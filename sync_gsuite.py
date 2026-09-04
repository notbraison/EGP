import json
import urllib.request
import openpyxl

EXCEL_FILE = "Procuring_Entities_Workflow.xlsx"
# Replace with the URL copied from Step 1
WEB_APP_URL = "https://script.google.com/macros/s/AKfycbwYm_HzK17LPJ3TerMslhBaYWXke887YW5NMc3nc8K-razyviUQPwu-7rmXAZXNg_UG/exec"


def sync_sheet1_to_google_sheet():
    wb = openpyxl.load_workbook(EXCEL_FILE, data_only=True)
    if "Sheet1" not in wb.sheetnames:
        raise ValueError("Sheet1 not found in local Excel file.")

    s1 = wb["Sheet1"]
    data = []

    # Read all non-empty rows from Sheet1
    for row in s1.iter_rows(values_only=True):
        # Stop reading if entire row is empty
        if any(cell is not None for cell in row):
            # Convert None cells to empty strings for JSON compatibility
            clean_row = [cell if cell is not None else "" for cell in row[:4]]
            data.append(clean_row)

    if not data:
        print("No data found in Sheet1.")
        return

    payload = json.dumps({"data": data}).encode("utf-8")
    req = urllib.request.Request(
        WEB_APP_URL,
        data=payload,
        headers={"Content-Type": "application/json"},
        method="POST",
    )

    print("Sending Sheet1 data to Google Sheet...")
    with urllib.request.urlopen(req) as response:
        result = json.loads(response.read().decode())
        if result.get("status") == "success":
            print(
                f"Successfully updated 'Budget Totals' tab with {len(data)}"
                " rows!"
            )
        else:
            print("Failed to update Google Sheet.")


if __name__ == "__main__":
    sync_sheet1_to_google_sheet()