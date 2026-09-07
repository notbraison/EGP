from copy import copy
import openpyxl
from openpyxl.styles import Border, Font, PatternFill

EXCEL_FILE = "Procuring_Entities.xlsx"
BUDGET_SHEET = "Budget Totals"
ENTITIES_SHEET = "latest_entities"


def update_budget_totals(file_path=EXCEL_FILE):
    wb = openpyxl.load_workbook(file_path)

    if BUDGET_SHEET not in wb.sheetnames or ENTITIES_SHEET not in wb.sheetnames:
        raise ValueError(
            f"Workbook must contain both '{BUDGET_SHEET}' and"
            f" '{ENTITIES_SHEET}'"
        )

    s1 = wb[BUDGET_SHEET]
    s2 = wb[ENTITIES_SHEET]

    # Index existing budgets, prequalifications, and statuses across 5 columns
    existing_map = {}
    for r in range(3, s1.max_row + 1):
        entity = s1.cell(row=r, column=2).value
        prequal = s1.cell(row=r, column=3).value
        budget = s1.cell(row=r, column=4).value
        status = s1.cell(row=r, column=5).value

        if entity:
            clean_entity = str(entity).replace("\xa0", " ").strip()
            key = "".join(e for e in clean_entity.upper() if e.isalnum())
            existing_map[key] = {
                "prequal": prequal,
                "budget": budget,
                "status": status,
            }

    new_entities = []
    for r in range(2, s2.max_row + 1):
        sr_no = s2.cell(row=r, column=1).value
        entity = s2.cell(row=r, column=3).value

        if entity:
            clean_entity = str(entity).replace("\xa0", " ").strip()
            new_entities.append((sr_no, clean_entity))

    # Template cells span 5 columns
    template_row = [s1.cell(row=3, column=c) for c in range(1, 6)]

    for i, (sr_no, entity_name) in enumerate(new_entities):
        row_idx = i + 3
        key = "".join(e for e in entity_name.upper() if e.isalnum())
        matched = existing_map.get(key, {"prequal": "", "budget": "", "status": ""})

        status_val = (
            str(matched["status"]).strip().lower() if matched["status"] else ""
        )

        s1.cell(row=row_idx, column=1).value = sr_no or (i + 1)
        s1.cell(row=row_idx, column=2).value = entity_name
        s1.cell(row=row_idx, column=3).value = matched.get("prequal", "")
        s1.cell(row=row_idx, column=4).value = matched["budget"]
        s1.cell(row=row_idx, column=5).value = matched["status"]

        for c in range(1, 6):
            cell = s1.cell(row=row_idx, column=c)
            tmpl = template_row[c - 1]

            cell.border = copy(tmpl.border)
            cell.alignment = copy(tmpl.alignment)
            cell.font = copy(tmpl.font)
            cell.number_format = copy(tmpl.number_format)

        # Status styling applied to Column 5
        status_cell = s1.cell(row=row_idx, column=5)
        if status_val == "done":
            status_cell.fill = PatternFill(
                start_color="E2EFDA", end_color="E2EFDA", fill_type="solid"
            )
            status_cell.font = Font(
                name=template_row[4].font.name or "Arial",
                size=template_row[4].font.size or 11,
                bold=True,
                color="375623",
            )
        else:
            status_cell.fill = PatternFill(fill_type=None)
            status_cell.font = Font(
                name=template_row[4].font.name or "Arial",
                size=template_row[4].font.size or 11,
                bold=False,
                color="000000",
            )

    # Clean up orphan trailing rows across all 5 columns
    last_valid_row = len(new_entities) + 2
    if s1.max_row > last_valid_row:
        for r in range(last_valid_row + 1, s1.max_row + 1):
            for c in range(1, 6):
                cell = s1.cell(row=r, column=c)
                cell.value = None
                cell.border = Border()
                cell.fill = PatternFill(fill_type=None)

    wb.save(file_path)
    print(f"Updated '{BUDGET_SHEET}' with {len(new_entities)} entities.")


if __name__ == "__main__":
    update_budget_totals()