"""Detailed inspector for the sample Excel attachment."""

from pathlib import Path
import openpyxl


def inspect_excel(filepath: Path):
    print("=" * 80)
    print(f"INSPECTING EXCEL FILE: {filepath.name}")
    print("=" * 80)

    wb = openpyxl.load_workbook(filepath, data_only=False)
    print(f"Sheet names: {wb.sheetnames}")
    
    # Metadata / Properties
    props = wb.properties
    print(f"Title: {props.title}, Creator: {props.creator}, Last Modified By: {props.lastModifiedBy}")
    print(f"Created: {props.created}, Modified: {props.modified}")

    for sheetname in wb.sheetnames:
        ws = wb[sheetname]
        print("\n" + "-" * 50)
        print(f"SHEET: '{sheetname}'")
        print("-" * 50)
        print(f"Dimensions: {ws.dimensions} (max_row={ws.max_row}, max_column={ws.max_column})")
        print(f"Merged cell ranges: {[str(m) for m in ws.merged_cells.ranges]}")
        
        # Hidden rows / cols
        hidden_cols = [col_letter for col_letter, col_dim in ws.column_dimensions.items() if col_dim.hidden]
        hidden_rows = [row_idx for row_idx, row_dim in ws.row_dimensions.items() if row_dim.hidden]
        print(f"Hidden columns: {hidden_cols}")
        print(f"Hidden rows: {hidden_rows}")

        # Inspect all rows up to max_row (or first 30 rows)
        print(f"\n--- Row inspection (first 30 rows) ---")
        populated_row_count = 0
        empty_row_count = 0
        formulas_found = []

        for r_idx in range(1, ws.max_row + 1):
            row_vals = [ws.cell(r_idx, c_idx).value for c_idx in range(1, ws.max_column + 1)]
            # check if empty
            if all(v is None or str(v).strip() == "" for v in row_vals):
                empty_row_count += 1
            else:
                populated_row_count += 1

            for c_idx, val in enumerate(row_vals, start=1):
                if isinstance(val, str) and val.startswith("="):
                    formulas_found.append((f"R{r_idx}C{c_idx}", val))

            if r_idx <= 25:
                # format row for printing
                filtered_vals = [v for v in row_vals if v is not None]
                data_types = [type(v).__name__ for v in row_vals if v is not None]
                print(f"Row {r_idx:2d} ({len(filtered_vals)} populated cells): {row_vals[:12]}")
                if r_idx in (1, 2, 3, 4, 5):
                    print(f"       Types: {data_types[:12]}")

        print(f"\nTotal populated rows: {populated_row_count}")
        print(f"Total empty rows: {empty_row_count}")
        print(f"Formulas found: {len(formulas_found)}")
        if formulas_found:
            for cell_ref, fmla in formulas_found[:10]:
                print(f"  {cell_ref}: {fmla}")


if __name__ == "__main__":
    p = Path("tests/fixtures/attachments/Alokasi Logitech BIC1 to iBox BCA XE05 7 September.xlsx")
    inspect_excel(p)
