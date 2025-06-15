#!/usr/bin/env python3
import pandas as pd
from openpyxl.utils import get_column_letter

# ——— Configuration ———
INPUT_FILE        = "files/Discount-Detail-Report-12_1_2024-6_14_2025.xlsx"
OUTPUT_FILE       = "weedmaps_report.xlsx"
DESC_FILTER       = "Weedmaps 2% online order"
PERCENT_THRESHOLD = 2.0

KEEP_COLS = [
    "Location Name",
    "Order ID",
    "Order Time",
    "Customer Name",
    "Product Name",
    "Unit Price",
    "Gross Sales",
    "Discounted Amount",
    "Discount Percent",
    "Net Sales",
    "Discount Name",
    "Budtender Name",
    "Discount Approved By",
]
# ————————————

def main():
    # 1. Load using row 5 as header (zero-based 4)
    df = pd.read_excel(INPUT_FILE, header=4, engine="openpyxl")

    # 2. Normalize & convert Discount Percent to float
    df["Discount Percent"] = (
        df["Discount Percent"]
          .astype(str)
          .str.rstrip("%")
          .replace("", "0")
          .astype(float)
    )

    # 3. Filter by description & threshold
    mask_desc = (
        df["Discount Description"]
          .astype(str)
          .str.contains(DESC_FILTER, case=False, na=False)
    )
    mask_pct = df["Discount Percent"] > PERCENT_THRESHOLD
    filtered = df[mask_desc & mask_pct].copy()

    # 4. Sort descending
    filtered.sort_values("Discount Percent", ascending=False, inplace=True)

    # 5. Compute the two counts
    filtered["Usage Count"] = (
        filtered.groupby("Customer Name")["Customer Name"]
                .transform("count")
    )
    filtered["Approver Count"] = (
        filtered.groupby("Discount Approved By")["Discount Approved By"]
                .transform("count")
    )

    # 6. Select only the columns you want + our new counts
    final_cols = KEEP_COLS + ["Usage Count", "Approver Count"]
    filtered = filtered[final_cols]

    # 7. Identify “abuse” cases (customers who used it > once)
    abuse = filtered[filtered["Usage Count"] > 1]

    # 8. Write both sheets into one workbook; then auto-fit and freeze header
    with pd.ExcelWriter(OUTPUT_FILE, engine="openpyxl") as writer:
        filtered.to_excel(writer, sheet_name="All Weedmaps >2%", index=False)
        abuse.to_excel(writer, sheet_name="Abuse (Multiple Uses)", index=False)

        for sheet in writer.sheets.values():
            max_row = sheet.max_row
            max_col = sheet.max_column

            # Auto-fit columns
            for col_idx in range(1, max_col + 1):
                col_letter = get_column_letter(col_idx)
                max_length = 0
                for row_idx in range(1, max_row + 1):
                    val = sheet.cell(row=row_idx, column=col_idx).value
                    if val is not None:
                        val_length = len(str(val))
                        if val_length > max_length:
                            max_length = val_length
                sheet.column_dimensions[col_letter].width = max_length + 2

            # Freeze the top header row
            sheet.freeze_panes = "A2"

    print(f"✔ Report written to '{OUTPUT_FILE}'")
    print(f"  • Total qualifying rows: {len(filtered)}")
    print(f"  • Abuse rows (Usage Count > 1): {len(abuse)}")

if __name__ == "__main__":
    main()
