import os
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import CellIsRule

def ensure_dir_exists(directory):
    """Creates the directory if it doesn't already exist."""
    if not os.path.exists(directory):
        os.makedirs(directory)

def apply_conditional_formatting(workbook, sheet_name, margin_col_index):
    """
    Applies conditional formatting to the Margin column:
    - Red fill if margin < 0.40
    - Green fill if margin >= 0.40
    """
    ws = workbook[sheet_name]
    max_row = ws.max_row
    margin_col_letter = get_column_letter(margin_col_index)

    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    # “Less than 0.4” → Red
    ws.conditional_formatting.add(
        f"{margin_col_letter}2:{margin_col_letter}{max_row}",
        CellIsRule(operator='lessThan', formula=['0.4'], fill=red_fill)
    )
    # “Greater or equal 0.4” → Green
    ws.conditional_formatting.add(
        f"{margin_col_letter}2:{margin_col_letter}{max_row}",
        CellIsRule(operator='greaterThanOrEqual', formula=['0.4'], fill=green_fill)
    )

def process_margin_file(csv_path, output_dir):
    """
    Reads one CSV, calculates margin (assuming a 40% discount on Price),
    writes to an Excel file sorted by margin ascending, and applies conditional formatting.

    Margin Formula:
      effective_price = Price * 0.6   # 40% discount
      margin = (effective_price - Cost) / effective_price
    """
    df = pd.read_csv(csv_path)

    # Check for required columns: 'Available', 'Cost', 'Price'
    required_cols = ['Available', 'Cost', 'Price']
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        print(f"Skipping {csv_path} - missing columns: {missing_cols}")
        return

    # Filter out rows:
    # 1) Where Cost <= 1
    df = df[df['Cost'] > 1.0]
    # 2) Where Available = 0
    df = df[df['Available'] != 0]
    # 3) Price is null or 0
    df = df.dropna(subset=['Price'])
    df = df[df['Price'] != 0]

    # Calculate effective price after 40% discount
    df['Effective_Price'] = (df['Price'] * 0.7) - ((df['Price'] * 0.7)/10) 

    # Remove any rows where Effective_Price might be 0 or negative (extreme edge cases)
    df = df[df['Effective_Price'] > 0]

    # Calculate margin
    df['Margin'] = (df['Effective_Price'] - df['Cost']) / df['Effective_Price']

    # Sort by margin ascending
    df.sort_values(by='Margin', inplace=True)

    # Build output filename
    base_name = os.path.splitext(os.path.basename(csv_path))[0]
    output_file = os.path.join(output_dir, f"{base_name}_margin_report.xlsx")

    # Write to Excel
    df.to_excel(output_file, index=False, sheet_name='MarginReport')

    # Apply conditional formatting with openpyxl
    wb = load_workbook(output_file)
    ws = wb['MarginReport']

    # Find which column index is 'Margin'
    margin_col_idx = None
    for idx, cell in enumerate(ws[1], start=1):
        if cell.value == 'Margin':
            margin_col_idx = idx
            break

    if margin_col_idx is not None:
        apply_conditional_formatting(wb, 'MarginReport', margin_col_idx)

    wb.save(output_file)
    print(f"Created: {output_file}")

def main():
    # Define input folder (where your CSVs are) and output folder
    input_folder = "files"
    output_folder = "margin_reports"

    ensure_dir_exists(output_folder)

    # Loop through CSV files in the input folder
    for filename in os.listdir(input_folder):
        if filename.lower().endswith(".csv"):
            csv_path = os.path.join(input_folder, filename)
            process_margin_file(csv_path, output_folder)

    print("\nAll margin reports have been generated.")

if __name__ == "__main__":
    main()
