#!/usr/bin/env python3
import os
import re
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
import locale
locale.setlocale(locale.LC_ALL, '')  # Use '' for system's default locale (e.g., USD for the US)

def process_file(file_path):
    if not os.path.exists(file_path):
        print(f"Error: The file at path {file_path} does not exist.")
        return None

    df = pd.read_excel(file_path, header=4)
    df.columns = df.columns.str.strip().str.lower()
    df.columns = [
        "order id", "order time", "budtender name", "customer name", "customer type",
        "vendor name", "product name", "category", "package id", "batch id",
        "external package id", "total inventory sold", "unit weight sold", "total weight sold",
        "gross sales", "inventory cost", "discounted amount", "loyalty as discount",
        "net sales", "return date", "upc gtin (canada)", "provincial sku (canada)",
        "producer", "order profit"
    ]

    df['order time'] = pd.to_datetime(df['order time'], errors='coerce')
    df['day of week'] = df['order time'].dt.strftime('%A')
    return df

def apply_discounts_and_kickbacks(data, discount, kickback):
    """Adds discount/kickback columns to the DataFrame."""
    data['discount amount'] = data['gross sales'] * discount
    data['kickback amount'] = data['inventory cost'] * kickback
    return data

brand_criteria = {
    'Hashish': {
        'vendors': ['BTC Ventures', 'Zenleaf LLC', 'Garden Of Weeden Inc.'],
        'days': ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
        'discount': 0.50,
        'kickback': 0.25,
        'categories': ['Concentrate'],
        'brands': ['Hashish']
    },
    'Jeeter': {
        'vendors': ['Med For America Inc.'],
        'days': ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
        'discount': 0.50,
        'kickback': 0.25,
        'categories': ['Pre-Rolls'],
        'brands': ['Jeeter'],
        'excluded_phrases': ['(3pk)','Jeeter | SVL']
    },
    'Kiva': {
        'vendors': ['KIVA / LCISM CORP', 'Vino & Cigarro, LLC'],
        'days': ['Monday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Terra', 'Petra', 'Kiva', 'Lost Farms', 'Camino']
    },
    'BigPetes': {
        'vendors': ['KIVA / LCISM CORP', 'Vino & Cigarro, LLC'],
        'days': ['Tuesday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Big Pete']
    },
    'HolySmoke/Water': {
        'vendors': ['Heritage Holding of Califonia, Inc.', 'Barlow Printing LLC'],
        'days': ['Sunday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Holy Smokes', 'Holy Water']
    },
    'Dawoods': {
        'vendors': ['The Clear Group Inc.'],
        'days': ['Friday','Saturday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Dabwoods']
    },
    'Time Machine': {
        'vendors': ['Vino & Cigarro, LLC'],
        'days': ['Tuesday','Thursday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Time Machine']
    },
    'Pacific Stone': {
        'vendors': ['Vino & Cigarro, LLC'],
        'days': ['Friday','Monday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Pacific Stone']
    },
    'Heavy Hitters': {
        'vendors': ['Fluids Manufacturing Inc.'],
        'days': ['Friday','Saturday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Heavy Hitters']
    },
    'WYLD/GoodTide': {
        'vendors': ['2020 Long Beach LLC'],
        'days': ['Friday','Saturday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Wyld', 'Good Tide']
    },
    'Jetty': {
        'vendors': ['KIVA / LCISM CORP', 'Vino & Cigarro, LLC','Garden Of Weeden Inc.'],
        'days': ['Thursday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Jetty']
    }
}
def style_summary_sheet(sheet, brand_name):
    """
    Styles the Summary sheet:
      - A bold title in row 1
      - Headers in row 2 (gray background, centered)
      - Data starts in row 3
      - Freeze pane at A3
      - Banded row styling for data
      - Currency/date formatting as needed
    """
    max_col = sheet.max_column
    max_row = sheet.max_row

    # 1) Big title in row 1
    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    title_cell = sheet.cell(row=1, column=1)
    title_cell.value = f"{brand_name.upper()} SUMMARY REPORT"
    title_cell.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    # 2) Style header row (row 2)
    header_row_idx = 2
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )

    for col_idx in range(1, max_col + 1):
        cell = sheet.cell(row=header_row_idx, column=col_idx)
        # Make it bold, white text, center aligned, gray fill
        cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
        cell.border = thin_border

    # 3) Freeze panes at row 3 (so row 1 & 2 stay visible)
    sheet.freeze_panes = "A3"

    # 4) Style data rows (row 3 downward)
    for row_idx in range(3, max_row + 1):  # Adjusted to start styling from row 3
        for col_idx in range(1, max_col + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell.border = thin_border

            # Check the header text
            hdr_val = sheet.cell(row=header_row_idx, column=col_idx).value
            if hdr_val and ("owed" in str(hdr_val).lower()):
                # Format as currency
                cell.number_format = '"$"#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif hdr_val and ("date" in str(hdr_val).lower()):
                # Format as date
                cell.number_format = "YYYY-MM-DD"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                # Default alignment is left
                cell.alignment = Alignment(horizontal="left", vertical="center")

            # Banded row coloring for readability
            if row_idx % 2 == 1:  # Odd data row
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    # 5) Auto-fit column widths
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

def style_worksheet(sheet):
    """
    Similar styling for other sheets like MV_Sales, LM_Sales, etc.
    """

    max_col = sheet.max_column
    # Make header row bold and center-aligned
    for col_idx in range(1, max_col + 1):
        cell = sheet.cell(row=1, column=col_idx)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Auto-fit column width
    for col_idx in range(1, max_col + 1):
        column_letter = get_column_letter(col_idx)
        max_length = 0
        for row_idx in range(1, sheet.max_row + 1):
            val = sheet.cell(row=row_idx, column=col_idx).value
            try:
                max_length = max(max_length, len(str(val)) if val else 0)
            except:
                pass
        sheet.column_dimensions[column_letter].width = max_length + 2
        sheet.freeze_panes = "A2"
def style_top_sellers_sheet(sheet):
    """
    Styles a 'Top Sellers' sheet:
      - Bold header
      - Currency formatting for Gross Sales
      - Alternating row colors
      - Auto-fit columns
    """
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    max_col = sheet.max_column

    # Header row
    for col_idx in range(1, max_col + 1):
        cell = sheet.cell(row=1, column=col_idx)
        cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    # Data rows
    for row_idx in range(2, sheet.max_row + 1):
        for col_idx in range(1, max_col + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            if col_idx == 2:  # "Gross Sales" column
                cell.number_format = '"$"#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")
            if row_idx % 2 == 1:
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    # Auto-fit columns
    for col_idx in range(1, max_col + 1):
        column_letter = get_column_letter(col_idx)
        max_length = 0
        for row_idx in range(1, sheet.max_row + 1):
            val = sheet.cell(row=row_idx, column=col_idx).value
            try:
                max_length = max(max_length, len(str(val)) if val else 0)
            except:
                pass
        sheet.column_dimensions[column_letter].width = max_length + 2

def run_deals_reports():
    """
    1) Reads salesMV.xlsx, salesLM.xlsx
    2) Generates brand_reports/<brand>_report_...xlsx
       - Summary sheet first
       - Reorders columns so store is first, then Kickback Owed, Days Active, Date Range, etc.
       - If brand runs all 7 days, show 'Everyday' instead of listing them all.
    3) Returns list of dict: [{"brand":..., "owed":..., "start":..., "end":...}, ...]
    """
    output_dir = 'brand_reports'
    Path(output_dir).mkdir(exist_ok=True)

    mv_data = process_file('files/salesMV.xlsx')
    lm_data = process_file('files/salesLM.xlsx')
    if mv_data is None or lm_data is None:
        print("One or both sales files missing; no data returned.")
        return []

    # We'll define all days for easy check
    ALL_DAYS = {"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"}

    consolidated_summary = []
    results_for_app = []

    for brand, criteria in brand_criteria.items():
        # Gather brand data for MV
        mv_brand_data = mv_data[
            (mv_data['vendor name'].isin(criteria['vendors'])) &
            (mv_data['day of week'].isin(criteria['days']))
        ].copy()

        # Gather brand data for LM
        lm_brand_data = lm_data[
            (lm_data['vendor name'].isin(criteria['vendors'])) &
            (lm_data['day of week'].isin(criteria['days']))
        ].copy()

        # Filter categories, if any
        if 'categories' in criteria:
            mv_brand_data = mv_brand_data[mv_brand_data['category'].isin(criteria['categories'])]
            lm_brand_data = lm_brand_data[lm_brand_data['category'].isin(criteria['categories'])]

        # Filter brand names, if any
        if 'brands' in criteria:
            mv_brand_data = mv_brand_data[mv_brand_data['product name'].apply(
                lambda x: any(b in x for b in criteria['brands'] if isinstance(x, str))
            )]
            lm_brand_data = lm_brand_data[lm_brand_data['product name'].apply(
                lambda x: any(b in x for b in criteria['brands'] if isinstance(x, str))
            )]

        # Exclude certain phrases
        if 'excluded_phrases' in criteria:
            for phrase in criteria['excluded_phrases']:
                pat = re.escape(phrase)
                mv_brand_data = mv_brand_data[~mv_brand_data['product name'].str.contains(pat, na=False)]
                lm_brand_data = lm_brand_data[~lm_brand_data['product name'].str.contains(pat, na=False)]

        # Skip if both are empty
        if mv_brand_data.empty and lm_brand_data.empty:
            continue

        # Apply discount & kickback
        mv_brand_data = apply_discounts_and_kickbacks(mv_brand_data, criteria['discount'], criteria['kickback'])
        lm_brand_data = apply_discounts_and_kickbacks(lm_brand_data, criteria['discount'], criteria['kickback'])

        # Figure out date range
        if not mv_brand_data.empty:
            start_mv = mv_brand_data['order time'].min().strftime('%Y-%m-%d')
            end_mv = mv_brand_data['order time'].max().strftime('%Y-%m-%d')
        else:
            start_mv = end_mv = None

        if not lm_brand_data.empty:
            start_lm = lm_brand_data['order time'].min().strftime('%Y-%m-%d')
            end_lm = lm_brand_data['order time'].max().strftime('%Y-%m-%d')
        else:
            start_lm = end_lm = None

        possible_starts = [s for s in [start_mv, start_lm] if s]
        possible_ends = [e for e in [end_mv, end_lm] if e]
        if not possible_starts or not possible_ends:
            continue

        start_date = min(possible_starts)
        end_date = max(possible_ends)
        date_range = f"{start_date}_to_{end_date}"

        # Summaries
        mv_summary = mv_brand_data.agg({
            'gross sales': 'sum',
            'inventory cost': 'sum',
            'discount amount': 'sum',
            'kickback amount': 'sum'
        }).to_frame().T
        mv_summary['location'] = 'Misson Valley'

        lm_summary = lm_brand_data.agg({
            'gross sales': 'sum',
            'inventory cost': 'sum',
            'discount amount': 'sum',
            'kickback amount': 'sum'
        }).to_frame().T
        lm_summary['location'] = 'La Mesa'

        # Combine them
        brand_summary = pd.concat([mv_summary, lm_summary], ignore_index=True)

        # If the brand runs all days, show 'Everyday', else show them
        if set(criteria['days']) == ALL_DAYS:
            days_text = "Everyday"
        else:
            days_text = ", ".join(criteria['days'])

        # We want columns: 
        # 1) Store (renamed from location)
        # 2) Kickback Owed (from kickback amount)
        # 3) Days Active
        # 4) Date Range
        # [We can also keep 'Brand' if you'd like, or remove other columns.]

        # Let's rename existing columns to make the final DataFrame neat:
        brand_summary.rename(columns={
            'location': 'Store',
            'kickback amount': 'Kickback Owed'
        }, inplace=True)
    
        # Add the new columns
        brand_summary['Days Active'] = days_text
        brand_summary['Date Range'] = f"{start_date} to {end_date}"

        # Reorder columns as requested: Store, Kickback Owed, Days Active, Date Range
        # Optionally, we can also keep 'gross sales', 'inventory cost', etc. if needed.
        # We'll keep them for reference but show the important ones first.
        col_order = ['Store', 'Kickback Owed', 'Days Active', 'Date Range',
                     'gross sales', 'inventory cost', 'discount amount', 'Brand']
        
        # brand_summary may not have 'Brand' yet, so let's add it

        brand_summary['Brand'] = brand  # to show the brand name

# Ensure the column order includes only columns that exist in the DataFrame
        final_cols = [c for c in col_order if c in brand_summary.columns]
        brand_summary = brand_summary[final_cols]

        # We'll also store brand_summary into consolidated_summary for the final report
        consolidated_summary.append(brand_summary)

        # Generate brand-level file
        safe_brand_name = brand.replace("/", " ")
        output_filename = os.path.join(output_dir, f"{safe_brand_name}_report_{date_range}.xlsx")

        # Create top sellers if desired
        combined_df = pd.concat([mv_brand_data, lm_brand_data], ignore_index=True)
        top_sellers_df = (combined_df.groupby('product name', as_index=False)
                          .agg({'gross sales': 'sum'})
                          .sort_values(by='gross sales', ascending=False)
                          .head(10))
        top_sellers_df.rename(columns={'product name': 'Product Name', 'gross sales': 'Gross Sales'}, inplace=True)

        with pd.ExcelWriter(output_filename) as writer:
            # Summary first
            brand_summary.to_excel(writer, sheet_name='Summary', index=False, startrow=1)

            # MV Sales
            mv_brand_data.to_excel(writer, sheet_name='MV_Sales', index=False)
            # LM Sales
            lm_brand_data.to_excel(writer, sheet_name='LM_Sales', index=False)
            # Top Sellers
            top_sellers_df.to_excel(writer, sheet_name='Top Sellers', index=False)

        # Apply styling
        wb = load_workbook(output_filename)
        if 'Summary' in wb.sheetnames:
            style_summary_sheet(wb['Summary'], brand)

        if 'MV_Sales' in wb.sheetnames:
            style_worksheet(wb['MV_Sales'])

        if 'LM_Sales' in wb.sheetnames:
            style_worksheet(wb['LM_Sales'])

        if 'Top Sellers' in wb.sheetnames:
            style_top_sellers_sheet(wb['Top Sellers'])

        wb.save(output_filename)

        # total owed = sum of all 'Kickback Owed' for that brand (MV + LM)
        total_owed = brand_summary['Kickback Owed'].sum()
        results_for_app.append({
            "brand": brand,
            "owed": float(total_owed),
            "start": start_date,
            "end": end_date
        })

    # Finally, build the consolidated report if we have data
    if consolidated_summary:
        final_df = pd.concat(consolidated_summary, ignore_index=True)
        overall_range = f"{start_date}_to_{end_date}"
        consolidated_file = os.path.join(output_dir, f"consolidated_brand_report_{overall_range}.xlsx")
        with pd.ExcelWriter(consolidated_file) as writer:
            final_df.to_excel(writer, sheet_name='Consolidated_Summary', index=False)

        # Style the consolidated summary
        wb = load_workbook(consolidated_file)
        if 'Consolidated_Summary' in wb.sheetnames:
            sheet = wb['Consolidated_Summary']
            style_summary_sheet(sheet, safe_brand_name)
        wb.save(consolidated_file)

        print("Individual brand reports and a consolidated report have been saved.")
    else:
        print("No brand data found; no Excel files generated.")

    return results_for_app

if __name__ == "__main__":
    data = run_deals_reports()
    print("Results for app:", data)
