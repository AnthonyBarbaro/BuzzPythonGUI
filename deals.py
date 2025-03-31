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
import shutil

# Global dictionary to map real names -> pseudonyms
NAME_MAP = {}
GLOBAL_COUNTER = 1

def pseudonymize_name(name):
    """
    Replace a customer's real name with a consistent pseudonym.
    Example: "John Smith" -> "Customer_1"
    The same name always maps to the same pseudonym during a single script run.
    """
    global GLOBAL_COUNTER
    
    if pd.isnull(name) or not isinstance(name, str):
        return ""

    name = name.strip()
    if name not in NAME_MAP:
        # Create a new pseudonym
        NAME_MAP[name] = f"Customer_{GLOBAL_COUNTER}"
        GLOBAL_COUNTER += 1

    return NAME_MAP[name]

locale.setlocale(locale.LC_ALL, '')  # Use '' for system's default locale (e.g., USD for the US)

def process_file(file_path):
    """Reads an Excel file with a known structure (header=4),
    standardizes columns, and adds a 'day of week' column."""
    if not os.path.exists(file_path):
        print(f"Error: The file at path {file_path} does not exist.")
        return None

    df = pd.read_excel(file_path, header=4)
    df.columns = df.columns.str.strip().str.lower()

    # Standard set of columns for Dutchie sales exports
    df.columns = [
        "order id", "order time", "budtender name", "customer name", "customer type",
        "vendor name", "product name", "category", "package id", "batch id",
        "external package id", "total inventory sold", "unit weight sold", "total weight sold",
        "gross sales", "inventory cost", "discounted amount", "loyalty as discount",
        "net sales", "return date", "upc gtin (canada)", "provincial sku (canada)",
        "producer", "order profit"
    ]

    # Convert order time to datetime, then create day-of-week
    df['order time'] = pd.to_datetime(df['order time'], errors='coerce')
    df['day of week'] = df['order time'].dt.strftime('%A')
    df['customer name'] = df['customer name'].apply(pseudonymize_name)
    # Debug: show shape and columns
    print(f"DEBUG: Successfully read {file_path}")
    print(f"DEBUG: {file_path} shape: {df.shape}")
    print(f"DEBUG: {file_path} columns: {list(df.columns)}")
    return df

import numpy as np
def apply_discounts_and_kickbacks(data, discount, kickback):
    """
    Adds discount/kickback columns and extra calculated metrics to the DataFrame.
    """

    # 1) Original discount/kickback
    data['discount amount'] = data['gross sales'] * discount
    data['kickback amount'] = data['inventory cost'] * kickback

    # 2) Net Profit = Gross Sales - Inventory Cost - Discount Amount
    data['net profit'] = data['gross sales'] - data['inventory cost'] - data['discount amount']

    # 3) Gross Margin % = ((Gross Sales - Inventory Cost) / Gross Sales) * 100
    #    Avoid division by zero using np.where
    data['gross margin %'] = np.where(
        data['gross sales'] != 0,
        (data['gross sales'] - data['inventory cost']) / data['gross sales'] * 100,
        0
    )

    # 4) Discount % = (Discount Amount / Gross Sales) * 100
    data['discount %'] = np.where(
        data['gross sales'] != 0,
        (data['discount amount'] / data['gross sales']) * 100,
        0
    )

    # 5) Profit Margin % = (Net Profit / Gross Sales) * 100
    data['profit margin %'] = np.where(
        data['gross sales'] != 0,
        (data['net profit'] / data['gross sales']) * 100,
        0
    )

    # 6) Break-Even Sales = Inventory Cost + Discount Amount
    data['break-even sales'] = data['inventory cost'] + data['discount amount']

    # 7) Efficiency Ratio = Gross Sales / Inventory Cost
    data['efficiency ratio'] = np.where(
        data['inventory cost'] != 0,
        data['gross sales'] / data['inventory cost'],
        0
    )

    # 8) Discount Impact % = (Discount Amount / Inventory Cost) * 100
    data['discount impact %'] = np.where(
        data['inventory cost'] != 0,
        (data['discount amount'] / data['inventory cost']) * 100,
        0
    )

    # 9) Sales to Cost Ratio = Gross Sales / Inventory Cost
    data['sales to cost ratio'] = np.where(
        data['inventory cost'] != 0,
        data['gross sales'] / data['inventory cost'],
        0
    )

    return data
#Month to month
brand_criteria3 = { #Started March 1
    'Made': {
        'vendors': ['Garden Of Weeden Inc.'],
        'days': ['Saturday'],
        'discount': 0.50,
        'kickback': 0.25,
        #'categories': ['Cartridges'], 
        'brands': ['Made']
    }
}
brand_criteria2 = {
    'Stiiizy': {
        'vendors': ['Elevation (Stiiizy)'],
        'days': ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
        'discount': 0.50,
        'kickback': 0.30,
        'brands': ['Stiiizy']
    }
}
brand_criteria4 = {
    'Wyld-FEB': {
        'vendors': ['2020 Long Beach LLC'],
        'days': ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
        'discount': 0.30,
        'kickback': 0.0,
        'brands': ['Wyld', 'Good Tide']
    },'GoodTide-FEB': {
        'vendors': ['2020 Long Beach LLC'],
        'days': ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
        'discount': 0.30,
        'kickback': 0.0,
        'brands': ['Good Tide']
    },
}

# Define brand-based criteria
brand_criteria1 = {
    'Monday': {
        'vendors': ['Evergreen SR. Distribution Inc.'],
        'days': ['Monday'],
        'discount': 0.30,
        'kickback': 0.0,
        'brands': ['Smackers']
    },
    'Tuesday': {
        'vendors': ['Evergreen SR. Distribution Inc.'],
        'days': ['Tuesday'],
        'discount': 0.30,
        'kickback': 0.0,
        'brands': ['Smackers']
    },
    'Wednesday': {
        'vendors': ['Evergreen SR. Distribution Inc.'],
        'days': ['Wednesday'],
        'discount': 0.30,
        'kickback': 0.0,
        'brands': ['Smackers']
    },
    'Thursday': {
        'vendors': ['Evergreen SR. Distribution Inc.'],
        'days': ['Thursday'],
        'discount': 0.30,
        'kickback': 0.0,
        'brands': ['Smackers']
    },
    'Friday': {
        'vendors': ['Evergreen SR. Distribution Inc.'],
        'days': ['Friday'],
        'discount': 0.30,
        'kickback': 0.0,
        'brands': ['Smackers']
    },
    'Saturday': {
        'vendors': ['Evergreen SR. Distribution Inc.'],
        'days': ['Saturday'],
        'discount': 0.30,
        'kickback': 0.0,
        'brands': ['Smackers']
    },
    'Sunday': {
        'vendors': ['Evergreen SR. Distribution Inc.'],
        'days': ['Sunday'],
        'discount': 0.30,
        'kickback': 0.0,
        'brands': ['Smackers']
    },
    'Smackers': {
        'vendors': ['Evergreen SR. Distribution Inc.'],
        'days': ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
        'discount': 0.30,
        'kickback': 0.0,
        'brands': ['Smackers']
    }
}
brand_criteria = {
    'Hashish': {
        'vendors': ['Zenleaf LLC'],
        'days': ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
        'discount': 0.50,
        'kickback': 0.25,
        #'categories': ['Concentrate'], 
        'brands': ['Hashish']
    },
    'Jeeter': {
        'vendors': ['Med For America Inc.'],
        'days': ['Monday','Tuesday','Wednesday','Thursday','Friday','Saturday','Sunday'],
        'discount': 0.40,
        'kickback': 0.23,
        'categories': ['Pre-Rolls'],
        'brands': ['Jeeter'],
        'include_phrases': ['LRO','2G','5pk','1G'],
        #'excluded_phrases': ['(3pk)','SVL']
    },
    'Kiva': {
        'vendors': ['KIVA / LCISM CORP', 'Vino & Cigarro, LLC'],
        'days': ['Monday','Wednesday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Terra', 'Petra', 'KIVA', 'Lost Farms', 'Camino']
    },
    'BigPetes': {
        'vendors': ["Big Pete's | LCISM Corp","Vino & Cigarro, LLC"],
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
    'Dabwoods': {
        'vendors': ['The Clear Group Inc.','Decoi','Garden Of Weeden Inc.'],
        'days': ['Friday','Saturday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Dabwoods']
    },
    'Time Machine': {
        'vendors': ['Vino & Cigarro, LLC'],
        'days': ['Tuesday','Friday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Time Machine']
    },
    'Pacific Stone': {
        'vendors': ['Vino & Cigarro, LLC'],
        'days': ['Monday','Thursday'],
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
    'Almora': {
        'vendors': ['Fluids Manufacturing Inc.'],
        'days': ['Sunday','Saturday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Almora']
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
        #'excluded_phrases': ['Jetty | Cart 1g |'],
        'include_phrases': ['SVL'],
        'brands': ['Jetty']
    },
    'Dr.Norm': {
        'vendors': ['Punch Media, LLC'],
        'days': ['Thursday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Dr. Norms']
    },
    'Smokiez': {
        'vendors': ['Garden Of Weeden Inc.'],
        'days': ['Sunday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Smokies']
    },
    'Preferred': {
        'vendors': ['Garden Of Weeden Inc.','Helios | Hypeereon Corporation'],
        'days': ['Monday','Wednesday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Preferred Gardens',]
    },
    'Kikoko': {
        'vendors': ['Garden Of Weeden Inc.'],
        'days': ['Wednesday'],
        'discount': 0.50,
        'kickback': 0.30,
        'brands': ['Kikoko']
    },
    'JoshWax': {
        'vendors': ['Zasp'],
        'days': ['Friday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Josh Wax']
    },
    'TreeSap': {
        'vendors': ['Zenleaf LLC'],
        'days': ['Thursday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['TreeSap']
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
        cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="808080", end_color="808080", fill_type="solid")
        cell.border = thin_border

    # 3) Freeze panes at row 3
    sheet.freeze_panes = "A3"

    # 4) Style data rows (row 3 downward)
    for row_idx in range(3, max_row + 1):
        for col_idx in range(1, max_col + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell.border = thin_border

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
                cell.alignment = Alignment(horizontal="left", vertical="center")

            # Banded row coloring
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
    Similar styling for other sheets like MV_Sales, LM_Sales, SV_Sales, etc.
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
            if val is not None:
                val_length = len(str(val))
                if val_length > max_length:
                    max_length = val_length
        sheet.column_dimensions[column_letter].width = max_length + 2

    # Freeze row 1
    sheet.freeze_panes = "A2"

def style_top_sellers_sheet(sheet):
    """
    Styles a 'Top Sellers' sheet:
      - Bold header
      - Currency formatting for "Gross Sales"
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

            # "Gross Sales" is column 2 in "Top Sellers"
            if col_idx == 2:
                cell.number_format = '"$"#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")

            # Alternating row color
            if row_idx % 2 == 1:
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    # Auto-fit columns
    for col_idx in range(1, max_col + 1):
        column_letter = get_column_letter(col_idx)
        max_length = 0
        for row_idx in range(1, sheet.max_row + 1):
            val = sheet.cell(row=row_idx, column=col_idx).value
            if val is not None:
                val_length = len(str(val))
                if val_length > max_length:
                    max_length = val_length
        sheet.column_dimensions[column_letter].width = max_length + 2

def run_deals_reports():
    """
    Reads salesMV.xlsx, salesLM.xlsx, and salesSV.xlsx (if present).
    Generates brand_reports/<brand>_report_...xlsx:
      - Summary sheet first
      - Optional MV_Sales, LM_Sales, SV_Sales (if data exists)
      - Top Sellers
      - Consolidated summary across all brands if any data is found
    Returns a list of dict: [{"brand":..., "owed":..., "start":..., "end":...}, ...]
    """
    # Archive old reports before generating new ones
    output_dir = 'brand_reports'
    old_dir = os.path.join('old')
    Path(output_dir).mkdir(parents=True, exist_ok=True)  # <-- Ensures folder exists

    for file in os.listdir(output_dir):
        full_path = os.path.join(output_dir, file)
        if file.endswith(".xlsx") and os.path.isfile(full_path):
            dest_path = os.path.join(old_dir, file)
            if os.path.exists(dest_path):
                os.remove(dest_path)  # ✅ Remove existing file first
            shutil.move(full_path, dest_path)  # ✅ Safe move


    # Debug: Attempt to read each file
    mv_data = process_file('files/salesMV.xlsx')
    lm_data = process_file('files/salesLM.xlsx')
    sv_data = process_file('files/salesSV.xlsx')  # NEW - If missing, returns None

    # If a store file is None, we skip that store
    if mv_data is None:
        print("DEBUG: MV data not found or empty. Skipping Mission Valley.")
        mv_data = pd.DataFrame()
    if lm_data is None:
        print("DEBUG: LM data not found or empty. Skipping La Mesa.")
        lm_data = pd.DataFrame()
    if sv_data is None:
        print("DEBUG: SV data not found or empty. Skipping Sorrento Valley.")
        sv_data = pd.DataFrame()

    ALL_DAYS = {"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"}
    consolidated_summary = []
    results_for_app = []

    # For each brand, gather data from whichever stores are not empty
    for brand, criteria in brand_criteria.items():

        # ----- Mission Valley ----- #
        mv_brand_data = pd.DataFrame()
        if not mv_data.empty:
            mv_brand_data = mv_data[
                (mv_data['vendor name'].isin(criteria['vendors'])) &
                (mv_data['day of week'].isin(criteria['days']))
            ].copy()

        # ----- La Mesa ----- #
        lm_brand_data = pd.DataFrame()
        if not lm_data.empty:
            lm_brand_data = lm_data[
                (lm_data['vendor name'].isin(criteria['vendors'])) &
                (lm_data['day of week'].isin(criteria['days']))
            ].copy()

        # ----- Sorrento Valley ----- #
        sv_brand_data = pd.DataFrame()
        if not sv_data.empty:
            sv_brand_data = sv_data[
                (sv_data['vendor name'].isin(criteria['vendors'])) &
                (sv_data['day of week'].isin(criteria['days']))
            ].copy()

        # Debug: shapes before further filtering
        print(f"\nDEBUG: {brand} - Initial shapes => MV: {mv_brand_data.shape}, LM: {lm_brand_data.shape}, SV: {sv_brand_data.shape}")

        # Filter categories
        if 'categories' in criteria:
            if not mv_brand_data.empty:
                mv_brand_data = mv_brand_data[mv_brand_data['category'].isin(criteria['categories'])]
            if not lm_brand_data.empty:
                lm_brand_data = lm_brand_data[lm_brand_data['category'].isin(criteria['categories'])]
            if not sv_brand_data.empty:
                sv_brand_data = sv_brand_data[sv_brand_data['category'].isin(criteria['categories'])]

        # Filter brand names
        if 'brands' in criteria:
            brand_list = criteria['brands']
            if not mv_brand_data.empty:
                mv_brand_data = mv_brand_data[mv_brand_data['product name'].apply(
                    lambda x: any(b in x for b in brand_list if isinstance(x, str))
                )]
            if not lm_brand_data.empty:
                lm_brand_data = lm_brand_data[lm_brand_data['product name'].apply(
                    lambda x: any(b in x for b in brand_list if isinstance(x, str))
                )]
            if not sv_brand_data.empty:
                sv_brand_data = sv_brand_data[sv_brand_data['product name'].apply(
                    lambda x: any(b in x for b in brand_list if isinstance(x, str))
                )]
        # Include phrases filter (if provided, these take priority)
        if 'include_phrases' in criteria:
            include_patterns = [re.escape(p) for p in criteria['include_phrases']]

            def matches_include(x):
                return any(re.search(pat, x, re.IGNORECASE) for pat in include_patterns if isinstance(x, str))

            if not mv_brand_data.empty:
                mv_brand_data = mv_brand_data[mv_brand_data['product name'].apply(matches_include)]
            if not lm_brand_data.empty:
                lm_brand_data = lm_brand_data[lm_brand_data['product name'].apply(matches_include)]
            if not sv_brand_data.empty:
                sv_brand_data = sv_brand_data[sv_brand_data['product name'].apply(matches_include)]

        # Excluded phrases
        if 'excluded_phrases' in criteria:
            for phrase in criteria['excluded_phrases']:
                pat = re.escape(phrase)

                if not mv_brand_data.empty:
                    mv_brand_data = mv_brand_data[~mv_brand_data['product name'].str.contains(pat, na=False)]
                if not lm_brand_data.empty:
                    lm_brand_data = lm_brand_data[~lm_brand_data['product name'].str.contains(pat, na=False)]
                if not sv_brand_data.empty:
                    sv_brand_data = sv_brand_data[~sv_brand_data['product name'].str.contains(pat, na=False)]

        # Debug: shapes after filtering
        print(f"DEBUG: {brand} - After filtering => MV: {mv_brand_data.shape}, LM: {lm_brand_data.shape}, SV: {sv_brand_data.shape}")

        # Skip brand if all stores are empty
        if mv_brand_data.empty and lm_brand_data.empty and sv_brand_data.empty:
            print(f"DEBUG: No data remains for brand '{brand}'. Skipping.")
            continue

        # Apply discount/kickback if columns exist
        required_cols = {'gross sales','inventory cost'}
        if not mv_brand_data.empty and required_cols.issubset(mv_brand_data.columns):
            mv_brand_data = apply_discounts_and_kickbacks(mv_brand_data, criteria['discount'], criteria['kickback'])
        if not lm_brand_data.empty and required_cols.issubset(lm_brand_data.columns):
            lm_brand_data = apply_discounts_and_kickbacks(lm_brand_data, criteria['discount'], criteria['kickback'])
        if not sv_brand_data.empty and required_cols.issubset(sv_brand_data.columns):
            sv_brand_data = apply_discounts_and_kickbacks(sv_brand_data, criteria['discount'], criteria['kickback'])

        # Determine date ranges
        def get_date_range(df):
            if df.empty:
                return None, None
            return df['order time'].min(), df['order time'].max()

        start_mv, end_mv = get_date_range(mv_brand_data)
        start_lm, end_lm = get_date_range(lm_brand_data)
        start_sv, end_sv = get_date_range(sv_brand_data)

        possible_starts = [d for d in [start_mv, start_lm, start_sv] if d]
        possible_ends = [d for d in [end_mv, end_lm, end_sv] if d]
        if not possible_starts or not possible_ends:
            print(f"DEBUG: Brand '{brand}' had data, but no valid date range. Skipping.")
            continue

        # Convert to strings
        start_date = min(possible_starts).strftime('%Y-%m-%d')
        end_date = max(possible_ends).strftime('%Y-%m-%d')
        date_range = f"{start_date}_to_{end_date}"

        # Summaries for each store
        # Create an aggregator function only if the DataFrame is not empty
        def build_summary(df, store_name):
            if df.empty:
                # Return an empty summary with the same columns
                return pd.DataFrame(columns=['gross sales','inventory cost','discount amount','kickback amount','location'])
            summary = df.agg({
                'gross sales': 'sum',
                'inventory cost': 'sum',
                'discount amount': 'sum',
                'kickback amount': 'sum'
            }).to_frame().T
            summary['location'] = store_name
            return summary

        mv_summary = build_summary(mv_brand_data, 'Mission Valley')
        lm_summary = build_summary(lm_brand_data, 'La Mesa')
        sv_summary = build_summary(sv_brand_data, 'Sorrento Valley')

        # Combine them
        brand_summary = pd.concat([mv_summary, lm_summary, sv_summary], ignore_index=True)

        # If the brand runs all days, show 'Everyday'
        if set(criteria['days']) == ALL_DAYS:
            days_text = "Everyday"
        else:
            days_text = ", ".join(criteria['days'])

        # Rename columns and add brand
        brand_summary.rename(columns={
            'location': 'Store',
            'kickback amount': 'Kickback Owed'
        }, inplace=True)

        brand_summary['Days Active'] = days_text
        brand_summary['Date Range'] = f"{start_date} to {end_date}"
        brand_summary['Brand'] = brand
        brand_summary['Margin'] = None
        # Reorder columns
        col_order = [
            'Store', 'Kickback Owed', 'Days Active', 'Date Range',
            'gross sales', 'inventory cost', 'discount amount','Margin','Brand'
        ]
        final_cols = [c for c in col_order if c in brand_summary.columns]
        brand_summary = brand_summary[final_cols]

        # Save to consolidated summary
        consolidated_summary.append(brand_summary)

        # --- CREATE BRAND-LEVEL EXCEL ---
        safe_brand_name = brand.replace("/", " ")
        output_filename = os.path.join(output_dir, f"{safe_brand_name}_report_{date_range}.xlsx")
        print(f"DEBUG: Creating {output_filename} for brand '{brand}'...")

        # Build "Top Sellers" with combined data from MV + LM + SV
        combined_df = pd.concat([mv_brand_data, lm_brand_data, sv_brand_data], ignore_index=True)
        if not combined_df.empty and 'gross sales' in combined_df.columns:
            top_sellers_df = (
                combined_df.groupby('product name', as_index=False)
                .agg({'gross sales': 'sum'})
                .sort_values(by='gross sales', ascending=False)
                .head(20)
            )
            top_sellers_df.rename(columns={'product name': 'Product Name', 'gross sales': 'Gross Sales'}, inplace=True)
        else:
            # No data => create empty
            top_sellers_df = pd.DataFrame(columns=['Product Name','Gross Sales'])

        with pd.ExcelWriter(output_filename) as writer:
            # Summary
            brand_summary.to_excel(writer, sheet_name='Summary', index=False, startrow=1)
            
            # MV Sales (if not empty)
            if not mv_brand_data.empty:
                mv_brand_data.to_excel(writer, sheet_name='MV_Sales', index=False)
            # LM Sales (if not empty)
            if not lm_brand_data.empty:
                lm_brand_data.to_excel(writer, sheet_name='LM_Sales', index=False)
            # SV Sales (if not empty)
            if not sv_brand_data.empty:
                sv_brand_data.to_excel(writer, sheet_name='SV_Sales', index=False)
            # Top Sellers
            top_sellers_df.to_excel(writer, sheet_name='Top Sellers', index=False)
        wb = load_workbook(output_filename)
        summary_sheet = wb['Summary']

        # The header is in row 2, data starts in row 3
        data_start_row = 3
        data_end_row = data_start_row + brand_summary.shape[0] - 1

        # We know from col_order that:
        #   E => "gross sales"  => column 5
        #   F => "inventory cost" => column 6
        #   G => "discount amount" => column 7
        #   H => "Margin" => column 8
        # So for each row of data, we set the formula in col 8
        for row_idx in range(data_start_row, data_end_row + 1):
            # Example formula: =((E3-G3)-(F3*0.75))/(E3-G3)
            summary_sheet.cell(row=row_idx, column=8).value = (
                f"=((E{row_idx}-G{row_idx})-(F{row_idx}-B{row_idx}))/(E{row_idx}-G{row_idx})"
            )
        wb.save(output_filename)
        # Style
        wb = load_workbook(output_filename)
        if 'Summary' in wb.sheetnames:
            style_summary_sheet(wb['Summary'], brand)
        if 'MV_Sales' in wb.sheetnames:
            style_worksheet(wb['MV_Sales'])
        if 'LM_Sales' in wb.sheetnames:
            style_worksheet(wb['LM_Sales'])
        if 'SV_Sales' in wb.sheetnames:
            style_worksheet(wb['SV_Sales'])
        if 'Top Sellers' in wb.sheetnames:
            style_top_sellers_sheet(wb['Top Sellers'])
        wb.save(output_filename)

        # Kickback owed = sum of 'Kickback Owed' from all stores
        total_owed = brand_summary['Kickback Owed'].sum()
        results_for_app.append({
            "brand": brand,
            "owed": float(total_owed),
            "start": start_date,
            "end": end_date
        })

    # Build a consolidated summary (all brands) if we have data
    if consolidated_summary:
        final_df = pd.concat(consolidated_summary, ignore_index=True)
        # For last brand, we used date_range, but let's re-derive overall range
        # in case different brands have different date ranges
        # We can keep it simple and just use the last brand's date_range or pick min->max from results_for_app
        if results_for_app:
            all_starts = [r['start'] for r in results_for_app if r['start']]
            all_ends = [r['end'] for r in results_for_app if r['end']]
            if all_starts and all_ends:
                overall_start = min(all_starts)
                overall_end = max(all_ends)
                overall_range = f"{overall_start}_to_{overall_end}"
            else:
                # fallback
                overall_range = date_range
        else:
            overall_range = date_range

        consolidated_file = os.path.join(output_dir, f"consolidated_brand_report_{overall_range}.xlsx")
        print(f"DEBUG: Creating consolidated summary => {consolidated_file}")
        with pd.ExcelWriter(consolidated_file) as writer:
            final_df.to_excel(writer, sheet_name='Consolidated_Summary', index=False, startrow=1)


        # Style consolidated summary
        wb = load_workbook(consolidated_file)
        if 'Consolidated_Summary' in wb.sheetnames:
            sheet = wb['Consolidated_Summary']
            # The header row is row=2, so data starts at row=3
            data_start_row = 3
            data_end_row = data_start_row + final_df.shape[0] - 1

            # We know from your column order:
            # 1) Store
            # 2) Kickback Owed
            # 3) Days Active
            # 4) Date Range
            # 5) gross sales      (column E)
            # 6) inventory cost   (column F)
            # 7) discount amount  (column G)
            # 8) Margin           (column H)
            # 9) Brand
            #
            # Let's inject your formula in column H for each row:
            for row_idx in range(data_start_row, data_end_row + 1):
                sheet.cell(row=row_idx, column=8).value = (
                    f"=((E{row_idx}-G{row_idx})-(F{row_idx}-B{row_idx}))/(E{row_idx}-G{row_idx})"
                )
            # We'll just style it with "ALL_BRANDS" 
            style_summary_sheet(sheet, "ALL_BRANDS")
        wb.save(consolidated_file)

        print("Individual brand reports + consolidated report have been saved.")
    else:
        print("No brand data found; no Excel files generated.")

    return results_for_app

if __name__ == "__main__":
    data = run_deals_reports()
    print("Results for app:", data)
