#!/usr/bin/env python3
import os
import re
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
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
        'kickback': 0.20,
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

def run_deals_for_store(store):
    """
    Reads data for the specified store (MV or LM).
    Calculates inventory cost and applies percentage-based calculations per brand.
    Returns a list of dicts: [{'store': ..., 'brand': ..., 'inventory_cost': ..., 'kickback': ..., 'start': ..., 'end': ...}, ...]
    """
    if store == 'MV':
        file_path = 'files/salesMV1.xlsx'
    elif store == 'LM':
        file_path = 'files/salesLM1.xlsx'
    else:
        raise ValueError(f"Invalid store: {store}")

    df = process_file(file_path)
    if df is None:
        print(f"No data found for store {store} at {file_path}.")
        return []

    output_dir = f'brand_reports_{store}'
    os.makedirs(output_dir, exist_ok=True)  # Ensure base directory exists

    results = []

    for brand, criteria in brand_criteria.items():
        # Filter data by vendor and days
        store_data = df[
            (df['vendor name'].isin(criteria['vendors'])) &
            (df['day of week'].isin(criteria['days']))
        ].copy()

        # Additional filters
        if 'categories' in criteria:
            store_data = store_data[store_data['category'].isin(criteria['categories'])]
        if 'brands' in criteria:
            store_data = store_data[
                store_data['product name'].apply(
                    lambda x: any(b in x for b in criteria['brands'] if isinstance(x, str))
                )
            ]

        if store_data.empty:
            continue

        # Apply calculations for percentages (kickback, discount)
        store_data = apply_discounts_and_kickbacks(store_data, criteria['discount'], criteria['kickback'])

        # Get the start and end dates
        start_date = store_data['order time'].min().strftime('%Y-%m-%d')
        end_date = store_data['order time'].max().strftime('%Y-%m-%d')

        # Calculate summary statistics
        inventory_cost = store_data['inventory cost'].sum()
        total_kickback = inventory_cost * criteria['kickback']

        # Debugging outputs
        print(f"DEBUG: Store='{store}', Brand='{brand}', Inventory Cost={inventory_cost}, Kickback={total_kickback}")

        # Ensure the subdirectory for this brand exists
        safe_brand_name = brand.replace("/", "_")  # Replace invalid characters
        brand_output_dir = os.path.join(output_dir, safe_brand_name)
        os.makedirs(brand_output_dir, exist_ok=True)  # Create the brand directory if it doesn't exist

        # Save brand-specific Excel report
        output_file = os.path.join(brand_output_dir, f"{safe_brand_name}_{store}_{start_date}_to_{end_date}.xlsx")
        with pd.ExcelWriter(output_file) as writer:
            store_data.to_excel(writer, sheet_name=f"{store}_Sales", index=False)
            summary = pd.DataFrame({
                'Brand': [brand],
                'Store': [store],
                'Inventory Cost': [inventory_cost],
                'Kickback': [total_kickback]
            })
            summary.to_excel(writer, sheet_name="Summary", index=False)

        # Format the Excel file
        wb = load_workbook(output_file)
        for sheet_name in wb.sheetnames:
            sheet = wb[sheet_name]
            for column in sheet.columns:
                max_length = max((len(str(cell.value)) for cell in column if cell.value is not None), default=10)
                sheet.column_dimensions[column[0].column_letter].width = max_length + 2
            for cell in sheet[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal="center")
        wb.save(output_file)
        total_kickback = inventory_cost * criteria['kickback']
        formatted_kickback = locale.currency(total_kickback, grouping=True)
        # Append results for this brand
        results.append({
            "store": store,
            "brand": brand,
            "inventory_cost": inventory_cost,
            "kickback": formatted_kickback,
            "start": start_date,
            "end": end_date
        })

    return results


def run_deals_reports():
    """
    1) Reads salesMV.xlsx, salesLM.xlsx
    2) Generates brand_reports/<brand>_report_...xlsx
    3) Returns list of dict: [{"brand":..., "owed":..., "start":..., "end":...}, ...]
    """
    output_dir = 'brand_reports'
    Path(output_dir).mkdir(exist_ok=True)

    mv_data = process_file('files/salesMV.xlsx')
    lm_data = process_file('files/salesLM.xlsx')
    if mv_data is None or lm_data is None:
        print("One or both sales files missing; no data returned.")
        return []

    consolidated_summary = []
    results_for_app = []

    for brand, criteria in brand_criteria.items():
        mv_brand_data = mv_data[
            (mv_data['vendor name'].isin(criteria['vendors'])) &
            (mv_data['day of week'].isin(criteria['days']))
        ].copy()
        lm_brand_data = lm_data[
            (lm_data['vendor name'].isin(criteria['vendors'])) &
            (lm_data['day of week'].isin(criteria['days']))
        ].copy()

        if 'categories' in criteria:
            mv_brand_data = mv_brand_data[mv_brand_data['category'].isin(criteria['categories'])]
            lm_brand_data = lm_brand_data[lm_brand_data['category'].isin(criteria['categories'])]

        if 'brands' in criteria:
            mv_brand_data = mv_brand_data[mv_brand_data['product name'].apply(
                lambda x: any(b in x for b in criteria['brands'] if isinstance(x, str))
            )]
            lm_brand_data = lm_brand_data[lm_brand_data['product name'].apply(
                lambda x: any(b in x for b in criteria['brands'] if isinstance(x, str))
            )]

        if 'excluded_phrases' in criteria:
            for phrase in criteria['excluded_phrases']:
                pat = re.escape(phrase)
                mv_brand_data = mv_brand_data[~mv_brand_data['product name'].str.contains(pat, na=False)]
                lm_brand_data = lm_brand_data[~lm_brand_data['product name'].str.contains(pat, na=False)]

        if mv_brand_data.empty and lm_brand_data.empty:
            continue

        mv_brand_data = apply_discounts_and_kickbacks(mv_brand_data, criteria['discount'], criteria['kickback'])
        lm_brand_data = apply_discounts_and_kickbacks(lm_brand_data, criteria['discount'], criteria['kickback'])

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

        mv_summary = mv_brand_data.agg({
            'gross sales': 'sum',
            'inventory cost': 'sum',
            'discount amount': 'sum',
            'kickback amount': 'sum'
        }).to_frame().T
        mv_summary['location'] = 'MV'

        lm_summary = lm_brand_data.agg({
            'gross sales': 'sum',
            'inventory cost': 'sum',
            'discount amount': 'sum',
            'kickback amount': 'sum'
        }).to_frame().T
        lm_summary['location'] = 'LM'

        brand_summary = pd.concat([mv_summary, lm_summary])
        brand_summary['brand'] = brand
        brand_summary['days active'] = ', '.join(criteria['days'])
        consolidated_summary.append(brand_summary)

        # Save brand-level Excel
        output_filename = os.path.join(output_dir, f"{brand.replace('/', ' ')}_report_{date_range}.xlsx")
        with pd.ExcelWriter(output_filename) as writer:
            mv_brand_data.to_excel(writer, sheet_name='MV_Sales', index=False)
            lm_brand_data.to_excel(writer, sheet_name='LM_Sales', index=False)
            brand_summary.to_excel(writer, sheet_name='Summary', index=False)

        wb = load_workbook(output_filename)
        for sname in wb.sheetnames:
            sheet = wb[sname]
            for col in sheet.columns:
                mlen = max((len(str(cell.value)) for cell in col if cell.value is not None), default=10)
                sheet.column_dimensions[col[0].column_letter].width = mlen + 2
            for cell in sheet[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
        wb.save(output_filename)

        # total owed is sum of 'kickback amount'
        total_owed = brand_summary['kickback amount'].sum()
        results_for_app.append({
            "brand": brand,
            "owed": float(total_owed),
            "start": start_date,
            "end": end_date
        })

    if consolidated_summary:
        final_df = pd.concat(consolidated_summary, ignore_index=True)
        overall_range = f"{start_date}_to_{end_date}"
        consolidated_file = os.path.join(output_dir, f"consolidated_brand_report_{overall_range}.xlsx")
        with pd.ExcelWriter(consolidated_file) as writer:
            final_df.to_excel(writer, sheet_name='Consolidated_Summary', index=False)

        wb = load_workbook(consolidated_file)
        sh = wb['Consolidated_Summary']
        for col in sh.columns:
            mlen = max((len(str(cell.value)) for cell in col if cell.value is not None), default=10)
            sh.column_dimensions[col[0].column_letter].width = mlen + 2
        for cell in sh[1]:
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='center')
        wb.save(consolidated_file)

        print("Individual brand reports and a consolidated report have been saved.")
    else:
        print("No brand data found; no Excel files generated.")

    return results_for_app

if __name__ == "__main__":
    data = run_deals_reports()
    print("Results for app:", data)

