#!/usr/bin/env python3
import os
import re
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment
from pathlib import Path

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
        'brands': ['Jeeter']
    },
    'Kiva': {
        'vendors': ['KIVA / LCISM CORP', 'Vino & Cigarro, LLC'],
        'days': ['Monday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Terra', 'Petra', 'Kiva', 'Lost Farms', 'Camino']
    },
    'Big Petes': {
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
        'vendors': ['KIVA / LCISM CORP', 'Vino & Cigarro, LLC'],
        'days': ['Thursday'],
        'discount': 0.50,
        'kickback': 0.25,
        'brands': ['Jetty']
    }
}

def run_deals_for_store(store='MV'):
    """
    1) If store='MV', read files/salesMV.xlsx
       If store='LM', read files/salesLM.xlsx
    2) Create brand reports in brand_reports_<store>/ folder
    3) Return a list of dict: [
        {
          "store": <'MV' or 'LM'>,
          "brand": <str>,
          "owed": <float>,
          "start": <YYYY-MM-DD>,
          "end": <YYYY-MM-DD>
        },
        ...
    ]
    """
    # Decide which file
    if store == 'MV':
        file_path = 'files/salesMV.xlsx'
    else:
        file_path = 'files/salesLM.xlsx'

    df = process_file(file_path)
    if df is None:
        print(f"No data found for {store} at {file_path}.")
        return []

    # Output folder
    output_dir = f'brand_reports_{store}'
    Path(output_dir).mkdir(exist_ok=True)

    results_for_app = []

    for brand, criteria in brand_criteria.items():
        # Filter by vendor/days
        store_data = df[
            (df['vendor name'].isin(criteria['vendors'])) &
            (df['day of week'].isin(criteria['days']))
        ].copy()

        if 'categories' in criteria:
            store_data = store_data[store_data['category'].isin(criteria['categories'])]

        if 'brands' in criteria:
            store_data = store_data[
                store_data['product name'].apply(
                    lambda x: any(b in x for b in criteria['brands'] if isinstance(x, str))
                )
            ]

        if 'excluded_phrases' in criteria:
            for phrase in criteria['excluded_phrases']:
                pat = re.escape(phrase)
                store_data = store_data[
                    ~store_data['product name'].str.contains(pat, na=False)
                ]

        if store_data.empty:
            continue

        # Apply discount/kickback
        store_data = apply_discounts_and_kickbacks(
            store_data,
            criteria['discount'],
            criteria['kickback']
        )

        start_date = store_data['order time'].min().strftime('%Y-%m-%d')
        end_date   = store_data['order time'].max().strftime('%Y-%m-%d')
        date_range = f"{start_date}_to_{end_date}"

        # Summaries
        summary_df = store_data.agg({
            'gross sales': 'sum',
            'inventory cost': 'sum',
            'discount amount': 'sum',
            'kickback amount': 'sum'
        }).to_frame().T
        summary_df['store'] = store
        summary_df['brand'] = brand

        # Create Excel
        output_file = os.path.join(output_dir, f"{brand}_{store}_{date_range}.xlsx")
        with pd.ExcelWriter(output_file) as writer:
            store_data.to_excel(writer, sheet_name=f'{store}_Sales', index=False)
            summary_df.to_excel(writer, sheet_name='Summary', index=False)

        # Format columns
        wb = load_workbook(output_file)
        for sname in wb.sheetnames:
            sheet = wb[sname]
            for col in sheet.columns:
                mlen = max(len(str(cell.value)) for cell in col if cell.value is not None)
                sheet.column_dimensions[col[0].column_letter].width = mlen + 2
            for cell in sheet[1]:
                cell.font = Font(bold=True)
                cell.alignment = Alignment(horizontal='center')
        wb.save(output_file)

        total_owed = summary_df['kickback amount'].iloc[0]
        results_for_app.append({
            "store": store,
            "brand": brand,
            "owed": float(total_owed),
            "start": start_date,
            "end": end_date
        })

    return results_for_app