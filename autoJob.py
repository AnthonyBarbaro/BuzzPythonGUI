#!/usr/bin/env python3

import os
import re
import subprocess
import time
import traceback
import datetime
import calendar
from datetime import date, timedelta, datetime as dt
from pathlib import Path

##############################################################################
# 1) LOGIC FOR LAST MONDAY TO SUNDAY
##############################################################################
def get_last_monday_sunday():
    """
    Returns (start_date, end_date) as Python date objects, representing
    last Monday through Sunday.

    Example: if today is Monday 2025-01-20, 
    this returns Monday 2025-01-13 and Sunday 2025-01-19.
    """
    today = date.today()
    # Monday of THIS current week:
    monday_this_week = today - timedelta(days=today.weekday())
    # Last Monday is 7 days before the Monday of this week
    last_monday = monday_this_week - timedelta(days=7)
    # Last Sunday is last_monday + 6 days
    last_sunday = last_monday + timedelta(days=6)
    return last_monday, last_sunday


##############################################################################
# 2) GETCATALOG LOGIC
##############################################################################
def run_get_catalog():
    """
    Calls getCatalog.py with a simple subprocess. 
    Make sure getCatalog.py is in the same directory or specify the full path.
    """
    print("\n===== Running getCatalog.py to download Catalog files... =====\n")
    try:
        subprocess.check_call(["python", "getCatalog.py"])
    except subprocess.CalledProcessError as e:
        print(f"[ERROR] getCatalog.py failed: {e}")
    except FileNotFoundError:
        print("[ERROR] getCatalog.py not found. Please check the script name/path.")


##############################################################################
# 3) GETSALESREPORT LOGIC (HEADLESS, NO GUI)
#    We replicate the essential parts of your getSalesReport.py but skip GUI.
##############################################################################
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

# Credentials for Dutchie Backoffice login
# Make sure you have login.py or define them here:
try:
    from login import username, password
except ImportError:
    username = "YOUR_USERNAME"
    password = "YOUR_PASSWORD"

store_abbr_map = {
    "Buzz Cannabis - Mission Valley": "MV",
    "Buzz Cannabis-La Mesa": "LM"
}

def launch_sales_browser():
    """
    Launches Chrome in 'headless' or regular mode to access the Sales Report page.
    """
    files_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "files")
    os.makedirs(files_dir, exist_ok=True)

    chrome_options = Options()
    chrome_options.add_argument("start-maximized")
    # You can run headless if you like:
    # chrome_options.add_argument("--headless")

    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

    prefs = {
        "download.default_directory": files_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get("https://dusk.backoffice.dutchie.com/reports/sales/sales-report")
    return driver

def login_dutchie(driver):
    wait = WebDriverWait(driver, 15)
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='auth_input_username']"))).send_keys(username)
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='auth_input_password']"))).send_keys(password)
    login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-testid='auth_button_go-green']")))
    login_button.click()
    time.sleep(2)

def click_store_dropdown(driver):
    wait = WebDriverWait(driver, 10)
    dropdown_xpath = "/html/body/div[1]/div/div[1]/div[2]/div[2]/div"
    try:
        wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.sc-ppyJt.jlHGrm")))
    except TimeoutException:
        pass
    dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
    driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", dropdown)
    dropdown.click()
    time.sleep(1)

def select_store(driver, store_name):
    wait = WebDriverWait(driver, 10)
    try:
        click_store_dropdown(driver)
        xpath = f"//li[@data-testid='rebrand-header_menu-item_{store_name}']"
        item = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        item.click()
        print(f"Selected store: {store_name}")
        time.sleep(1)
        return True
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error selecting store '{store_name}': {e}")
        return False

def set_date_range(driver, start_date, end_date):
    """
    Sets the 'Start Date' and 'End Date' fields on the Sales Report page,
    then clicks 'Run'.
    """
    wait = WebDriverWait(driver, 10)
    start_input_str = start_date.strftime("%m/%d/%Y")
    end_input_str   = end_date.strftime("%m/%d/%Y")

    date_inputs = wait.until(EC.presence_of_all_elements_located((By.ID, "input-input_")))
    if len(date_inputs) < 2:
        raise Exception("Could not find start/end date inputs on the page.")

    # Clear and input start date
    date_inputs[0].send_keys(Keys.CONTROL, "a")
    date_inputs[0].send_keys(Keys.DELETE)
    date_inputs[0].send_keys(start_input_str)

    # Clear and input end date
    date_inputs[1].send_keys(Keys.CONTROL, "a")
    date_inputs[1].send_keys(Keys.DELETE)
    date_inputs[1].send_keys(end_input_str)

    print(f"Date range set: {start_input_str} to {end_input_str}")

    # Click "Run"
    run_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Run')]")))
    run_button.click()
    print("Clicked 'Run' button...")
    time.sleep(2)

def export_sales_report(driver, store_name):
    """
    After setting date range and seeing results, clicks "Actions" -> "Export"
    Waits for file to download, renames it accordingly to salesMV.xlsx or salesLM.xlsx
    """
    wait = WebDriverWait(driver, 30)
    files_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "files")

    # State before download
    before_files = set(os.listdir(files_dir))

    try:
        # The "Actions" button
        actions_button = wait.until(EC.element_to_be_clickable((By.ID, 'actions-menu-button')))
        actions_button.click()
        time.sleep(1)

        # The "Export" option
        export_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(),'Export')]")))
        export_option.click()
        time.sleep(3)  # wait for file to appear in folder

        # Wait up to 120 seconds for new file
        start_time = time.time()
        new_file = None
        while time.time() - start_time < 120:
            after_files = set(os.listdir(files_dir))
            added_files = after_files - before_files
            # exclude partial .crdownload
            added_files = {f for f in added_files if not f.endswith('.crdownload')}
            if added_files:
                # assume the first is our new file
                new_file = list(added_files)[0]
                break
            time.sleep(1)

        if not new_file:
            print("No new file downloaded within 2 minutes for store:", store_name)
            return

        # rename it
        old_path = os.path.join(files_dir, new_file)
        if store_name == "Buzz Cannabis - Mission Valley":
            new_name = "salesMV.xlsx"
        else:
            new_name = "salesLM.xlsx"

        new_path = os.path.join(files_dir, new_name)
        os.rename(old_path, new_path)
        print(f"Exported {store_name} sales data to {new_name}")
    except Exception as e:
        print(f"[ERROR] export_sales_report failed for {store_name}:", e)

def run_sales_report(start_date, end_date):
    """
    Complete function that:
    - Launches browser
    - Logs in
    - For each store in store_names, sets date range, exports sales to an .xlsx
    - Quits the browser
    """
    store_names = [
        "Buzz Cannabis - Mission Valley",
        "Buzz Cannabis-La Mesa"
    ]
    driver = launch_sales_browser()
    login_dutchie(driver)

    for store_name in store_names:
        ok = select_store(driver, store_name)
        if not ok:
            continue
        set_date_range(driver, start_date, end_date)
        export_sales_report(driver, store_name)

    driver.quit()


##############################################################################
# 4) RUN deals.py (Brand-Level Deals Report)
##############################################################################
# For brevity, we'll paste your deals.py code as a function:

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

def style_summary_sheet(sheet, brand_name):
    max_col = sheet.max_column
    max_row = sheet.max_row

    sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max_col)
    title_cell = sheet.cell(row=1, column=1)
    title_cell.value = f"{brand_name.upper()} SUMMARY REPORT"
    title_cell.font = Font(name="Calibri", size=16, bold=True, color="FFFFFF")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    title_cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

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

    sheet.freeze_panes = "A3"

    for row_idx in range(3, max_row + 1):
        for col_idx in range(1, max_col + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            hdr_val = sheet.cell(row=header_row_idx, column=col_idx).value
            if hdr_val and ("owed" in str(hdr_val).lower()):
                cell.number_format = '"$"#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif hdr_val and ("date" in str(hdr_val).lower()):
                cell.number_format = "YYYY-MM-DD"
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

            if row_idx % 2 == 1:
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

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
    max_col = sheet.max_column
    for col_idx in range(1, max_col + 1):
        cell = sheet.cell(row=1, column=col_idx)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        max_length = 0
        for row_idx in range(1, sheet.max_row + 1):
            val = sheet.cell(row=row_idx, column=col_idx).value
            try:
                max_length = max(max_length, len(str(val)) if val else 0)
            except:
                pass
        sheet.column_dimensions[col_letter].width = max_length + 2
    sheet.freeze_panes = "A2"

def style_top_sellers_sheet(sheet):
    thin_border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    max_col = sheet.max_column

    for col_idx in range(1, max_col + 1):
        cell = sheet.cell(row=1, column=col_idx)
        cell.font = Font(name="Calibri", size=12, bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    for row_idx in range(2, sheet.max_row + 1):
        for col_idx in range(1, max_col + 1):
            cell = sheet.cell(row=row_idx, column=col_idx)
            cell.border = thin_border
            if col_idx == 2:
                cell.number_format = '"$"#,##0.00'
                cell.alignment = Alignment(horizontal="right")
            else:
                cell.alignment = Alignment(horizontal="left")
            if row_idx % 2 == 1:
                cell.fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    for col_idx in range(1, max_col + 1):
        col_letter = get_column_letter(col_idx)
        max_length = 0
        for row_idx in range(1, sheet.max_row + 1):
            val = sheet.cell(row=row_idx, column=col_idx).value
            try:
                max_length = max(max_length, len(str(val)) if val else 0)
            except:
                pass
        sheet.column_dimensions[col_letter].width = max_length + 2


def run_deals():
    """
    This runs your deals.py main function (run_deals_reports).
    It will read salesMV.xlsx, salesLM.xlsx from 'files' directory,
    generate brand reports in 'brand_reports'.
    """
    # Here is a direct inline copy of your run_deals_reports() logic:
    print("\n===== Running deals.py logic to generate brand-level deals reports... =====\n")
    import locale
    locale.setlocale(locale.LC_ALL, '')

    # We rely on 'files/salesMV.xlsx' and 'files/salesLM.xlsx' existing
    # (exported from run_sales_report above).

    # We define brand_criteria inline, same as your deals.py
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
        data['discount amount'] = data['gross sales'] * discount
        data['kickback amount'] = data['inventory cost'] * kickback
        return data

    mv_data = process_file('files/salesMV.xlsx')
    lm_data = process_file('files/salesLM.xlsx')
    if mv_data is None or lm_data is None:
        print("One or both sales files missing; no data to process in deals.py.")
        return

    ALL_DAYS = {"Monday","Tuesday","Wednesday","Thursday","Friday","Saturday","Sunday"}
    consolidated_summary = []
    results_for_app = []

    output_dir = 'brand_reports'
    Path(output_dir).mkdir(exist_ok=True)

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

        # Date range logic
        start_mv = mv_brand_data['order time'].min() if not mv_brand_data.empty else None
        end_mv   = mv_brand_data['order time'].max() if not mv_brand_data.empty else None
        start_lm = lm_brand_data['order time'].min() if not lm_brand_data.empty else None
        end_lm   = lm_brand_data['order time'].max() if not lm_brand_data.empty else None

        possible_starts = [d for d in [start_mv, start_lm] if d]
        possible_ends   = [d for d in [end_mv, end_lm] if d]
        if not possible_starts or not possible_ends:
            continue

        start_date = min(possible_starts).strftime('%Y-%m-%d')
        end_date   = max(possible_ends).strftime('%Y-%m-%d')
        date_range = f"{start_date}_to_{end_date}"

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

        brand_summary = pd.concat([mv_summary, lm_summary], ignore_index=True)

        if set(criteria['days']) == ALL_DAYS:
            days_text = "Everyday"
        else:
            days_text = ", ".join(criteria['days'])

        brand_summary.rename(columns={'location': 'Store',
                                      'kickback amount': 'Kickback Owed'},
                             inplace=True)
        brand_summary['Days Active'] = days_text
        brand_summary['Date Range']  = f"{start_date} to {end_date}"
        brand_summary['Brand']       = brand

        col_order = ['Store', 'Kickback Owed', 'Days Active', 'Date Range',
                     'gross sales', 'inventory cost', 'discount amount', 'Brand']
        final_cols = [c for c in col_order if c in brand_summary.columns]
        brand_summary = brand_summary[final_cols]
        consolidated_summary.append(brand_summary)

        combined_df = pd.concat([mv_brand_data, lm_brand_data], ignore_index=True)
        top_sellers_df = (combined_df.groupby('product name', as_index=False)
                          .agg({'gross sales': 'sum'})
                          .sort_values(by='gross sales', ascending=False)
                          .head(10))
        top_sellers_df.rename(columns={'product name': 'Product Name', 'gross sales': 'Gross Sales'}, inplace=True)

        safe_brand_name = brand.replace("/", " ")
        output_filename = os.path.join(output_dir, f"{safe_brand_name}_report_{date_range}.xlsx")

        with pd.ExcelWriter(output_filename) as writer:
            brand_summary.to_excel(writer, sheet_name='Summary', index=False, startrow=1)
            mv_brand_data.to_excel(writer, sheet_name='MV_Sales', index=False)
            lm_brand_data.to_excel(writer, sheet_name='LM_Sales', index=False)
            top_sellers_df.to_excel(writer, sheet_name='Top Sellers', index=False)

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

        total_owed = brand_summary['Kickback Owed'].sum()
        results_for_app.append({
            "brand": brand,
            "owed": float(total_owed),
            "start": start_date,
            "end": end_date
        })

    if consolidated_summary:
        final_df = pd.concat(consolidated_summary, ignore_index=True)
        overall_start = final_df['Date Range'].apply(lambda x: x.split(' to ')[0]).min()
        overall_end   = final_df['Date Range'].apply(lambda x: x.split(' to ')[1]).max()
        overall_range = f"{overall_start}_to_{overall_end}"
        consolidated_file = os.path.join(output_dir, f"consolidated_brand_report_{overall_range}.xlsx")

        with pd.ExcelWriter(consolidated_file) as writer:
            final_df.to_excel(writer, sheet_name='Consolidated_Summary', index=False)

        wb = load_workbook(consolidated_file)
        if 'Consolidated_Summary' in wb.sheetnames:
            sheet = wb['Consolidated_Summary']
            # We'll pick the last brand name for styling or pass an arbitrary string
            style_summary_sheet(sheet, "All Brands")
        wb.save(consolidated_file)

        print("Individual brand reports + consolidated report saved to 'brand_reports/'.")
    else:
        print("No brand data found; no Excel files generated in deals.py.")


##############################################################################
# 5) RUN BRAND_INVENTORY.PY FOR 'Hashish' ONLY
##############################################################################
def run_brand_inventory_hashish():
    """
    We'll replicate a minimal version of brand_inventory.py logic,
    forcing brand='Hashish' only.
    We'll assume we want to parse the 'files' directory for new CSVs,
    output to 'done', and only keep lines for brand 'Hashish'.
    Then we apply openpyxl formatting (freeze panes, column widths, row height).
    """
    print("\n===== Running brand_inventory.py logic ONLY for brand='Hashish'... =====\n")

    import pandas as pd
    import re
    from datetime import datetime as dt
    from openpyxl import load_workbook
    from openpyxl.styles import Font, Alignment
    
    input_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "files")
    output_directory = os.path.join(os.path.dirname(os.path.abspath(__file__)), "done")
    os.makedirs(output_directory, exist_ok=True)

    def is_empty_or_numbers(val):
        if not isinstance(val, str):
            return True
        val_str = val.strip()
        return val_str == "" or val_str.isdigit()

    def extract_strain_type(product_name: str):
        if not isinstance(product_name, str):
            return ""
        name = " " + product_name.upper() + " "
        if re.search(r'\bS\b', name):
            return 'S'
        if re.search(r'\bH\b', name):
            return 'H'
        if re.search(r'\bI\b', name):
            return 'I'
        return ""

    def extract_product_details(product_name: str):
        if not isinstance(product_name, str):
            return "", ""
        name_upper = product_name.upper()
        weight_match = re.search(r'(\d+(\.\d+)?)G', name_upper)
        weight = weight_match.group(0) if weight_match else ""
        sub_type = ""
        if " HH " in f" {name_upper} ":
            sub_type = "HH"
        elif " IN " in f" {name_upper} ":
            sub_type = "IN"
        return weight, sub_type

    INPUT_COLUMNS = ['Available', 'Product', 'Category', 'Brand']

    for filename in os.listdir(input_directory):
        if filename.lower().endswith('.csv'):
            file_path = os.path.join(input_directory, filename)
            try:
                df = pd.read_csv(file_path)
            except Exception as e:
                print(f"[ERROR] reading CSV {filename}: {e}")
                continue

            # Filter to required columns
            use_cols = [c for c in INPUT_COLUMNS if c in df.columns]
            if not use_cols:
                continue
            df = df[use_cols]

            # Only brand=Hashish
            if 'Brand' in df.columns:
                df = df[df['Brand'] == 'Hashish']

            # Separate available vs. unavailable
            if 'Available' not in df.columns:
                continue
            unavailable_data = df[df['Available'] == 0]
            available_data   = df[df['Available'] != 0]

            # Parse product columns for the 'available' subset
            if not available_data.empty and 'Product' in available_data.columns:
                available_data['Strain_Type'] = available_data['Product'].apply(extract_strain_type)
                available_data[['Product_Weight','Product_SubType']] = available_data['Product'].apply(
                    lambda x: pd.Series(extract_product_details(x))
                )
                # Remove rows with empty or numeric product name
                available_data = available_data[~available_data['Product'].apply(is_empty_or_numbers)]

            # Sort by Category, Strain_Type, Product_Weight, Product_SubType, and Product
            sort_cols = []
            if 'Category' in available_data.columns:
                sort_cols.append('Category')
            sort_cols += ['Strain_Type','Product_Weight','Product_SubType']
            if 'Product' in available_data.columns:
                sort_cols.append('Product')
            available_data.sort_values(by=sort_cols, inplace=True, na_position='last')

            # Prepare output path
            base_name = os.path.splitext(filename)[0]  # e.g. "GreenHalo"
            today_str = dt.now().strftime("%m-%d-%Y")
            out_subdir = os.path.join(output_directory, base_name)
            os.makedirs(out_subdir, exist_ok=True)

            # --- Construct final Excel filename with "Hashish" + "_" + <inputfile base name> ---
            # e.g. "Hashish_GreenHalo.xlsx"
            out_file = os.path.join(out_subdir, f"Hashish_{base_name}.xlsx")

            # Write to Excel using pandas
            with pd.ExcelWriter(out_file) as writer:
                available_data.to_excel(writer, index=False, sheet_name='Available')
                if not unavailable_data.empty:
                    unavailable_data.to_excel(writer, index=False, sheet_name='Unavailable')

            # Apply formatting with openpyxl
            workbook = load_workbook(out_file)
            for sheet_name in workbook.sheetnames:
                sheet = workbook[sheet_name]

                # Freeze the first row
                sheet.freeze_panes = "A2"

                # Auto-adjust column widths
                for column in sheet.columns:
                    max_length = max(len(str(cell.value)) if cell.value is not None else 0 
                                     for cell in column)
                    sheet.column_dimensions[column[0].column_letter].width = max_length + 2

                # Set a default row height
                for row in sheet.iter_rows():
                    sheet.row_dimensions[row[0].row].height = 17

                # Make the first row bold & center-aligned
                for cell in sheet["1:1"]:
                    cell.font = Font(bold=True)
                    cell.alignment = Alignment(horizontal='center')

            workbook.save(out_file)
            print(f"Hashish brand inventory saved & formatted -> {out_file}")


##############################################################################
# 6) GOOGLE DRIVE UPLOADER
##############################################################################
def run_drive_upload():
    """
    Upload brand_reports/*.xlsx + any done/**/*Hashish_*.xlsx to 
    Google Drive folder "2025_Kickback -> <week range>", 
    writing all links into links.txt
    """
    print("\n===== Running googleDriveUploader logic... =====\n")
    import google.auth
    import google.auth.transport.requests
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    from google.oauth2.credentials import Credentials

    SCOPES = ['https://www.googleapis.com/auth/drive.file']
    LINKS_FILE = "links.txt"
    PARENT_FOLDER_NAME = "2025_Kickback"
    REPORTS_FOLDER = "brand_reports"

    def authenticate_drive_api():
        creds = None
        token_file = "token.json"
        if os.path.exists(token_file):
            creds = Credentials.from_authorized_user_file(token_file, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(google.auth.transport.requests.Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
                creds = flow.run_local_server(port=0)
            with open(token_file, "w") as token:
                token.write(creds.to_json())
        return build("drive","v3", credentials=creds)

    def get_week_range_str():
        lm, ls = get_last_monday_sunday()
        return f"{lm.strftime('%b %d')} to {ls.strftime('%b %d')}"

    def find_or_create_folder(service, folder_name, parent_id=None):
        query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}'"
        if parent_id:
            query += f" and '{parent_id}' in parents"
        resp = service.files().list(q=query, spaces='drive', fields='files(id,name)').execute()
        items = resp.get('files', [])
        if items:
            return items[0]['id']
        else:
            meta = {
                "name": folder_name,
                "mimeType": "application/vnd.google-apps.folder"
            }
            if parent_id:
                meta["parents"] = [parent_id]
            f = service.files().create(body=meta, fields="id").execute()
            return f["id"]

    def upload_file(service, path, parent_id):
        fname = os.path.basename(path)
        body = {
            "name": fname,
            "parents": [parent_id]
        }
        media = MediaFileUpload(path, resumable=True)
        f = service.files().create(body=body, media_body=media, fields="id").execute()
        return f["id"]

    def make_public(service, file_id):
        try:
            perm = {"type":"anyone","role":"reader"}
            service.permissions().create(fileId=file_id, body=perm).execute()
            info = service.files().get(fileId=file_id, fields="webViewLink").execute()
            return info.get("webViewLink")
        except:
            return None

    service = authenticate_drive_api()
    parent_id = find_or_create_folder(service, PARENT_FOLDER_NAME, None)

    week_range = get_week_range_str()
    week_folder_id = find_or_create_folder(service, week_range, parent_id)

    with open(LINKS_FILE,"w") as lf:
        # 1) Upload brand_reports
        if os.path.isdir(REPORTS_FOLDER):
            for fname in os.listdir(REPORTS_FOLDER):
                if fname.endswith(".xlsx"):
                    full_path = os.path.join(REPORTS_FOLDER,fname)
                    # Upload to the *same* week folder (no sub-subfolders)
                    file_id = upload_file(service, full_path, week_folder_id)
                    link = make_public(service, file_id)
                    if link:
                        lf.write(f"{fname}: {link}\n")
                        print(f"Uploaded {fname} => {link}")

        # 2) Also upload done/**/*Hashish_*.xlsx
        done_dir = "done"
        if os.path.isdir(done_dir):
            for root, dirs, files in os.walk(done_dir):
                for f in files:
                    if f.endswith(".xlsx") and "Hashish_" in f:
                        full_path = os.path.join(root,f)
                        file_id = upload_file(service, full_path, week_folder_id)
                        link = make_public(service, file_id)
                        if link:
                            lf.write(f"{f}: {link}\n")
                            print(f"Uploaded {f} => {link}")

    print("All files uploaded. Links stored in links.txt.")


##############################################################################
# 7) EMAIL THE LINKS.TXT + HASHISH BRAND REPORT (Optional)
##############################################################################
def send_email_with_gmail(subject, body, recipients, attachments=None):
    """
    Sends an email (plain text) via Gmail API with optional attachments.
    """
    print("\n===== Sending Email via Gmail API... =====\n")
    import base64
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email.mime.text import MIMEText
    from email.utils import formatdate
    from email import encoders
    import google.auth.transport.requests
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    GMAIL_SCOPES = ['https://www.googleapis.com/auth/gmail.send']

    creds = None
    gmail_token = "token_gmail.json"
    if os.path.exists(gmail_token):
        creds = Credentials.from_authorized_user_file(gmail_token, GMAIL_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(google.auth.transport.requests.Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(gmail_token,"w") as t:
            t.write(creds.to_json())

    service = build('gmail','v1', credentials=creds)

    if isinstance(recipients, str):
        recipients = [recipients]

    msg = MIMEMultipart()
    msg['From'] = "me"
    msg['To'] = ", ".join(recipients)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    if attachments:
        for path in attachments:
            if not os.path.isfile(path):
                continue
            fn = os.path.basename(path)
            with open(path,"rb") as f:
                part = MIMEBase("application","octet-stream")
                part.set_payload(f.read())
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{fn}"')
            msg.attach(part)

    raw_msg = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    send_msg = {'raw': raw_msg}
    try:
        sent = service.users().messages().send(userId='me', body=send_msg).execute()
        print(f"Email sent! ID: {sent['id']}")
    except Exception as e:
        print("[ERROR] Could not send Gmail:", e)
def send_email_with_gmail_html(subject, html_body, recipients, attachments=None):
    """
    Sends an HTML email via Gmail API with optional attachments.
    """
    import base64
    from email.mime.multipart import MIMEMultipart
    from email.mime.base import MIMEBase
    from email.mime.text import MIMEText
    from email.utils import formatdate
    from email import encoders
    import google.auth.transport.requests
    from google_auth_oauthlib.flow import InstalledAppFlow
    from google.oauth2.credentials import Credentials
    from googleapiclient.discovery import build

    GMAIL_SCOPES = ['https://www.googleapis.com/auth/gmail.send']
    creds = None
    gmail_token = "token_gmail.json"

    if os.path.exists(gmail_token):
        creds = Credentials.from_authorized_user_file(gmail_token, GMAIL_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(google.auth.transport.requests.Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(gmail_token, "w") as f:
            f.write(creds.to_json())

    service = build('gmail', 'v1', credentials=creds)

    if isinstance(recipients, str):
        recipients = [recipients]

    # Create a MIMEMultipart message for HTML
    msg = MIMEMultipart('alternative')
    msg['From'] = "me"
    msg['To'] = ", ".join(recipients)
    msg['Date'] = formatdate(localtime=True)
    msg['Subject'] = subject

    # Attach the HTML body
    part_html = MIMEText(html_body, 'html')
    msg.attach(part_html)

    # Optionally attach files
    if attachments:
        for file_path in attachments:
            if not os.path.isfile(file_path):
                continue
            filename = os.path.basename(file_path)
            with open(file_path, "rb") as fp:
                file_data = fp.read()
            part = MIMEBase("application", "octet-stream")
            part.set_payload(file_data)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{filename}"')
            msg.attach(part)

    raw_message = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    send_req = {'raw': raw_message}

    try:
        sent = service.users().messages().send(userId='me', body=send_req).execute()
        print(f"HTML Email sent! ID: {sent['id']}")
    except Exception as e:
        print("[ERROR] Could not send HTML email:", e)


##############################################################################
# MAIN: ORCHESTRATE ALL STEPS
##############################################################################
def main():
    print("===== Starting autoJob.py =====")

    last_monday, last_sunday = get_last_monday_sunday()
    date_range_str = f"{last_monday} to {last_sunday}"
    print(f"Processing for last week range: {date_range_str}")

    # 1) Catalog
    run_get_catalog()

    # 2) Sales
    run_sales_report(last_monday, last_sunday)

    # 3) Deals
    run_deals()
    time.sleep(2)

    # 4) Brand Inventory (Hashish)
    run_brand_inventory_hashish()
    time.sleep(2)

    # 5) Drive Upload (both brand_reports + done/Hashish)
    run_drive_upload()

    # 6) TWO EMAILS

    # 6a) Parse links.txt to separate "Hashish" lines from "non-Hashish" lines,
    #     and build HTML bullet lists for each group.

    links_file = "links.txt"
    hashish_links = []
    non_hashish_links = []

    if os.path.exists(links_file):
        with open(links_file, "r", encoding="utf-8") as lf:
            lines = lf.readlines()
        for line in lines:
            line = line.strip()
            # Typical format: "filename.xlsx: https://drive.google.com/..."
            if "Hashish_" in line:
                hashish_links.append(line)
            else:
                non_hashish_links.append(line)

    # Convert lines to HTML bullet points
    def make_html_list(link_lines):
        """
        Each line is 'filename.xlsx: https://...'
        We'll parse into <li><b>filename.xlsx</b>: <a href='url'>url</a></li>
        """
        if not link_lines:
            return "<p>No links found.</p>"
        html_list = "<ul>\n"
        for line in link_lines:
            if ":" in line:
                filename, url = line.split(":", 1)
                filename = filename.strip()
                url = url.strip()
                # Make an HTML bullet with clickable link
                html_list += f"<li><strong>{filename}</strong>: <a href='{url}'>{url}</a></li>\n"
            else:
                html_list += f"<li>{line}</li>\n"
        html_list += "</ul>\n"
        return html_list

    non_hashish_html = make_html_list(non_hashish_links)
    hashish_html = make_html_list(hashish_links)

    # 6b) For the second email, we want to also parse the "Hashish" brand report summary
    #     to get Store + Kickback Owed. We'll load whichever "Hashish_*.xlsx" brand report
    #     was uploaded. If you have multiple, you may need additional logic.

    import openpyxl

    hashish_summary_rows = []
    brand_reports_dir = "brand_reports"
    hashish_report_path = None

    if os.path.isdir(brand_reports_dir):
        # We pick the first Hashish_ file. If you have multiple, adjust logic.
        for f in os.listdir(brand_reports_dir):
            if f.startswith("Hashish_") and f.endswith(".xlsx"):
                hashish_report_path = os.path.join(brand_reports_dir, f)
                break

    if hashish_report_path and os.path.isfile(hashish_report_path):
        # Open "Summary" sheet and read columns A (Store) & B (Kickback Owed)
        wb = openpyxl.load_workbook(hashish_report_path, data_only=True)
        if "Summary" in wb.sheetnames:
            sh = wb["Summary"]
            # We'll gather (Store, Kickback Owed) from row 2 downward
            # (row 1 is typically headers: "Store", "Kickback Owed", etc.)
            for row_idx in range(2, sh.max_row + 1):
                store_val = sh.cell(row=row_idx, column=1).value  # Column A
                owed_val  = sh.cell(row=row_idx, column=2).value  # Column B
                if store_val is not None and owed_val is not None:
                    hashish_summary_rows.append((store_val, owed_val))
        wb.close()

    # Build an HTML table for the Store / Kickback Owed data
    def make_hashish_owed_table(rows):
        """
        rows should be a list of tuples: [(store, owed), (store, owed), ...]
        where 'store' might be "Misson Valley" / "La Mesa" and 'owed' is numeric or string.
        
        This function:
        - Skips any row if store == "Store" or owed == "Kickback Owed" (i.e. a header).
        - Converts owed to a currency string with 2 decimals.
        - Builds an HTML table of <Store, Kickback Owed>.
        """
        if not rows:
            return "<p>No store / owed data found.</p>"
        
        # Filter out any row that looks like the header or empty data
        filtered_data = []
        for (store, owed) in rows:
            # If store or owed is literally "Store", "Kickback Owed", or empty, skip
            if not store or not owed:
                continue
            if str(store).lower() in ["store", "kickback owed"]:
                continue
            if str(owed).lower() in ["store", "kickback owed"]:
                continue

            # Attempt to format owed as currency
            try:
                owed_float = float(owed)
                owed_str = f"${owed_float:,.2f}"  # e.g. $876.92
            except (ValueError, TypeError):
                # fallback if it's not numeric
                owed_str = str(owed)

            filtered_data.append((str(store), owed_str))

        if not filtered_data:
            return "<p>No store / owed data after removing headers/empties.</p>"

        html = """
        <table border='1' cellpadding='5' cellspacing='0'>
        <thead>
            <tr><th>Store</th><th>Kickback Owed</th></tr>
        </thead>
        <tbody>
        """

        for (store_val, owed_val) in filtered_data:
            html += f"<tr><td>{store_val}</td><td>{owed_val}</td></tr>\n"

        html += "</tbody></table>\n"
        return html


    hashish_owed_table_html = make_hashish_owed_table(hashish_summary_rows)

    # ============= EMAIL #1 => "Hello Donna" (Non-Hashish links in body) =============
    # We'll embed non-hashish links in the email body (HTML).
    email1_subject = f"Weekly Kickback Links (Donna) ({date_range_str})"
    email1_body_html = f"""
    <html>
    <body>
    <p>Hello Donna,</p>
    <p>Please find below the Google Drive links for the week {date_range_str}:</p>
    {non_hashish_html}
    <p><strong>please include/contact anthony@buzzcannabis.com in all emails regarding these credits. </strong></p>
    <p>Regards,<br>Anthony</p>
    </body>
    </html>
    """

    # We'll send an HTML email so the <a href='...'> is clickable
    send_email_with_gmail_html(
        subject=email1_subject,
        html_body=email1_body_html,
        recipients=["anthony@barbaro.tech"],
        attachments=None  # or pass a list of attachments if you want any files
    )

    # ============= EMAIL #2 => "Hello Ryan" (Hashish links + owed data) =============
    # We'll embed hashish links and the store/owed table in the email body (HTML).
    email2_subject = f"Weekly Kickback (Ryan) (Hashish) ({date_range_str})"
    email2_body_html = f"""
    <html>
    <body>
    <p>Hello Ryan,</p>
    <p>Here are the <strong>Hashish brand</strong> links for {date_range_str}:</p>
    {hashish_html}
    <p>Additionally, the store / Kickback Owed data is:</p>
    {hashish_owed_table_html}
    <p>Regards,<br>Anthony</p>
    </body>
    </html>
    """

    send_email_with_gmail_html(
        subject=email2_subject,
        html_body=email2_body_html,
        recipients=["anthony@barbaro.tech"],
        attachments=None
    )

    print("\n===== autoJob.py completed successfully. =====")


if __name__ == "__main__":
    main()