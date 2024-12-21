import os
import re
import subprocess
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.ttk import Combobox
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
import traceback
from datetime import datetime, timedelta
import calendar
import time
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from login import username, password

CONFIG_FILE = "config.txt"
INPUT_COLUMNS = ['Available', 'Product', 'Category', 'Brand']

store_abbr_map = {
    "Buzz Cannabis - Mission Valley": "MV",
    "Buzz Cannabis-La Mesa": "LM"
}

start_str = None
end_str = None
driver = None

def wait_for_new_file(download_directory, before_files, timeout=30):
    end_time = time.time() + timeout
    while time.time() < end_time:
        after_files = set(os.listdir(download_directory))
        new_files = after_files - before_files
        if new_files:
            return list(new_files)[0]
        time.sleep(1)
    return None

def launchBrowser():
    files_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "files")
    os.makedirs(files_dir, exist_ok=True)

    chrome_options = Options()
    chrome_options.add_argument("start-maximized")
    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])

    prefs = {
        "download.default_directory": files_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)

    d = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    d.get("https://dusk.backoffice.dutchie.com/reports/sales/sales-report")
    return d

def login():
    wait = WebDriverWait(driver, 10)
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='auth_input_username']"))).send_keys(username)
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='auth_input_password']"))).send_keys(password)
    login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-testid='auth_button_go-green']")))
    login_button.click()
    time.sleep(1)

def click_dropdown():
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

def select_dropdown_item(item_text):
    wait = WebDriverWait(driver, 10)
    try:
        click_dropdown()
        xpath = f"//li[@data-testid='rebrand-header_menu-item_{item_text}']"
        item = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        item.click()
        print(f"Selected store: {item_text}")
        time.sleep(1)
        return True
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error selecting store '{item_text}': {e}")
        return False

from selenium.webdriver.common.keys import Keys

def set_date_range(start_date, end_date):
    global start_str, end_str

    start_str = start_date.strftime("%m-%d-%Y")
    end_str = end_date.strftime("%m-%d-%Y")

    start_input_str = start_date.strftime("%m/%d/%Y")
    end_input_str = end_date.strftime("%m/%d/%Y")

    wait = WebDriverWait(driver, 10)
    date_inputs = wait.until(EC.presence_of_all_elements_located((By.ID, "input-input_")))

    # Clear and input start date
    date_inputs[0].send_keys(Keys.CONTROL, "a")
    date_inputs[0].send_keys(Keys.DELETE)
    date_inputs[0].send_keys(start_input_str)

    # Clear and input end date
    date_inputs[1].send_keys(Keys.CONTROL, "a")
    date_inputs[1].send_keys(Keys.DELETE)
    date_inputs[1].send_keys(end_input_str)

    print(f"Set date range: {start_input_str} to {end_input_str}")
    time.sleep(1)

def click_run_button():
    wait = WebDriverWait(driver, 10)
    run_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(),'Run')]")))
    run_button.click()
    print("Run button clicked successfully.")
    time.sleep(1)

def clickActionsAndExport(current_store):
    try:
        time.sleep(10)
        wait = WebDriverWait(driver, 10)
        files_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "files")
        before_files = set(os.listdir(files_dir))

        actions_button = wait.until(EC.element_to_be_clickable((By.ID, 'actions-menu-button')))
        actions_button.click()
        print("Actions button clicked successfully.")
        time.sleep(1)
        
        export_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(),'Export')]")))
        export_option.click()
        print("Export option clicked successfully.")
        time.sleep(1)

        export_csv_button = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "body > div.sc-jYnRlT.kGxwGQ.sc-heKhxA.Bgfyt.MuiDialog-root.sc-bBPnyn.hrApOB.MuiModal-root "
                              "> div.sc-fhrEpP.dWiAWv.MuiDialog-container.MuiDialog-scrollPaper "
                              "> div > div.sc-iyxVF.bdkceX.MuiDialogActions-root.MuiDialogActions-spacing.sc-fopvND.finsng "
                              "> div.primary-actions > button:nth-child(1)")))
        export_csv_button.click()
        print("Export CSV button clicked successfully.")

        new_file = wait_for_new_file(files_dir, before_files, timeout=60)
        if new_file:
            print(f"New file downloaded: {new_file}")
            store_abbr = store_abbr_map.get(current_store, "UNK")
            extension = os.path.splitext(new_file)[1]

            # If you want the format:
            # SR_{store_abbr}_{start_str}-{end_str}{extension}
            # Example: SR_MV_12-09-2024-12-15-2024.csv
            new_filename = f"SR_{store_abbr}_{start_str}-{end_str}{extension}"

            original_path = os.path.join(files_dir, new_file)
            new_path = os.path.join(files_dir, new_filename)

            # If file with same name exists, delete it to replace
            if os.path.exists(new_path):
                os.remove(new_path)
                print(f"Removed existing file: {new_filename}")

            try:
                os.rename(original_path, new_path)
                print(f"Renamed {new_file} to {new_filename}")
            except Exception as e:
                print(f"Error renaming file: {e}")
        else:
            print("No new file detected after export.")


        
        time.sleep(1)
    except TimeoutException:
        print("An element could not be found or clicked within the timeout period.")
    except Exception as e:
        print(f"An error occurred during export: {e}")

def update_days_combobox(year_combo, month_combo, day_combo):
    # Weekday abbreviations
    weekday_abbr = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    y = int(year_combo.get())
    m = month_combo.current()+1
    now = datetime.now()
    # last day of month
    last_day = calendar.monthrange(y, m)[1]

    # If current month & year, limit days to today
    if y == now.year and m == now.month:
        last_day = min(last_day, now.day)

    # build list with weekdays
    day_values = []
    for day_num in range(1, last_day+1):
        dt = datetime(y, m, day_num)
        wday_abbr = weekday_abbr[dt.weekday()]  # Monday=0
        day_values.append(f"{day_num} ({wday_abbr})")

    day_combo['values'] = day_values
    # if current selection too large, reset
    current_day_idx = day_combo.current()
    if current_day_idx == -1 or current_day_idx >= len(day_values):
        day_combo.current(0)

def open_gui_and_run():
    root = tk.Tk()
    root.title("Select Date Range")

    this_year = datetime.now().year
    YEAR_RANGE = [str(this_year-1), str(this_year)]
    MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"]

    def create_date_selector(frame, label_text):
        tk.Label(frame, text=label_text, font=("Arial", 12, "bold")).pack(pady=(10,5))
        
        subframe = tk.Frame(frame)
        subframe.pack(pady=5)

        year_combo = Combobox(subframe, values=YEAR_RANGE, state='readonly', width=8)
        year_combo.current(YEAR_RANGE.index(str(this_year)))
        year_combo.grid(row=0, column=0, padx=5)

        month_combo = Combobox(subframe, values=MONTHS, state='readonly', width=10)
        current_month = datetime.now().month
        month_combo.current(current_month-1)
        month_combo.grid(row=0, column=1, padx=5)

        day_combo = Combobox(subframe, state='readonly', width=10)  # widened to accommodate " (Mon)"
        day_combo.grid(row=0, column=2, padx=5)

        def on_year_month_change(*args):
            update_days_combobox(year_combo, month_combo, day_combo)

        year_combo.bind("<<ComboboxSelected>>", on_year_month_change)
        month_combo.bind("<<ComboboxSelected>>", on_year_month_change)

        # initial populate days
        update_days_combobox(year_combo, month_combo, day_combo)
        selected_year = int(year_combo.get())
        selected_month = month_combo.current()+1
        now = datetime.now()
        if selected_year == now.year and selected_month == now.month:
            today_day = now.day
            day_combo.current(today_day-1)
        else:
            day_combo.current(0)

        return year_combo, month_combo, day_combo

    # GUI Layout
    main_frame = tk.Frame(root)
    main_frame.pack(pady=20, padx=20)

    start_year_combo, start_month_combo, start_day_combo = create_date_selector(main_frame, "Select Start Date:")
    end_year_combo, end_month_combo, end_day_combo = create_date_selector(main_frame, "Select End Date:")

    def on_ok():
        sy = int(start_year_combo.get())
        sm = start_month_combo.current()+1
        # day format: "1 (Mon)", split by space
        sday_str = start_day_combo.get().split()[0]
        sd = int(sday_str)

        ey = int(end_year_combo.get())
        em = end_month_combo.current()+1
        eday_str = end_day_combo.get().split()[0]
        ed = int(eday_str)

        start_date = datetime(sy, sm, sd)
        end_date = datetime(ey, em, ed)

        if start_date > end_date:
            messagebox.showerror("Error", "Start date cannot be after End date.")
            return

        root.destroy()

        global driver
        driver = launchBrowser()
        login()

        store_names = ["Buzz Cannabis - Mission Valley", "Buzz Cannabis-La Mesa"]
        for store in store_names:
            if not select_dropdown_item(store):
                break
            set_date_range(start_date, end_date)
            click_run_button()
            clickActionsAndExport(store)

        driver.quit()

    tk.Button(root, text="OK", command=on_ok, font=("Arial", 12, "bold"), bg="lightblue").pack(pady=10)

    root.mainloop()

# Main execution through GUI
open_gui_and_run()
