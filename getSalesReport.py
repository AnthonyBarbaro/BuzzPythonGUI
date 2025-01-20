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

def wait_for_new_file(download_directory, before_files, timeout=12):
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

def monitor_folder_for_new_file(folder_path, before_files, timeout=120):
    """Monitor a folder for new files."""
    start_time = time.time()
    while time.time() - start_time < timeout:
        current_files = set(os.listdir(folder_path))
        new_files = current_files - before_files
        if new_files:
            # Return the first fully downloaded file
            for file in new_files:
                if not file.endswith('.crdownload'):  # Exclude partially downloaded files
                    return file
        time.sleep(1)
    return None

def wait_until_file_is_stable(file_path, stable_time=2, max_wait=30):
    """Wait until a file's size is stable."""
    start_time = time.time()
    last_size = -1
    stable_start = None

    while True:
        try:
            current_size = os.path.getsize(file_path)
        except FileNotFoundError:
            current_size = -1

        if current_size == last_size and current_size != -1:
            if stable_start is None:
                stable_start = time.time()
            elif time.time() - stable_start >= stable_time:
                return True
        else:
            stable_start = None

        last_size = current_size
        if time.time() - start_time > max_wait:
            return False
        time.sleep(0.5)

def clickActionsAndExport(current_store):
    try:
        print(f"\n=== Exporting data for store: {current_store} ===")
        wait = WebDriverWait(driver, 5)
        files_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "files")

        # Capture initial state of the download folder
        before_files = set(os.listdir(files_dir))
        print("Files before download:", before_files)

        # Click the Actions button
        actions_button = wait.until(EC.element_to_be_clickable((By.ID, 'actions-menu-button')))
        actions_button.click()
        print("Actions button clicked successfully.")
        time.sleep(2)

        # Select the Export option
        export_option = wait.until(EC.element_to_be_clickable((By.XPATH, "//li[contains(text(),'Export')]")))
        export_option.click()
        print("Export option clicked successfully.")
        time.sleep(1)

        # Wait for a new file to appear in the directory
        downloaded_file = wait_for_new_file(files_dir, before_files, timeout=120)
        if downloaded_file:
            original_path = os.path.join(files_dir, downloaded_file)
            print(f"New file detected: {downloaded_file}")
            
        # Check the folder contents after download
        time.sleep(1)  # Short wait to ensure the file is written
        after_files = set(os.listdir(files_dir))
        print("Files after download:", after_files)

        # Identify new files
        new_files = after_files - before_files
        if new_files:
            downloaded_file = max(new_files, key=lambda f: os.path.getctime(os.path.join(files_dir, f)))
            print(f"New file detected: {downloaded_file}")
            original_path = os.path.join(files_dir, downloaded_file)

            # Ensure file is stable before renaming
            if wait_until_file_is_stable(original_path):
                # Generate a new filename
                timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
                if current_store == "Buzz Cannabis - Mission Valley":
                    new_filename = f"salesMV.xlsx"
                elif current_store == "Buzz Cannabis-La Mesa":
                    new_filename = f"salesLM.xlsx"
                elif current_store == "Buzz Cannabis - SORRENTO VALLEY":
                    new_filename = f"salesSV.xlsx"
                else:
                    new_filename = f"sales_{current_store}_{timestamp}.xlsx"

                new_path = os.path.join(files_dir, new_filename)

                # Rename the file
                try:
                    os.rename(original_path, new_path)
                    print(f"Renamed file to: {new_filename}")
                except Exception as e:
                    print(f"Error renaming file: {e}")
            else:
                print("File did not stabilize in time.")
        else:
            print("No new file detected after export.")

    except TimeoutException:
        print("An element could not be found or clicked within the timeout period.")
    except Exception as e:
        print(f"An error occurred during export: {traceback.format_exc()}")

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

def create_store_checkboxes(frame):
    """
    Creates three checkboxes for the three store locations,
    with each checkbox selected by default.
    Returns a dict of {store_name: IntVar}.
    """
    store_vars = {}

    # Mission Valley
    varMV = tk.IntVar(value=1)
    cbMV = tk.Checkbutton(frame, text="Buzz Cannabis - Mission Valley", variable=varMV)
    cbMV.pack(anchor='w')
    store_vars["Buzz Cannabis - Mission Valley"] = varMV

    # La Mesa
    varLM = tk.IntVar(value=1)
    cbLM = tk.Checkbutton(frame, text="Buzz Cannabis-La Mesa", variable=varLM)
    cbLM.pack(anchor='w')
    store_vars["Buzz Cannabis-La Mesa"] = varLM

    # Sorrento Valley
    varSV = tk.IntVar(value=1)
    cbSV = tk.Checkbutton(frame, text="Buzz Cannabis - SORRENTO VALLEY", variable=varSV)
    cbSV.pack(anchor='w')
    store_vars["Buzz Cannabis - SORRENTO VALLEY"] = varSV

    return store_vars

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

    # Create date selectors
    start_year_combo, start_month_combo, start_day_combo = create_date_selector(main_frame, "Select Start Date:")
    end_year_combo, end_month_combo, end_day_combo = create_date_selector(main_frame, "Select End Date:")

    # Create checkboxes for selecting stores
    tk.Label(main_frame, text="Select Store(s):", font=("Arial", 12, "bold")).pack(pady=(10,5), anchor='w')
    store_vars = create_store_checkboxes(main_frame)

    def on_ok():
        # Gather date info
        sy = int(start_year_combo.get())
        sm = start_month_combo.current()+1
        sday_str = start_day_combo.get().split()[0]  # "1 (Mon)" -> "1"
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

        # Determine which stores are selected
        selected_stores = []
        for store_name, var in store_vars.items():
            if var.get() == 1:  # if box checked
                selected_stores.append(store_name)

        # If none selected, default to all
        if not selected_stores:
            selected_stores = [
                "Buzz Cannabis - Mission Valley",
                "Buzz Cannabis-La Mesa",
                "Buzz Cannabis - SORRENTO VALLEY"
            ]

        # Close GUI
        root.destroy()

        # Launch browser, login, iterate over stores
        global driver
        driver = launchBrowser()
        login()

        for store in selected_stores:
            if not select_dropdown_item(store):
                break
            set_date_range(start_date, end_date)
            click_run_button()
            clickActionsAndExport(store)

        driver.quit()

    tk.Button(root, text="OK", command=on_ok, font=("Arial", 12, "bold"), bg="lightblue").pack(pady=10)

    root.mainloop()

# Main execution through GUI
if __name__ == "__main__":
    open_gui_and_run()
