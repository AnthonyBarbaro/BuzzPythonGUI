import os
import time
import traceback
from datetime import datetime, timedelta
import calendar

import tkinter as tk
from tkinter import messagebox
from tkinter.ttk import Combobox

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException, ElementClickInterceptedException

# Import your login credentials
from login import username, password


# ---------------------------------------------------------------------
# 1) Original script logic
# ---------------------------------------------------------------------

def launchBrowser():
    """Launch Chrome, go to the Dusk Closing Report page."""
    chrome_options = Options()
    chrome_options.add_argument("start-maximized")
    chrome_options.add_experimental_option("detach", True)
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get("https://dusk.backoffice.dutchie.com/reports/closing-report/registers")
    return driver

def login(driver):
    """Login using your stored credentials."""
    wait = WebDriverWait(driver, 10)
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='auth_input_username']"))).send_keys(username)
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='auth_input_password']"))).send_keys(password)
    login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-testid='auth_button_go-green']")))
    login_button.click()

def click_dropdown(driver):
    """ Click the store dropdown to open the list of store options. """
    wait = WebDriverWait(driver, 10)
    dropdown_xpath = "//div[@data-testid='header_select_location']"
    try:
        dropdown = wait.until(EC.element_to_be_clickable((By.XPATH, dropdown_xpath)))
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", dropdown)
        dropdown.click()
        time.sleep(1)  # small delay for options to load
    except TimeoutException:
        print("Dropdown not found or not clickable")

def select_store(driver, store_name):
    """
    Select the given store from the dropdown menu. 
    store_name should match exactly how it appears in the site’s dropdown.
    """
    wait = WebDriverWait(driver, 10)
    try:
        # 1) Click the dropdown:
        click_dropdown(driver)

        # 2) Locate the store option by text:
        #    Adjust XPATH if needed to match how your site represents store names
        xpath = f"//li[contains(text(), '{store_name}')]"
        item = wait.until(EC.visibility_of_element_located((By.XPATH, xpath)))
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", item)
        item.click()
        time.sleep(1)
        return True
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error while trying to select '{store_name}' from the dropdown: {e}")
        return False

def click_date_input_field(driver):
    """ Click on the date input field to open the date-picker. """
    wait = WebDriverWait(driver, 7)
    date_input_id = "input-input_"
    date_input = wait.until(EC.element_to_be_clickable((By.ID, date_input_id)))
    date_input.click()

def click_dates_in_calendar(driver, day_of_month):
    """
    For the 'closing-report/registers' page, select a single day at a time.
    Then click the 'Run' button to refresh the table, using JavaScript clicks 
    and small waits to avoid intercept issues.
    """
    wait = WebDriverWait(driver, 10)
    try:
        # 1) Click the day in the datepicker (via JS to reduce "intercept" issues)
        day_div_xpath = f"//div[text()='{day_of_month}']"
        day_div = wait.until(EC.element_to_be_clickable((By.XPATH, day_div_xpath)))
        driver.execute_script("arguments[0].click();", day_div)
        time.sleep(0.5)

        # 2) Press ESC to close date picker if it remains open
        body = driver.find_element(By.TAG_NAME, "body")
        body.send_keys(Keys.ESCAPE)
        time.sleep(0.5)

        # OPTIONAL: If you know an overlay is blocking, you can wait for invisibility.
        # Example CSS from your error might be "div.sc-kwdoa-D.loTCwi"
        # try:
        #    wait.until(EC.invisibility_of_element_located((By.CSS_SELECTOR, "div.sc-kwdoa-D.loTCwi")))
        # except TimeoutException:
        #    pass

        # 3) Use JavaScript click on the 'Run' button
        run_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Run')]")))
        driver.execute_script("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", run_button)
        driver.execute_script("arguments[0].click();", run_button)

    except (TimeoutException, ElementClickInterceptedException) as e:
        print("Could not click the day or the Run button.")
        print(f"Error details: {e}")

def extract_monetary_values(driver):
    """
    Extract the first 3 right-aligned table cells and parse them as float. 
    """
    css_selector = "[class$='table-cell-right-']"
    time.sleep(3)
    elements = WebDriverWait(driver, 35).until(
        EC.presence_of_all_elements_located((By.CSS_SELECTOR, css_selector))
    )
    monetary_values = []
    for element in elements[:3]:
        value_text = element.text
        try:
            numeric_value = float(value_text
                .replace('$', '')
                .replace(',', '')
                .replace('(', '-')
                .replace(')', '')
            )
            monetary_values.append(numeric_value)
        except ValueError:
            print(f"Could not convert '{value_text}' to float.")
    return monetary_values

def process_single_day(driver, date_to_run):
    """
    Given a Python date object (date_to_run),
    1) Click into date field,
    2) Select that day in the calendar,
    3) Extract monetary values,
    4) Print out your result (like your original script).
    """
    # Convert day of month to a string (for the datepicker)
    day_str = str(date_to_run.day)

    # Click date input, then click the day in the datepicker
    click_date_input_field(driver)
    click_dates_in_calendar(driver, day_str)

    # Extract values
    gross = extract_monetary_values(driver)

    # Format the date mm/dd
    formatted_date = date_to_run.strftime('%m/%d')
    if len(gross) >= 2:
        print("\033[1m--------------------------------\033[0m")
        print(f"\033[1m{formatted_date} {gross[0]} {gross[1]}\033[0m")
        print("\033[1m--------------------------------\033[0m")
        if gross[0] != 0:
            # ratio example
            print(float((-1 * gross[1]) / gross[0]))
        else:
            print("Gross[0] is zero, cannot compute ratio.")
    else:
        print(f"{formatted_date}: Not enough data to calculate sales.")


# ---------------------------------------------------------------------
# 2) Adding a GUI to pick the date range and stores
# ---------------------------------------------------------------------

def create_store_checkboxes(frame):
    """
    Create 4 checkboxes for the store locations you want to handle,
    each store is checked by default.
    Returns a dict of {store_name: IntVar}.
    """
    store_vars = {}

    store_map = [
        "Buzz Cannabis - Mission Valley",
        "Buzz Cannabis-La Mesa",
        "Buzz Cannabis - SORRENTO VALLEY",
        "Buzz Cannabis - Lemon Grove"
    ]

    for store_name in store_map:
        var = tk.IntVar(value=1)
        cb = tk.Checkbutton(frame, text=store_name, variable=var)
        cb.pack(anchor='w')
        store_vars[store_name] = var

    return store_vars

def update_days_combobox(year_combo, month_combo, day_combo):
    """
    Refreshes the day_combobox when year or month changes,
    taking into account actual days in the selected month/year.
    """
    weekday_abbr = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
    y = int(year_combo.get())
    m = month_combo.current() + 1
    now = datetime.now()

    # last day of that month
    last_day = calendar.monthrange(y, m)[1]
    # If user picked the current month/year, limit day up to 'today'
    if y == now.year and m == now.month:
        last_day = min(last_day, now.day)

    day_values = []
    for day_num in range(1, last_day + 1):
        dt = datetime(y, m, day_num)
        wday_abbr = weekday_abbr[dt.weekday()]
        day_values.append(f"{day_num} ({wday_abbr})")

    day_combo['values'] = day_values
    if day_combo.current() == -1 or day_combo.current() >= len(day_values):
        day_combo.current(0)

def create_date_selector(frame, label_text, year_options):
    """
    Creates a row of combo boxes for picking Year, Month, Day,
    plus a label. Returns (year_combo, month_combo, day_combo).
    """
    tk.Label(frame, text=label_text, font=("Arial", 12, "bold")).pack(pady=(10,5))
    
    subframe = tk.Frame(frame)
    subframe.pack(pady=5)

    year_combo = Combobox(subframe, values=year_options, state='readonly', width=8)
    year_combo.current(len(year_options)-1)  # default to most recent year
    year_combo.grid(row=0, column=0, padx=5)

    MONTHS = ["January","February","March","April","May","June","July","August","September","October","November","December"]
    month_combo = Combobox(subframe, values=MONTHS, state='readonly', width=10)
    current_month = datetime.now().month
    month_combo.current(current_month-1)
    month_combo.grid(row=0, column=1, padx=5)

    day_combo = Combobox(subframe, state='readonly', width=10)
    day_combo.grid(row=0, column=2, padx=5)

    def on_year_month_change(*args):
        update_days_combobox(year_combo, month_combo, day_combo)

    year_combo.bind("<<ComboboxSelected>>", on_year_month_change)
    month_combo.bind("<<ComboboxSelected>>", on_year_month_change)

    # Initial population of days
    update_days_combobox(year_combo, month_combo, day_combo)
    now = datetime.now()
    sel_year = int(year_combo.get())
    sel_month = month_combo.current()+1
    if sel_year == now.year and sel_month == now.month:
        day_combo.current(now.day - 1)
    else:
        day_combo.current(0)

    return year_combo, month_combo, day_combo

def get_date_from_comboboxes(year_combo, month_combo, day_combo):
    """
    Convert the user’s GUI selection into a datetime object.
    """
    y = int(year_combo.get())
    m = month_combo.current() + 1
    d = int(day_combo.get().split()[0])  # "15 (Mon)" -> 15
    return datetime(y, m, d)

def open_gui_and_run():
    """
    Launch a GUI that:
    1) Asks for a start date
    2) Asks for an end date
    3) Asks which stores to process
    4) Launches the closing-report logic for each day in the range + each store
    """
    root = tk.Tk()
    root.title("Select Date Range for Closing Report")

    main_frame = tk.Frame(root)
    main_frame.pack(pady=20, padx=20)

    this_year = datetime.now().year
    year_range = [str(this_year-1), str(this_year)]

    # --- Start Date ---
    start_year_combo, start_month_combo, start_day_combo = create_date_selector(
        main_frame, "Select Start Date:", year_range
    )
    # --- End Date ---
    end_year_combo, end_month_combo, end_day_combo = create_date_selector(
        main_frame, "Select End Date:", year_range
    )

    # --- Store checkboxes ---
    tk.Label(main_frame, text="Select Store(s):", font=("Arial", 12, "bold")).pack(pady=(10,5), anchor='w')
    store_vars = create_store_checkboxes(main_frame)

    def on_ok():
        # 1) Get start/end from combos
        start_date = get_date_from_comboboxes(start_year_combo, start_month_combo, start_day_combo)
        end_date = get_date_from_comboboxes(end_year_combo, end_month_combo, end_day_combo)

        if start_date > end_date:
            messagebox.showerror("Date Error", "Start date cannot be after End date.")
            return

        # 2) Build list of dates from start_date to end_date (inclusive)
        date_list = []
        current = start_date
        while current <= end_date:
            date_list.append(current)
            current += timedelta(days=1)

        # 3) Which stores are selected?
        selected_stores = [
            store_name for store_name, var in store_vars.items() if var.get() == 1
        ]
        if not selected_stores:
            messagebox.showinfo("No Store Selected", "No stores selected. Exiting.")
            root.destroy()
            return

        # 4) Close GUI
        root.destroy()

        # 5) Launch browser, login once
        driver = launchBrowser()
        login(driver)

        # 6) For each selected store:
        for store_name in selected_stores:
            if not select_store(driver, store_name):
                print(f"Skipping store {store_name} due to selection error.")
                continue
            print(f"\n\033[1m--- Processing store {store_name} ---\033[0m")

            # 7) For each date in the range, run your daily logic
            for date_to_run in date_list:
                # short delay between days (optional)
                time.sleep(2)
                process_single_day(driver, date_to_run)

        # 8) Done
        driver.quit()
        print("\nAll processing completed successfully.")

    # "OK" Button
    tk.Button(root, text="OK", command=on_ok, font=("Arial", 12, "bold"), bg="lightblue").pack(pady=10)
    root.mainloop()


# Run if called directly
if __name__ == "__main__":
    open_gui_and_run()
