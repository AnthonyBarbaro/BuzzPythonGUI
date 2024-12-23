import os
import re
import time
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from login import username, password

store_abbr_map = {
    "Buzz Cannabis - Mission Valley": "MV",
    "Buzz Cannabis-La Mesa": "LM"
}

def wait_for_new_file(download_directory, before_files, timeout=30):
    """
    Waits for a new file to appear in the download directory within the given timeout.
    Returns the new filename if found, otherwise None.
    """
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

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)
    driver.get("https://dusk.backoffice.dutchie.com/products/catalog")
    return driver

def login():
    wait = WebDriverWait(driver, 10)
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='auth_input_username']"))).send_keys(username)
    wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "input[data-testid='auth_input_password']"))).send_keys(password)
    login_button = wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button[data-testid='auth_button_go-green']")))
    login_button.click()

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

def select_dropdown_item(item_text):
    wait = WebDriverWait(driver, 10)
    try:
        click_dropdown()
        # Use data-testid attribute
        xpath = f"//li[@data-testid='rebrand-header_menu-item_{item_text}']"
        item = wait.until(EC.element_to_be_clickable((By.XPATH, xpath)))
        item.click()
        print(f"Selected store: {item_text}")
        return True
    except (TimeoutException, NoSuchElementException) as e:
        print(f"Error while trying to select '{item_text}' from the dropdown: {e}")
        return False

def clickActionsAndExport(current_store):
    try:
        time.sleep(12)  # Wait for the page to fully load
        wait = WebDriverWait(driver, 10)
        
        # Get the current file list before clicking export
        files_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "files")
        before_files = set(os.listdir(files_dir))

        actions_button = wait.until(EC.element_to_be_clickable((By.ID, 'actions-menu-button')))
        actions_button.click()
        print("Actions button clicked successfully.")
        
        export_option = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "li[data-testid='catalog-list-actions-menu-item-export']")))
        export_option.click()
        print("Export option clicked successfully.")
        
        export_csv_button = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, "body > div.sc-jYnRlT.kGxwGQ.sc-heKhxA.Bgfyt.MuiDialog-root.sc-bBPnyn.hrApOB.MuiModal-root > div.sc-fhrEpP.dWiAWv.MuiDialog-container.MuiDialog-scrollPaper > div > div.sc-iyxVF.bdkceX.MuiDialogActions-root.MuiDialogActions-spacing.sc-fopvND.finsng > div.primary-actions > button:nth-child(1)")))
        export_csv_button.click()
        print("Export CSV button clicked successfully.")

        # Wait for the new file to appear
        new_file = wait_for_new_file(files_dir, before_files, timeout=60)
        if new_file:
            print(f"New file downloaded: {new_file}")
            # Rename the file with date and store abbreviation
            store_abbr = store_abbr_map.get(current_store, "UNK")
            today_str = datetime.now().strftime("%m-%d-%Y")

            original_path = os.path.join(files_dir, new_file)
            # Assume original file is a CSV, adjust as needed
            extension = os.path.splitext(new_file)[1]
            new_filename = f"{today_str}_{store_abbr}{extension}"
            new_path = os.path.join(files_dir, new_filename)
            os.rename(original_path, new_path)
            print(f"Renamed {new_file} to {new_filename}")
        else:
            print("No new file detected after export.")
        
    except TimeoutException:
        print("An element could not be found or clicked within the timeout period.")
    except Exception as e:
        print(f"An error occurred: {e}")

# Main execution
driver = launchBrowser()
login()

store_names = ["Buzz Cannabis - Mission Valley", "Buzz Cannabis-La Mesa"]
for store in store_names:
    if not select_dropdown_item(store):
        break
    clickActionsAndExport(store)

driver.quit()