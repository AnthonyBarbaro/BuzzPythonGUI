#!/usr/bin/env python3
import time
import sys
import platform
import smtplib
from email.mime.text import MIMEText
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException
import os
from dotenv import load_dotenv

# ───────────────────── CONFIG ─────────────────────
CUSTOMER_NAME = "SAM AMIR EMADVAEZ"
#CUSTOMER_NAME = "ANTHONY BARBARO"
CHECK_INTERVAL = 4 # seconds between checks

# Email/SMS setup
load_dotenv()
SENDER_EMAIL = os.getenv("SENDER_EMAIL")
APP_PASSWORD = os.getenv("APP_PASSWORD")
RECIPIENTS = [r.strip() for r in os.getenv("RECIPIENT", "").split(",") if r.strip()]

# ───────────────────── ALERT SOUND ─────────────────────
def play_sound():
    if platform.system() == "Windows":
        import winsound
        winsound.Beep(1000, 2000)
    else:
        sys.stdout.write("\a")
        sys.stdout.flush()

# ───────────────────── SEND SMS ─────────────────────
def send_sms_notification(message: str):
    msg = MIMEText(message)
    msg["From"] = SENDER_EMAIL
    msg["To"] = ", ".join(RECIPIENTS)
    msg["Subject"] = "Check-in Alert"

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(SENDER_EMAIL, APP_PASSWORD)
        server.sendmail(SENDER_EMAIL, RECIPIENTS, msg.as_string())

# ───────────────────── SELENIUM SETUP ─────────────────────
options = Options()
options.add_argument("--start-maximized")

driver = webdriver.Chrome(service=Service(), options=options)
driver.get("https://dusk.pos.dutchie.com/")

input(">>> Press ENTER here once you’re logged in... ")

print(f"Listening for {CUSTOMER_NAME}... (Ctrl+C to stop)")
seen = False

# ───────────────────── MAIN LOOP ─────────────────────
try:
    while True:
        names = []
        try:
            elements = driver.find_elements(By.CSS_SELECTOR, "[data-testid='order-card_customer-name_p']")
            for el in elements:
                try:
                    names.append(el.text.strip().upper())
                except Exception:
                    # element went stale before we could read it
                    continue

            if CUSTOMER_NAME in names and not seen:
                print(f"⚡ {CUSTOMER_NAME} just checked in!")
                play_sound()
                send_sms_notification(f"{CUSTOMER_NAME} just checked in!")
                seen = True

            if CUSTOMER_NAME not in names:
                seen = False

        except NoSuchElementException:
            pass

        time.sleep(CHECK_INTERVAL)
except KeyboardInterrupt:
    print("\nStopped listening.")
    driver.quit()
