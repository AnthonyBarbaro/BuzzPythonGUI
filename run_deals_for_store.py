import datetime
import pandas as pd
import os

from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from deals import run_deals_for_store  # Existing function to fetch store data

# --------------------------------------------------------------------------
# SCOPES and SPREADSHEET
# --------------------------------------------------------------------------
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.file"
]
SPREADSHEET_ID = "1dTjTyKmERbXaSw-H7oTP3l2-d-taIEhIK5587p1yjx4"

# --------------------------------------------------------------------------
# AUTHENTICATE
# --------------------------------------------------------------------------
def authenticate_sheets():
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())

    service = build("sheets", "v4", credentials=creds)
    return service

# --------------------------------------------------------------------------
# COMBINED DATA
# --------------------------------------------------------------------------
def get_combined_deal_data():
    """Pulls data for MV & LM and combines them."""
    mv_deals = run_deals_for_store("MV")
    lm_deals = run_deals_for_store("LM")
    combined = []
    for deal in mv_deals + lm_deals:
        # Clean up formatted strings like '$3,690.24'
        owed_str = str(deal.get("kickback", "0")).replace("$", "").replace(",", "")
        try:
            owed = float(owed_str)
        except ValueError:
            print(f"ERROR: Could not convert {owed_str} to float. Defaulting to 0.")
            owed = 0.0

        combined.append({
            "brand": deal["brand"],
            "store": deal["store"],
            "inventory_cost": deal["inventory_cost"],
            "owed": owed,
            "start": deal["start"],
            "end": deal["end"],
            "location": "MV" if deal["store"] == "MV" else "LM"
        })
    return combined

def fetch_combined_deals():
    """Groups deals by brand -> week -> {owed_mv, owed_lm}."""
    data = get_combined_deal_data()
    weekly_data = {}

    for d in data:
        brand = d["brand"]
        start_date = d["start"]
        end_date = d["end"]
        week_key = f"{start_date} - {end_date}"

        if brand not in weekly_data:
            weekly_data[brand] = {}

        if week_key not in weekly_data[brand]:
            weekly_data[brand][week_key] = {"owed_mv": 0.0, "owed_lm": 0.0}

        if d["location"] == "MV":
            weekly_data[brand][week_key]["owed_mv"] += d["owed"]
        else:
            weekly_data[brand][week_key]["owed_lm"] += d["owed"]

    return weekly_data

# --------------------------------------------------------------------------
# UPDATE SHEET
# --------------------------------------------------------------------------
def update_sheet(service, combined_deals):
    """
    1) Checks "Processed Weeks" to skip duplicates.
    2) Adds brand headers, week data, total lines.
    3) Applies advanced formatting via batchUpdate.
    """
    sheet = service.spreadsheets()

    # --1-- Grab existing main sheet data
    main_data_resp = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Sheet1").execute()
    main_data = main_data_resp.get("values", [])

    if main_data:
        df = pd.DataFrame(main_data[1:], columns=main_data[0])
    else:
        df = pd.DataFrame(columns=["Month", "Week", "Brand", "Location",
                                   "Owed", "Invoice Numbers", "Payments", "Remaining"])

    # Convert numeric columns
    for col in ["Owed", "Payments", "Remaining"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0)

    # --2-- Grab existing processed weeks
    processed_tab_resp = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range="Processed Weeks!A1:A").execute()
    processed_data = processed_tab_resp.get("values", [])
    if processed_data and len(processed_data) > 1:
        processed_weeks = [row[0] for row in processed_data[1:] if row]
    else:
        processed_weeks = []

    # --3-- Build new rows
    rows_to_add = []
    def get_month_name(date_str):
        """Parse 'YYYY-MM-DD' -> Month Name (e.g. 'December')."""
        # Handle potential errors or partial dates
        try:
            return datetime.datetime.strptime(date_str, "%Y-%m-%d").strftime("%B")
        except:
            return ""

    new_weeks = []  # Track newly processed weeks

    for brand, weeks_data in combined_deals.items():
        # Brand Header Row (placeholder, we format it bold later)
        rows_to_add.append({
            "Month": "",
            "Week": "",
            "Brand": f"{brand}",
            "Location": "",
            "Owed": "",
            "Invoice Numbers": "",
            "Payments": "",
            "Remaining": "",
        })

        for week_key, amounts in weeks_data.items():
            if week_key in processed_weeks:
                print(f"Week '{week_key}' already processed. Skipping brand: {brand}")
                continue
            # If not in processed weeks, we add it
            new_weeks.append(week_key)

            start_date = week_key.split(" - ")[0]
            month_name = get_month_name(start_date)

            owed_mv = amounts["owed_mv"]
            owed_lm = amounts["owed_lm"]
            total_owed = owed_mv + owed_lm

            # Row for MV
            if owed_mv > 0:
                rows_to_add.append({
                    "Month": month_name,
                    "Week": week_key,
                    "Brand": "",
                    "Location": "MV",
                    "Owed": owed_mv,
                    "Invoice Numbers": "",
                    "Payments": 0,
                    "Remaining": owed_mv,
                })

            # Row for LM
            if owed_lm > 0:
                rows_to_add.append({
                    "Month": "",
                    "Week": "",
                    "Brand": "",
                    "Location": "LM",
                    "Owed": owed_lm,
                    "Invoice Numbers": "",
                    "Payments": 0,
                    "Remaining": owed_lm,
                })

            # Total Row
            rows_to_add.append({
                "Month": "",
                "Week": "",
                "Brand": f"{brand} Total",
                "Location": "",
                "Owed": total_owed,
                "Invoice Numbers": "",
                "Payments": 0,
                "Remaining": total_owed,
            })

        # Spacer row
        rows_to_add.append({
            "Month": "",
            "Week": "",
            "Brand": "",
            "Location": "",
            "Owed": "",
            "Invoice Numbers": "",
            "Payments": "",
            "Remaining": "",
        })

    # --4-- Merge new data into df
    if rows_to_add:
        df_new = pd.DataFrame(rows_to_add)
        df = pd.concat([df, df_new], ignore_index=True)

    # Drop duplicates
    df = df.drop_duplicates()

    # --5-- Convert back to values
    values = [df.columns.tolist()] + df.fillna("").values.tolist()
    body = {"values": values}

    # --6-- Update main sheet data
    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range="Sheet1",
        valueInputOption="RAW",
        body=body
    ).execute()

    print("Data updated in main sheet.")

    # --7-- Update 'Processed Weeks'
    newly_processed = sorted(set(processed_weeks + new_weeks))
    track_processed_weeks(service, newly_processed, sheet)

    # --8-- Apply advanced formatting
    apply_formatting(service, len(values), "Sheet1")
    print("Advanced formatting applied.")


def track_processed_weeks(service, processed_weeks, sheet=None):
    """
    Logs all processed weeks in a separate tab "Processed Weeks".
    Creates the tab if missing, then writes the unique weeks.
    """
    if sheet is None:
        sheet = service.spreadsheets()

    # Check if "Processed Weeks" tab exists
    sheet_metadata = sheet.get(spreadsheetId=SPREADSHEET_ID).execute()
    sheets = sheet_metadata.get("sheets", [])
    processed_tab_exists = any(s.get("properties", {}).get("title") == "Processed Weeks" for s in sheets)
    if not processed_tab_exists:
        # Create
        requests = [{
            "addSheet": {
                "properties": {"title": "Processed Weeks"}
            }
        }]
        sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body={"requests": requests}).execute()
        print("Created 'Processed Weeks' tab.")

    # Prepare logging data
    weeks_log = [[week] for week in processed_weeks]
    body = {"values": [["Processed Weeks"]] + weeks_log}

    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID,
        range="Processed Weeks!A1",
        valueInputOption="RAW",
        body=body
    ).execute()

    print("Processed weeks logged.")


def apply_formatting(service, row_count, sheet_name):
    """
    Uses batchUpdate for advanced formatting:
      - Bold brand headers
      - Bold brand total rows
      - Maybe color them differently
    """
    # We need to iterate over row data to find brand headers or totals
    # However, we only have row_count. Let's do a naive approach:
    # 1) We'll format all rows that contain "**" or "Total" in column C
    #    (But we no longer use "**", so let's detect brand headers or 'Total' text.)
    # 2) We'll color brand headers differently from brand totals.

    # Step: fetch the updated data from the sheet to read row contents
    sheet = service.spreadsheets()
    main_data_resp = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=f"{sheet_name}").execute()
    main_data = main_data_resp.get("values", [])

    # We'll build requests for each row we want to format
    requests = []
    if not main_data:
        return  # Nothing to format

    # Identify column indexes based on header
    headers = main_data[0]
    brand_col = headers.index("Brand") if "Brand" in headers else 2

    # Start from row 2 (index 1 in the data)
    for row_idx in range(1, len(main_data)):
        row_data = main_data[row_idx]
        if len(row_data) <= brand_col:
            continue
        brand_val = row_data[brand_col]

        # Check if it's a brand header or brand total
        if brand_val.endswith(" Total") and brand_val != "":
            # Bold and color for total row
            requests.append(format_row_request(
                sheet_id=0,  # assuming single sheet with ID=0
                start_row=row_idx,
                end_row=row_idx + 1,
                bold=True,
                bg_color=(0.9, 0.9, 0.9)  # light gray
            ))
        elif brand_val != "" and row_data[0] == "":  
            # If Month is empty but Brand is non-empty,
            # we interpret this as a brand header row
            requests.append(format_row_request(
                sheet_id=0,
                start_row=row_idx,
                end_row=row_idx + 1,
                bold=True,
                bg_color=(0.7, 0.85, 1)  # light blue
            ))

    # Send batchUpdate
    if requests:
        body = {"requests": requests}
        sheet.batchUpdate(spreadsheetId=SPREADSHEET_ID, body=body).execute()


def format_row_request(sheet_id, start_row, end_row, bold=False, bg_color=None):
    """
    Helper for building a batchUpdate request to format a row.
    - bg_color: (R, G, B) in [0, 1].
    """
    if bg_color is None:
        bg_color = (1, 1, 1)  # white

    return {
        "repeatCell": {
            "range": {
                "sheetId": sheet_id,
                "startRowIndex": start_row,
                "endRowIndex": end_row,
                "startColumnIndex": 0,
                "endColumnIndex": 8,  # adjust if more columns
            },
            "cell": {
                "userEnteredFormat": {
                    "textFormat": {"bold": bold},
                    "backgroundColor": {
                        "red": bg_color[0],
                        "green": bg_color[1],
                        "blue": bg_color[2]
                    }
                }
            },
            "fields": "userEnteredFormat(textFormat, backgroundColor)"
        }
    }


def main():
    service = authenticate_sheets()
    combined_deals = fetch_combined_deals()
    update_sheet(service, combined_deals)
    print("Done.")


if __name__ == "__main__":
    main()
