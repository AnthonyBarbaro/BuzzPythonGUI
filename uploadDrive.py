##############################################################################
# 6) GOOGLE DRIVE UPLOADER
##############################################################################
from datetime import date, timedelta, datetime as dt
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

def run_drive_upload():
    """
    Upload brand_reports/*.xlsx + any done/**/*Hashish_*.xlsx to 
    Google Drive folder "2025_Kickback -> <week range>", 
    writing all links into links.txt
    """
    print("\n===== Running googleDriveUploader logic... =====\n")
    import google.auth
    import os
    import google.auth.transport.requests
    from google_auth_oauthlib.flow import InstalledAppFlow
    from googleapiclient.discovery import build
    from googleapiclient.http import MediaFileUpload
    from google.oauth2.credentials import Credentials

    SCOPES = ['https://www.googleapis.com/auth/drive.file']
    LINKS_FILE = "links.txt"
    PARENT_FOLDER_NAME = "2026_Kickback"
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
        return f"{lm.strftime('%m-%d')} to {ls.strftime('%m-%d')}"
        

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
def main():
    run_drive_upload()

if __name__ == "__main__":
    main()
