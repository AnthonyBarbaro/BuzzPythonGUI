import os
import datetime
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# Google Drive API Scopes
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# Folder Structure
REPORTS_FOLDER = "brand_reports"
PARENT_FOLDER_NAME = "2025_Kickback"
LINKS_FILE = "links.txt"

# Authenticate and build the Google Drive service
def authenticate_drive_api():
    creds = None
    token_file = "token.json"

    if os.path.exists(token_file):
        creds = Credentials.from_authorized_user_file(token_file, SCOPES)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open(token_file, "w") as token:
            token.write(creds.to_json())
    
    return build("drive", "v3", credentials=creds)

# Calculate the previous week range (Monday to Sunday)
def get_previous_week_range():
    today = datetime.date.today()
    start_of_current_week = today - datetime.timedelta(days=today.weekday())
    end_of_previous_week = start_of_current_week - datetime.timedelta(days=1)
    start_of_previous_week = end_of_previous_week - datetime.timedelta(days=6)
    
    start_str = start_of_previous_week.strftime("%b %d")
    end_str = end_of_previous_week.strftime("%b %d")
    return f"{start_str} to {end_str}"

# Find or create a folder on Google Drive
def find_or_create_folder(service, folder_name, parent_id=None):
    query = f"mimeType='application/vnd.google-apps.folder' and name='{folder_name}'"
    if parent_id:
        query += f" and '{parent_id}' in parents"
    response = service.files().list(q=query, spaces='drive', fields='files(id, name)').execute()
    folders = response.get('files', [])

    if folders:
        folder_id = folders[0]['id']
        print(f"Folder '{folder_name}' found with ID: {folder_id}")
        return folder_id
    else:
        folder_metadata = {
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
            "parents": [parent_id] if parent_id else []  # Root directory if parent_id is None
        }
        folder = service.files().create(body=folder_metadata, fields="id").execute()
        print(f"Created folder '{folder_name}' with ID: {folder.get('id')}")
        return folder.get("id")

# Upload a file to a specific folder
def upload_file_to_drive(service, file_path, folder_id):
    file_name = os.path.basename(file_path)
    file_metadata = {
        "name": file_name,
        "parents": [folder_id]
    }
    media = MediaFileUpload(file_path, resumable=True)
    file = service.files().create(body=file_metadata, media_body=media, fields="id").execute()
    print(f"Uploaded {file_name} to folder ID {folder_id}. File ID: {file['id']}")
    return file.get("id")

# Set file permissions to "Anyone with the link"
def make_file_public(service, file_id):
    try:
        permission = {
            "type": "anyone",
            "role": "reader"
        }
        service.permissions().create(fileId=file_id, body=permission).execute()
        file = service.files().get(fileId=file_id, fields="webViewLink").execute()
        return file.get("webViewLink")
    except Exception as e:
        print(f"Failed to make file public. File ID: {file_id}. Error: {e}")
        return None

def main():
    # Authenticate Google Drive API
    service = authenticate_drive_api()

    # Find or create the main folder "2025_Kickback" in the root directory
    parent_folder_id = find_or_create_folder(service, PARENT_FOLDER_NAME, parent_id=None)
    print(f"Parent folder ID for '{PARENT_FOLDER_NAME}': {parent_folder_id}")

    # Get the previous week range and create/find a weekly folder
    week_folder_name = get_previous_week_range()
    week_folder_id = find_or_create_folder(service, week_folder_name, parent_id=parent_folder_id)
    print(f"Weekly folder ID for '{week_folder_name}': {week_folder_id}")

    # Open the links file for writing
    with open(LINKS_FILE, "w") as links_file:
        # Iterate through files in the local folder and upload
        if os.path.exists(REPORTS_FOLDER):
            for file_name in os.listdir(REPORTS_FOLDER):
                file_path = os.path.join(REPORTS_FOLDER, file_name)
                if os.path.isfile(file_path) and file_name.endswith(".xlsx"):
                    try:
                        # Extract the first word from the file name for the subfolder
                        first_word = file_name.split()[0]

                        # Create a subfolder using the first word
                        file_folder_id = find_or_create_folder(service, first_word, parent_id=week_folder_id)
                        print(f"Created subfolder for file: {first_word}")

                        # Upload the file to its subfolder
                        file_id = upload_file_to_drive(service, file_path, file_folder_id)

                        # Make the file public and get the shareable link
                        public_link = make_file_public(service, file_id)
                        if public_link:
                            print(f"Public link for {file_name}: {public_link}")

                            # Write the link to the links.txt file
                            links_file.write(f"{file_name}: {public_link}\n")
                    except Exception as e:
                        print(f"Failed to upload {file_name}: {e}")
        else:
            print(f"The folder '{REPORTS_FOLDER}' does not exist!")

if __name__ == "__main__":
    main()
