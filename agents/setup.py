"""
Run once to create the Price Tracker sheet and Competitor Reports doc in Drive.
Requires the Drive folder to be shared with the service account first.
"""
import json
from pathlib import Path
from google.oauth2 import service_account
from googleapiclient.discovery import build

BASE_DIR = Path(__file__).parent
SA_FILE = BASE_DIR / "service-account.json"
CONFIG_FILE = BASE_DIR / "config.json"
FOLDER_ID = "1HAYPV3CpIjf1SSKy038wPZ6eiC7bR2xv"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/documents",
]

COMPETITORS = [
    {"name": "Zara", "base_url": "https://www.zara.com/ar/"},
    {"name": "Indian", "base_url": "https://www.indian.com.ar/"},
    {"name": "Renner", "base_url": "https://www.renner.com.ar/"},
    {"name": "Le Utthe", "base_url": "https://www.leutthe.com.ar/"},
    {"name": "Cuesta Blanca", "base_url": "https://www.cuestablanca.com.ar/"},
    {"name": "Cheeky", "base_url": "https://www.cheeky.com.ar/"},
]

SHEET_HEADERS = [["Product Name", "Competitor", "URL", "Current Price", "Last Price", "Last Checked", "Delta %"]]


def create_sheet(drive, sheets):
    file_meta = {
        "name": "KIABI Price Tracker",
        "mimeType": "application/vnd.google-apps.spreadsheet",
        "parents": [FOLDER_ID],
    }
    f = drive.files().create(body=file_meta, fields="id").execute()
    sheet_id = f["id"]

    sheets.spreadsheets().batchUpdate(
        spreadsheetId=sheet_id,
        body={"requests": [{"updateSheetProperties": {
            "properties": {"sheetId": 0, "title": "Price Tracker"},
            "fields": "title"
        }}]}
    ).execute()

    sheets.spreadsheets().values().update(
        spreadsheetId=sheet_id,
        range="Price Tracker!A1:G1",
        valueInputOption="RAW",
        body={"values": SHEET_HEADERS},
    ).execute()

    return sheet_id


def create_doc(drive, docs):
    file_meta = {
        "name": "KIABI Competitor Reports",
        "mimeType": "application/vnd.google-apps.document",
        "parents": [FOLDER_ID],
    }
    f = drive.files().create(body=file_meta, fields="id").execute()
    doc_id = f["id"]

    docs.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": [{"insertText": {
            "location": {"index": 1},
            "text": "KIABI Argentina — Competitor Intelligence Reports\n\n"
        }}]}
    ).execute()

    return doc_id


def main():
    creds = service_account.Credentials.from_service_account_file(str(SA_FILE), scopes=SCOPES)
    drive = build("drive", "v3", credentials=creds)
    sheets = build("sheets", "v4", credentials=creds)
    docs = build("docs", "v1", credentials=creds)

    print("Creating Price Tracker sheet...")
    sheet_id = create_sheet(drive, sheets)
    print(f"  ✅ Sheet ID: {sheet_id}")

    print("Creating Competitor Reports doc...")
    doc_id = create_doc(drive, docs)
    print(f"  ✅ Doc ID: {doc_id}")

    config = {
        "sheet_id": sheet_id,
        "doc_id": doc_id,
        "drive_folder_id": FOLDER_ID,
        "alert_email": "gtm0zero@gmail.com",
        "competitors": COMPETITORS,
    }

    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)

    print(f"\nSheet: https://docs.google.com/spreadsheets/d/{sheet_id}")
    print(f"Doc:   https://docs.google.com/document/d/{doc_id}")
    print("\n✅ Setup complete. Open the sheet → Tools → Notification settings to enable email alerts.")


if __name__ == "__main__":
    main()
