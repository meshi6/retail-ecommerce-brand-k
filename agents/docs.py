"""
CLI helper for appending competitor reports to the KIABI Competitor Reports sheet.

Commands:
  append    Read report text from stdin, append as a dated row in the Reports tab
  setup     Create the Reports tab and headers (run once)
"""
import json
import sys
from datetime import datetime
from pathlib import Path

from google.oauth2 import service_account
from googleapiclient.discovery import build

BASE_DIR = Path(__file__).parent
SA_FILE = BASE_DIR / "service-account.json"
CONFIG_FILE = BASE_DIR / "config.json"
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]
TAB_NAME = "Reports"


def _sheets():
    creds = service_account.Credentials.from_service_account_file(str(SA_FILE), scopes=SCOPES)
    return build("sheets", "v4", credentials=creds)


def _config():
    with open(CONFIG_FILE) as f:
        return json.load(f)


def _ensure_reports_tab(svc, sheet_id):
    meta = svc.spreadsheets().get(spreadsheetId=sheet_id).execute()
    existing = [s["properties"]["title"] for s in meta["sheets"]]
    if TAB_NAME not in existing:
        svc.spreadsheets().batchUpdate(
            spreadsheetId=sheet_id,
            body={"requests": [{"addSheet": {"properties": {"title": TAB_NAME}}}]},
        ).execute()
        svc.spreadsheets().values().update(
            spreadsheetId=sheet_id,
            range=f"{TAB_NAME}!A1:C1",
            valueInputOption="RAW",
            body={"values": [["Date", "Period", "Report"]]},
        ).execute()
        print(f"✅ '{TAB_NAME}' tab created")
    return TAB_NAME in existing


def cmd_setup():
    cfg = _config()
    svc = _sheets()
    already_existed = _ensure_reports_tab(svc, cfg["doc_id"])
    if already_existed:
        print(f"✅ '{TAB_NAME}' tab already exists")
    print("Setup complete.")


def cmd_append():
    report_text = sys.stdin.read().strip()
    cfg = _config()
    svc = _sheets()
    _ensure_reports_tab(svc, cfg["doc_id"])

    date_str = datetime.now().strftime("%Y-%m-%d")
    now = datetime.now()
    period = f"{date_str}"

    svc.spreadsheets().values().append(
        spreadsheetId=cfg["doc_id"],
        range=f"{TAB_NAME}!A:C",
        valueInputOption="USER_ENTERED",
        body={"values": [[date_str, period, report_text]]},
    ).execute()

    print(json.dumps({
        "sheet_id": cfg["doc_id"],
        "date": date_str,
        "url": f"https://docs.google.com/spreadsheets/d/{cfg['doc_id']}",
    }))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "append":
        cmd_append()
    elif cmd == "setup":
        cmd_setup()
    else:
        print(f"Unknown command: {cmd}", file=sys.stderr)
        sys.exit(1)
