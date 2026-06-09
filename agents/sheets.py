"""
CLI helper for reading/writing the KIABI Price Tracker Google Sheet.

Commands:
  list              Print all tracked products as JSON (with row numbers)
  add NAME COMP URL Add a new product row
  update            Read JSON update payload from stdin, write new prices
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
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
]
SHEET_NAME = "Price Tracker"


def _sheets():
    creds = service_account.Credentials.from_service_account_file(str(SA_FILE), scopes=SCOPES)
    return build("sheets", "v4", credentials=creds)


def _config():
    with open(CONFIG_FILE) as f:
        return json.load(f)


def cmd_list():
    cfg = _config()
    svc = _sheets()
    result = svc.spreadsheets().values().get(
        spreadsheetId=cfg["sheet_id"],
        range=f"{SHEET_NAME}!A2:G",
    ).execute()
    rows = result.get("values", [])
    out = []
    for i, row in enumerate(rows):
        while len(row) < 7:
            row.append("")
        out.append({
            "row": i + 2,
            "name": row[0],
            "competitor": row[1],
            "url": row[2],
            "current_price": row[3],
            "last_price": row[4],
            "last_checked": row[5],
            "delta_pct": row[6],
        })
    print(json.dumps(out, ensure_ascii=False, indent=2))


def cmd_add(name, competitor, url):
    cfg = _config()
    svc = _sheets()
    svc.spreadsheets().values().append(
        spreadsheetId=cfg["sheet_id"],
        range=f"{SHEET_NAME}!A:G",
        valueInputOption="USER_ENTERED",
        body={"values": [[name, competitor, url, "", "", "", ""]]},
    ).execute()

    result = svc.spreadsheets().values().get(
        spreadsheetId=cfg["sheet_id"],
        range=f"{SHEET_NAME}!A2:A",
    ).execute()
    count = len(result.get("values", []))
    print(f"Added: {name} ({competitor}) — {count} product(s) tracked")


def cmd_update():
    payload = json.loads(sys.stdin.read())
    cfg = _config()
    svc = _sheets()
    now = datetime.now().strftime("%Y-%m-%d %H:%M")

    for u in payload:
        row = u["row"]
        result = svc.spreadsheets().values().get(
            spreadsheetId=cfg["sheet_id"],
            range=f"{SHEET_NAME}!D{row}",
        ).execute()
        prev_price = (result.get("values") or [[""]])[0][0]

        svc.spreadsheets().values().update(
            spreadsheetId=cfg["sheet_id"],
            range=f"{SHEET_NAME}!D{row}:G{row}",
            valueInputOption="USER_ENTERED",
            body={"values": [[u["current_price"], prev_price, now, u.get("delta_pct", "")]]},
        ).execute()

    print(json.dumps({"updated": len(payload), "timestamp": now}))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "list":
        cmd_list()
    elif cmd == "add":
        cmd_add(sys.argv[2], sys.argv[3], sys.argv[4])
    elif cmd == "update":
        cmd_update()
    else:
        print(f"Unknown command: {cmd}", file=sys.stderr)
        sys.exit(1)
