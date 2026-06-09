"""
CLI helper for appending reports to the KIABI Competitor Reports Google Doc.

Commands:
  append    Read report text from stdin, append as a dated section
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
    "https://www.googleapis.com/auth/documents",
]


def _docs():
    creds = service_account.Credentials.from_service_account_file(str(SA_FILE), scopes=SCOPES)
    return build("docs", "v1", credentials=creds)


def _config():
    with open(CONFIG_FILE) as f:
        return json.load(f)


def cmd_append():
    report_text = sys.stdin.read().strip()
    cfg = _config()
    svc = _docs()
    doc_id = cfg["doc_id"]

    doc = svc.documents().get(documentId=doc_id).execute()
    end_index = doc["body"]["content"][-1]["endIndex"] - 1

    date_str = datetime.now().strftime("%Y-%m-%d")
    divider = f"\n\n{'─' * 60}\n"
    header = f"Competitor Report — {date_str}\n{'─' * 60}\n\n"
    full_text = divider + header + report_text + "\n"

    svc.documents().batchUpdate(
        documentId=doc_id,
        body={"requests": [{"insertText": {
            "location": {"index": end_index},
            "text": full_text,
        }}]},
    ).execute()

    print(json.dumps({
        "doc_id": doc_id,
        "date": date_str,
        "url": f"https://docs.google.com/document/d/{doc_id}",
    }))


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else ""
    if cmd == "append":
        cmd_append()
    else:
        print(f"Unknown command: {cmd}", file=sys.stderr)
        sys.exit(1)
