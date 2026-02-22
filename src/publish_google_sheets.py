from __future__ import annotations

import csv
import os
from pathlib import Path

import gspread
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]


def _load_client() -> gspread.Client:
    creds_path = os.getenv("GOOGLE_APPLICATION_CREDENTIALS")
    if not creds_path:
        raise ValueError(
            "GOOGLE_APPLICATION_CREDENTIALS is required to publish to Google Sheets."
        )

    credentials = Credentials.from_service_account_file(creds_path, scopes=SCOPES)
    return gspread.authorize(credentials)


def _read_csv_rows(path: Path) -> list[list[str]]:
    with path.open("r", encoding="utf-8", newline="") as handle:
        return [row for row in csv.reader(handle)]


def _write_csv_to_worksheet(
    spreadsheet: gspread.Spreadsheet, worksheet_name: str, csv_path: Path
) -> None:
    rows = _read_csv_rows(csv_path)
    try:
        ws = spreadsheet.worksheet(worksheet_name)
        ws.clear()
    except gspread.WorksheetNotFound:
        ws = spreadsheet.add_worksheet(
            title=worksheet_name,
            rows=str(max(len(rows), 1)),
            cols=str(max((len(r) for r in rows), default=1)),
        )

    if rows:
        ws.update("A1", rows)


def publish_csvs(sheet_id: str | None, weekly_csv: Path, top_csv: Path) -> str:
    client = _load_client()

    if sheet_id:
        spreadsheet = client.open_by_key(sheet_id)
    else:
        spreadsheet = client.create("Sheets Automation Demo Output")
        sheet_id = spreadsheet.id

    _write_csv_to_worksheet(spreadsheet, "weekly_summary", weekly_csv)
    _write_csv_to_worksheet(spreadsheet, "top_products", top_csv)

    return sheet_id
