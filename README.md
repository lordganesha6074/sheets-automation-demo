# Sheets Automation Demo

This demo shows how to automate a weekly order reporting workflow that is often maintained manually in spreadsheets.

## Problem Statement
Operations teams need a reliable weekly order report, but manual spreadsheet updates are time-consuming and error-prone. This project demonstrates a lightweight Python pipeline that standardizes order data and generates ready-to-share weekly outputs for stakeholders.

## What `src/run.py` Does
`src/run.py` executes the full reporting flow end-to-end:

1. Loads raw order data.
2. Cleans and standardizes fields (for example, dates and numeric values).
3. Filters records for the reporting period.
4. Runs weekly aggregations to produce summary metrics.
5. Identifies top-performing products.
6. Exports CSV outputs and a formatted Excel report for distribution.

## Setup
1. Create and activate a virtual environment (optional but recommended):
   - macOS/Linux:
     ```bash
     python -m venv .venv
     source .venv/bin/activate
     ```
   - Windows (PowerShell):
     ```powershell
     python -m venv .venv
     .venv\Scripts\Activate.ps1
     ```
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Run
Execute the pipeline:

```bash
python src/run.py
```

## Outputs
After a successful run, the pipeline writes the following files:

- `data/processed/clean_orders.csv`
- `data/processed/weekly_summary.csv`
- `data/processed/top_products.csv`
- `data/processed/weekly_report.xlsx` (tabs: `Summary`, `Top Products`)

## Publish to Google Sheets (optional)

You can optionally publish `weekly_summary.csv` and `top_products.csv` into a Google Sheet as separate tabs (`weekly_summary`, `top_products`).

1. Create a Google Cloud service account and enable Google Sheets API.
2. Download the service account JSON key to a local path (for example: `/secure/path/service-account.json`).
3. Share the destination Google Sheet with the service account email, or allow the script to create a new sheet.
4. Set environment variables:
   - Required: `GOOGLE_APPLICATION_CREDENTIALS=/secure/path/service-account.json`
   - Optional: `GOOGLE_SHEET_ID=<existing_sheet_id>` (if omitted, a new sheet is created)
5. Run publishing:

```bash
GOOGLE_APPLICATION_CREDENTIALS=/secure/path/service-account.json python src/run.py --publish
```

The command prints the Google Sheet ID and URL when publish succeeds.

**Security note:** Never commit service account credential JSON files or secrets to source control. Keep credentials outside the repository.

## Next Steps
As an optional extension, you can integrate Google Sheets API publishing to push the generated report outputs directly into a shared stakeholder workbook.
