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

## Next Steps
As an optional extension, you can integrate Google Sheets API publishing to push the generated report outputs directly into a shared stakeholder workbook.
