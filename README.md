# Sheets Automation Demo

This repository automates a weekly e-commerce/payment reporting workflow that teams often handle with manual spreadsheet copy/paste.

Includes an example CSV export for demonstration.

**Demo video:** (coming soon)

## Who hires this (real scenarios)
- Ops or finance leads who need a reliable weekly order report without spending hours reconciling CSV exports by hand.
- E-commerce managers who want manual copy/paste into Google Sheets replaced with a repeatable script.
- Founders and small teams who need share-ready Excel plus optional Google Sheets tabs for stakeholder updates.
- Teams that need fewer reporting errors (dates, currency parsing, duplicate orders) before numbers go to leadership.
- Clients who value a clean, documented, maintainable workflow with fast communication and reliable support when requirements change.

## Problem Statement
Operations teams need a reliable weekly order report, but manual spreadsheet updates are time-consuming and error-prone. This project demonstrates a lightweight Python pipeline that standardizes order data and generates ready-to-share weekly outputs for stakeholders.

## Typical input data (what you export each week)
Common source exports this workflow is designed for:

- Shopify Orders export (CSV): **Orders → Export CSV**.
- WooCommerce orders export (CSV): filtered exports used for reporting/accounting.
- Stripe Payments export (CSV): **Payments → Export**, then choose date range and columns.
- Amazon Seller Central order reports (CSV): order/report downloads used by operations/accounting.

Expected/common columns (mapped by `src/run.py`):
- `order_id` (or order id / id)
- `order_date` (or order timestamp / order_datetime / date)
- `status`
- `channel` (or source / sales_channel)
- `product`
- `units` (quantity)
- `price` (or unit_price / amount)
- `revenue` (or total / total_revenue)

Typical real-world messiness handled in weekly exports:
- Mixed date formats that can produce unparseable timestamps.
- Currency symbols and formatting in amount fields (for example `$1,234.56`).
- Duplicate order IDs, missing required values, or incomplete rows.

## What `src/run.py` Does
`src/run.py` executes the full reporting flow end-to-end:

1. Loads raw order data.
2. Cleans and standardizes fields (for example, dates and numeric values).
3. Filters records for the reporting period.
4. Runs weekly aggregations to produce summary metrics.
5. Identifies top-performing products.
6. Exports CSV outputs and a formatted Excel report for distribution.

## What you get (deliverables)
- [x] Cleaned dataset: `clean_orders.csv`
- [x] Weekly summary: `weekly_summary.csv`
- [x] Top products: `top_products.csv`
- [x] Share-ready Excel report: `weekly_report.xlsx`
- [x] Optional Google Sheet with 2 tabs: `weekly_summary`, `top_products`

## How it works (in 60 seconds)
1. Export order/payment CSVs from your platforms.
2. Place the file in `data/raw/` (or pass a custom file path with `--input`).
3. Run the script.
4. Get cleaned outputs in `data/processed/`.
5. Optionally publish summary tabs to Google Sheets.

Data-quality behavior includes quarantining dropped rows into `data/processed/quarantine_bad_rows.csv` (for example: unparseable dates, missing required fields, invalid amounts, duplicates, or filtered-out rows).

## Setup
Python and pip are the actual runtime requirements. Recommended: use a virtual environment to avoid dependency conflicts.

1. Create and activate a virtual environment:
   - macOS/Linux:
     ```bash
     python -m venv .venv
     source .venv/bin/activate
     ```
   - Windows (PowerShell):
     ```powershell
     py -m venv .venv
     .\.venv\Scripts\Activate.ps1
     ```
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Run
Execute the pipeline with defaults:

```bash
python src/run.py
```

Common options:

```bash
python src/run.py --verbose
python src/run.py --input data/raw/orders_export.csv --outdir data/processed
python src/run.py --start-date 2024-01-01 --end-date 2024-03-31
python src/run.py --no-paid-only
```

## Windows (PowerShell) Support
Designed to run on Windows and Linux. Tested on Linux; Windows validation will be added via CI next. If you hit an issue, open a ticket and I’ll troubleshoot.

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python .\src\run.py --input .\data\raw\orders_export.csv --outdir .\data\processed
```

Notes:
- On Windows, prefer `py -m venv .venv`.
- Use Windows-style paths (`.\data\raw\...`) in PowerShell; forward slashes usually work, but backslashes are the safest default.

## Scheduling (optional)
- Linux: use `cron` to run `python src/run.py` weekly and write logs to a file.
- Windows: use Task Scheduler to run `python .\src\run.py` on a weekly trigger.

## Preview

Preview table below shows example output format; generated results vary by your exports and date range.

| week_start | week_end   | total_orders | total_revenue | average_order_value |
|------------|------------|--------------|---------------|---------------------|
| 2024-03-04 | 2024-03-10 | 128          | 18432.75      | 144.01              |
| 2024-03-11 | 2024-03-17 | 141          | 20109.20      | 142.62              |
| 2024-03-18 | 2024-03-24 | 136          | 19688.45      | 144.77              |
| 2024-03-25 | 2024-03-31 | 149          | 21452.10      | 143.97              |

## Outputs
After a successful run, the pipeline writes the following files:

- `data/processed/clean_orders.csv`
- `data/processed/weekly_summary.csv`
- `data/processed/top_products.csv`
- `data/processed/weekly_report.xlsx` (tabs: `Summary`, `Top Products`)
- `data/processed/quarantine_bad_rows.csv`

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
