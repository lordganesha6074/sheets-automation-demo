# Sheets Automation Demo

Turn a messy weekly export into a clean, share-ready reporting pack in one command.
This project shows a repeatable `export → run → review → share` workflow using Python.

## Before → After
Manual spreadsheet cleanup, filters, pivots, and copy/paste become an automated clean dataset, weekly summary, top products, Excel report, and optional Google Sheets publish.

## Impact (ROI)
If reporting takes **H hours/week**, this pipeline turns it into a one-command run plus review.

Simple calculator:

- Weekly time saved = `H - R`
- Weekly value saved = `(H - R) × hourly_rate`
- Annual value saved = `weekly_value_saved × 52`

Where:
- `H` = current manual hours/week
- `R` = run + review hours/week after automation
- `hourly_rate` = loaded internal cost per hour

Example (illustrative only):
- If `H = 2.0`, `R = 0.25`, and `hourly_rate = £25`:
- Weekly value = `(2.0 - 0.25) × £25 = £43.75`
- Annual value ≈ `£2,275`

## Who this is for
Practical fit for teams that need dependable reporting from recurring data exports, including:

- E-commerce operators exporting from Shopify, WooCommerce, Amazon Seller Central, Stripe, or similar platforms.
- Ops/finance teams producing weekly reporting packs and KPI rollups.
- Agencies consolidating multiple client exports into consistent deliverables.
- Any business exporting CSV/Excel files and needing repeatable reporting with less manual cleanup.

## What you get (Deliverables)
- Cleaned dataset: `data/processed/clean_orders.csv`
- Weekly summary: `data/processed/weekly_summary.csv`
- Top products table: `data/processed/top_products.csv`
- Share-ready Excel workbook: `data/processed/weekly_report.xlsx` (tabs: `Summary`, `Top Products`)
- Quarantine file for excluded rows: `data/processed/quarantine_bad_rows.csv`
- Run-level quality counters: `data/processed/data_quality_report.json`
- Optional Google Sheet publish (tabs: `weekly_summary`, `top_products`)

## Example metrics generated
Preview metric columns in reporting outputs include:
- `orders`
- `units`
- `revenue`
- `aov`
- `revenue_wow_pct`
- `channel_revenue_share_pct`

## How it works (60 seconds)
1. Export CSV from your source system.
2. Put it in `data/raw/` (or pass `--input`).
3. Run `python src/run.py`.
4. Review outputs in `data/processed/`.
5. (Optional) Publish summary outputs to Google Sheets with `--publish`.

## Reporting assumptions (Scope clarity)
Documented from `src/run.py` behavior.

- **Revenue field logic (trusted source with fallback):**
  - If a revenue column exists, `_revenue` trusts that source value when present.
  - If revenue is missing and a price column is available, fallback is `units × price`.
  - If no revenue column exists, revenue is computed from `units × price`.
  - Before aggregation, rows can be quarantined with `drop_reason=invalid_amount` when amount inputs are unusable for the active price/revenue-column setup.
  - This means unusable amount inputs are not always silently zero-filled; affected rows may be excluded before `_revenue` is summed.

- **Paid-only filter default:**
  - Default is `--paid-only` (enabled).
  - Status is normalized to lowercase/trimmed and only exact `paid` rows are kept.
  - Non-paid statuses (for example `cancelled`, `refunded`, `unpaid`) are quarantined with `drop_reason=status_not_paid`.
  - Use `--no-paid-only` to include all statuses.

- **Date used for reporting:**
  - Aggregation uses `order_date` (mapped from aliases like `order timestamp`, `order_datetime`, `date`).
  - Payment date is not used for time bucketing in this workflow.

- **Week definition:**
  - Week key is built with `to_period("W-MON").start_time`.
  - This means buckets are **Tuesday through Monday**, labeled by the **week start date** (Tuesday) in `YYYY-MM-DD`.

- **Dedup rule (deterministic):**
  - Dedup key: `order_id`.
  - Rows are stable-sorted by `order_date` ascending, then `_source_row` ascending.
  - Keep `last` per `order_id` (latest date wins; if same date, later source row wins).
  - Dropped duplicates are quarantined with `drop_reason=duplicate_order_id`.

## Controls & auditability
- Raw input file is read as-is; pipeline outputs are written to `data/processed/`.
- Run counters are logged and saved (`raw_rows`, dropped/filtered counts, dedup count, final rows).
- Every excluded row can be inspected in `quarantine_bad_rows.csv` with an explicit `drop_reason`.
- `data_quality_report.json` includes count metrics and run timestamp.
- Dedup is deterministic (same input → same dedup outcome).
- `--verbose` enables debug logging for reconciliation and troubleshooting.

## Quickstart
Fast path (recommended):
- macOS/Linux:
  ```bash
  ./scripts/run_demo.sh
  ```
- Windows (PowerShell):
  ```powershell
  powershell -ExecutionPolicy Bypass -File .\scripts\run_demo.ps1
  ```

Manual path (for transparency):
1. (Optional) Generate example input:
   - Includes an example export generated with realistic edge cases — replace with your real export.
   - macOS/Linux:
     ```bash
     python scripts/generate_orders_export.py
     ```
   - Windows (PowerShell):
     ```powershell
     py .\scripts\generate_orders_export.py
     ```
2. Run the pipeline:
   ```bash
   python src/run.py
   ```
3. Review outputs in `data/processed/`.

## Setup
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
Default run:

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

## Windows (PowerShell)
Supported path:

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python .\src\run.py --input .\data\raw\orders_export.csv --outdir .\data\processed
```

## Preview
Preview format (values vary by input/date range). Context columns remain `week` and `channel`; KPI columns are exactly `orders, units, revenue, aov, revenue_wow_pct, channel_revenue_share_pct`:

| week       | channel      | orders | units | revenue | aov    | revenue_wow_pct | channel_revenue_share_pct |
|------------|--------------|--------|-------|---------|--------|-----------------|---------------------------|
| 2024-12-31 | affiliate    | 9      | 16    | 1501.96 | 166.88 | 2.40            | 10.90                     |
| 2024-12-31 | marketplace  | 24     | 48    | 4296.33 | 179.01 | 5.10            | 31.30                     |
| 2024-12-31 | retail       | 15     | 27    | 2584.74 | 172.32 | -1.20           | 18.80                     |
| 2024-12-31 | web          | 31     | 59    | 5338.12 | 172.20 | 3.70            | 39.00                     |

Rounding conventions for previews/reports (matching `weekly_summary.csv` export formatting):
- Currency columns are rounded to 2 decimals.
- Percent KPI columns (`revenue_wow_pct`, `channel_revenue_share_pct`) are rounded to 2 decimals.

## Outputs
After a successful run, expected files:
- `data/processed/clean_orders.csv`
- `data/processed/weekly_summary.csv`
- `data/processed/top_products.csv`
- `data/processed/weekly_report.xlsx`
- `data/processed/quarantine_bad_rows.csv`
- `data/processed/data_quality_report.json`

Reference run metrics (using the included export generator):
- Reference run (included export generator): 210 rows in → 12 dropped (bad dates) → 161 paid → 157 after dedup
- Outputs generated in one run: weekly_summary + top_products + Excel report + quarantine + data quality report

## Publish to Google Sheets (optional)
Publish `weekly_summary.csv` and `top_products.csv` to two tabs (`weekly_summary`, `top_products`).

1. Create a Google Cloud service account and enable the Google Sheets API.
2. Download the service-account JSON key to a secure local path.
3. Share your destination sheet with the service-account email (or let script create a sheet).
4. Set environment variables:
   - Required: `GOOGLE_APPLICATION_CREDENTIALS=/secure/path/service-account.json`
   - Optional: `GOOGLE_SHEET_ID=<existing_sheet_id>`
5. Run:

```bash
GOOGLE_APPLICATION_CREDENTIALS=/secure/path/service-account.json python src/run.py --publish
```

## Scheduling (optional)
- Linux: cron job running `python src/run.py` (with logs redirected).
- Windows: Task Scheduler job running `python .\src\run.py` on a weekly trigger.

## Technical notes
- Python version in CI: **3.11** (`.github/workflows/ci.yml`)
- Core libraries: `pandas`, `python-dateutil`, `openpyxl`
- Optional publish libraries: `gspread`, `google-auth`
- Tests exist in `tests/` (pytest) and CI exists via GitHub Actions (`.github/workflows/ci.yml`)

Expected/common mapped fields:
- `order_id` (or `order id` / `id`)
- `order_date` (or `order timestamp` / `order_datetime` / `date`)
- `status`
- `channel` (or `source` / `sales_channel`)
- `product`
- `units` (or `quantity` / `qty`)
- `price` (or `unit_price` / `amount`)
- `revenue` (or `total` / `total_revenue`)
