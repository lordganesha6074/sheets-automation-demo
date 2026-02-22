# Video #1 Script (60–90s)

## Hook (say this in first 5–10s)
"Still cleaning CSV exports by hand every week? This turns that into a one-command reporting pipeline."

## 0–15s — Problem + setup
**Say:**
"Ops and finance teams lose hours every week fixing messy exports before reporting. In this demo, I’ll run one script that cleans order data, builds weekly summaries, and creates a share-ready report."

**Show on screen:**
- Repo root in terminal.
- Quick glance at `data/raw/orders_export.csv` in file explorer.
- `README.md` section mentioning outputs.

## 15–45s — Run the pipeline
**Say:**
"I’ll run `python src/run.py` on the sample export. The script standardizes dates and amounts, removes bad rows to quarantine, and generates all reporting outputs in one pass."

**Show on screen:**
- Terminal command: `python src/run.py`
- Console logs finishing successfully.
- Optional quick scroll of output lines that reference processing + file writes.

## 45–90s — Proof + publish option + CTA
**Say:**
"Now in `data/processed/` we have `clean_orders.csv`, `weekly_summary.csv`, `top_products.csv`, plus `weekly_report.xlsx` for sharing. If needed, I can also publish the summary tabs to Google Sheets with `--publish`."

"If you send me your export format, I can adapt this workflow to your business and remove manual reporting steps this week."

**Show on screen:**
- `data/processed/` folder contents.
- Open `weekly_summary.csv` and `top_products.csv`.
- Open `weekly_report.xlsx`.
- Briefly point to README publish section (`--publish`, Sheets tab names).

## Result sentence
"One raw export goes in; clean files, weekly metrics, and a stakeholder-ready report come out in under a minute."

## CTA
"DM me if you want this customized for your exact CSV exports and reporting cadence."
