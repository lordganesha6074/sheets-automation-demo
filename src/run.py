from __future__ import annotations

import argparse
import json
import logging
import os
from datetime import datetime, timezone
from pathlib import Path

import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
DEFAULT_INPUT_PATH = ROOT / "data" / "raw" / "orders_export.csv"
DEFAULT_OUTPUT_DIR = ROOT / "data" / "processed"

LOGGER = logging.getLogger(__name__)


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    lookup = {col.strip().lower(): col for col in df.columns}
    for candidate in candidates:
        if candidate.lower() in lookup:
            return lookup[candidate.lower()]
    return None


def to_numeric_currency(series: pd.Series) -> pd.Series:
    cleaned = (
        series.astype(str)
        .str.replace(r"[^\d.\-]", "", regex=True)
        .replace({"": pd.NA, ".": pd.NA, "-": pd.NA})
    )
    return pd.to_numeric(cleaned, errors="coerce")


def quarantine_rows(
    dropped_rows: list[pd.DataFrame],
    df: pd.DataFrame,
    mask: pd.Series,
    reason: str,
) -> pd.DataFrame:
    if not mask.any():
        return df

    quarantined = df.loc[mask].copy()
    quarantined["drop_reason"] = reason
    dropped_rows.append(quarantined)
    return df.loc[~mask].copy()


def main(
    publish: bool = False,
    input_path: Path = DEFAULT_INPUT_PATH,
    outdir: Path = DEFAULT_OUTPUT_DIR,
    paid_only: bool = True,
    start_date: str | None = None,
    end_date: str | None = None,
) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    clean_path = outdir / "clean_orders.csv"
    weekly_path = outdir / "weekly_summary.csv"
    top_products_path = outdir / "top_products.csv"
    report_path = outdir / "weekly_report.xlsx"
    quarantine_path = outdir / "quarantine_bad_rows.csv"
    quality_report_path = outdir / "data_quality_report.json"

    df = pd.read_csv(input_path)
    df["_source_row"] = range(len(df))
    counters = {
        "raw_rows": len(df),
        "dropped_unparseable_date": 0,
        "dropped_invalid_amount": 0,
        "dropped_missing_required": 0,
        "filtered_paid": 0,
        "deduped": 0,
        "final_rows": 0,
    }
    dropped_rows: list[pd.DataFrame] = []

    LOGGER.info("[rows] raw input: %s", counters["raw_rows"])
    LOGGER.debug("[io] input=%s outdir=%s", input_path, outdir)

    order_id_col = find_column(df, ["order_id", "order id", "id"])
    order_date_col = find_column(df, ["order_date", "order timestamp", "order_datetime", "date"])
    status_col = find_column(df, ["status", "payment_status"])
    channel_col = find_column(df, ["channel", "source", "sales_channel"])
    product_col = find_column(df, ["product", "product_name", "item"])
    units_col = find_column(df, ["units", "quantity", "qty"])
    price_col = find_column(df, ["price", "unit_price", "amount"])
    revenue_col = find_column(df, ["revenue", "total", "total_revenue"])

    required = {
        "order_id": order_id_col,
        "order_date": order_date_col,
        "status": status_col,
        "channel": channel_col,
        "product": product_col,
        "units": units_col,
    }
    missing = [name for name, col in required.items() if col is None]
    if missing:
        raise ValueError(f"Missing required column(s): {', '.join(missing)}")
    if price_col is None and revenue_col is None:
        raise ValueError("Need either a price column or a revenue column.")

    LOGGER.debug(
        (
            "[columns] order_id=%s order_date=%s status=%s channel=%s product=%s "
            "units=%s price=%s revenue=%s"
        ),
        order_id_col,
        order_date_col,
        status_col,
        channel_col,
        product_col,
        units_col,
        price_col,
        revenue_col,
    )

    missing_required_mask = df[[order_id_col, status_col, channel_col, product_col, units_col]]
    missing_required_mask = missing_required_mask.isna().any(axis=1)
    counters["dropped_missing_required"] = int(missing_required_mask.sum())
    df = quarantine_rows(dropped_rows, df, missing_required_mask, "missing_required")

    df[order_date_col] = pd.to_datetime(df[order_date_col], errors="coerce")
    invalid_date_mask = df[order_date_col].isna()
    counters["dropped_unparseable_date"] = int(invalid_date_mask.sum())
    if counters["dropped_unparseable_date"]:
        LOGGER.info(
            "[dates] dropped %s rows with unparseable order_date",
            counters["dropped_unparseable_date"],
        )
    df = quarantine_rows(dropped_rows, df, invalid_date_mask, "unparseable_order_date")
    LOGGER.info("[rows] after date parsing: %s", len(df))

    if start_date:
        start_dt = pd.to_datetime(start_date)
        before_start_mask = df[order_date_col] < start_dt
        LOGGER.debug("[window] applying start date >= %s", start_dt.date())
        df = quarantine_rows(dropped_rows, df, before_start_mask, "before_start_date")
    if end_date:
        end_dt = pd.to_datetime(end_date)
        after_end_mask = df[order_date_col] > end_dt
        LOGGER.debug("[window] applying end date <= %s", end_dt.date())
        df = quarantine_rows(dropped_rows, df, after_end_mask, "after_end_date")
    if start_date or end_date:
        LOGGER.info("[rows] after optional date window: %s", len(df))

    df[units_col] = pd.to_numeric(df[units_col], errors="coerce")
    invalid_units = df[units_col].isna().sum()
    if invalid_units:
        LOGGER.info(
            "[units] coerced %s missing/invalid values to 0 before aggregation",
            invalid_units,
        )
    df[units_col] = df[units_col].fillna(0)
    df[units_col] = df[units_col].round().astype("Int64")

    if price_col is not None:
        df[price_col] = to_numeric_currency(df[price_col])
        invalid_price = df[price_col].isna().sum()
        if invalid_price:
            LOGGER.info(
                "[price] %s missing/invalid values detected; defaulting to 0 "
                "when revenue is computed from units * price",
                invalid_price,
            )

    if revenue_col is not None:
        df[revenue_col] = to_numeric_currency(df[revenue_col])
        invalid_revenue = df[revenue_col].isna().sum()
        if invalid_revenue:
            LOGGER.info(
                "[revenue] %s missing/invalid values detected; backfilling with "
                "units * price where possible, otherwise 0",
                invalid_revenue,
            )

    invalid_amount_mask = pd.Series(False, index=df.index)
    if price_col is not None:
        invalid_amount_mask = invalid_amount_mask | df[price_col].isna()
    if revenue_col is not None:
        invalid_amount_mask = invalid_amount_mask & df[revenue_col].isna()

    counters["dropped_invalid_amount"] = int(invalid_amount_mask.sum())
    df = quarantine_rows(dropped_rows, df, invalid_amount_mask, "invalid_amount")

    df[status_col] = df[status_col].astype(str).str.strip().str.lower()
    if paid_only:
        paid_mask = df[status_col] == "paid"
        counters["filtered_paid"] = int((~paid_mask).sum())
        df = quarantine_rows(dropped_rows, df, ~paid_mask, "status_not_paid")
        LOGGER.info("[rows] after paid filter: %s", len(df))
    else:
        LOGGER.info("[rows] skipping paid-only filter")

    missing_order_ids = df[order_id_col].isna().sum()
    if missing_order_ids:
        LOGGER.info(
            "[order_id] excluding %s paid rows with missing order_id from dedup/order "
            "metrics",
            missing_order_ids,
        )

    df_with_id = df[df[order_id_col].notna()].copy()
    df_missing_id = df[df[order_id_col].isna()].copy()

    # Deterministic ordering ensures deduplication keeps a reproducible "latest" row.
    df_with_id = df_with_id.sort_values(
        by=[order_date_col, "_source_row"],
        ascending=[True, True],
        kind="mergesort",
    )
    dedup_mask = df_with_id.duplicated(subset=[order_id_col], keep="last")
    counters["deduped"] = int(dedup_mask.sum())
    deduped_rows = df_with_id.loc[dedup_mask].copy()
    if not deduped_rows.empty:
        deduped_rows["drop_reason"] = "duplicate_order_id"
        dropped_rows.append(deduped_rows)
    df_with_id = df_with_id.drop_duplicates(subset=[order_id_col], keep="last")
    LOGGER.info("[rows] after dedup: %s", len(df_with_id))

    # Keep rows with missing order_id for non-order metrics, but never include them in order counts.
    df = pd.concat([df_with_id, df_missing_id], ignore_index=True)
    counters["final_rows"] = len(df)

    units_revenue = df[units_col].astype(float)
    if price_col is not None:
        price_series = df[price_col].fillna(0)
    else:
        price_series = 0
    computed_revenue = units_revenue * price_series

    if revenue_col is not None:
        df["_revenue"] = df[revenue_col].fillna(computed_revenue)
    else:
        df["_revenue"] = computed_revenue

    df["week"] = df[order_date_col].dt.to_period("W-MON").dt.start_time.dt.strftime("%Y-%m-%d")

    clean_df = df.copy()
    clean_df[order_date_col] = clean_df[order_date_col].dt.strftime("%Y-%m-%d")
    clean_df = clean_df.drop(columns=["_source_row"])
    clean_df.to_csv(clean_path, index=False)

    weekly_summary = (
        df.groupby(["week", channel_col], dropna=False)
        .agg(
            orders=(order_id_col, "nunique"),
            units=(units_col, "sum"),
            revenue=("_revenue", "sum"),
        )
        .reset_index()
        .rename(columns={channel_col: "channel"})
        [["week", "channel", "orders", "units", "revenue"]]
        .sort_values(["week", "channel"])
    )

    # KPI columns appended after base summary columns.
    weekly_summary["aov"] = (
        weekly_summary["revenue"].div(weekly_summary["orders"]).where(weekly_summary["orders"] != 0, 0)
    )

    wow_sorted = weekly_summary.sort_values(["channel", "week"]).copy()
    prior_revenue = wow_sorted.groupby("channel")["revenue"].shift(1)
    wow_sorted["revenue_wow_pct"] = (
        wow_sorted["revenue"].sub(prior_revenue).div(prior_revenue).mul(100)
    )
    wow_sorted["revenue_wow_pct"] = wow_sorted["revenue_wow_pct"].where(
        prior_revenue.notna() & prior_revenue.ne(0),
        0,
    )

    weekly_total_revenue = wow_sorted.groupby("week")["revenue"].transform("sum")
    wow_sorted["channel_revenue_share_pct"] = (
        wow_sorted["revenue"].div(weekly_total_revenue).mul(100)
    )
    wow_sorted["channel_revenue_share_pct"] = wow_sorted["channel_revenue_share_pct"].where(
        weekly_total_revenue.ne(0),
        0,
    )

    weekly_summary = wow_sorted.sort_values(["week", "channel"])
    weekly_summary = weekly_summary[
        [
            "week",
            "channel",
            "orders",
            "units",
            "revenue",
            "aov",
            "revenue_wow_pct",
            "channel_revenue_share_pct",
        ]
    ]

    top_products = (
        df.groupby(product_col, dropna=False)
        .agg(
            units=(units_col, "sum"),
            revenue=("_revenue", "sum"),
        )
        .reset_index()
        .rename(columns={product_col: "product"})
        .sort_values(["revenue", "units"], ascending=[False, False])
    )


    weekly_summary["revenue"] = weekly_summary["revenue"].round(2)
    weekly_summary["aov"] = weekly_summary["aov"].round(2)
    weekly_summary["revenue_wow_pct"] = weekly_summary["revenue_wow_pct"].round(2)
    weekly_summary["channel_revenue_share_pct"] = weekly_summary["channel_revenue_share_pct"].round(2)
    top_products["revenue"] = top_products["revenue"].round(2)

    weekly_summary.to_csv(weekly_path, index=False, float_format="%.2f")
    top_products.to_csv(top_products_path, index=False, float_format="%.2f")

    with pd.ExcelWriter(report_path, engine="openpyxl") as writer:
        weekly_summary.to_excel(writer, sheet_name="Summary", index=False)
        top_products.to_excel(writer, sheet_name="Top Products", index=False)

        summary_sheet = writer.sheets["Summary"]
        summary_revenue_col = weekly_summary.columns.get_loc("revenue") + 1
        for row_idx in range(2, len(weekly_summary) + 2):
            summary_sheet.cell(row=row_idx, column=summary_revenue_col).number_format = "0.00"

        top_products_sheet = writer.sheets["Top Products"]
        top_products_revenue_col = top_products.columns.get_loc("revenue") + 1
        for row_idx in range(2, len(top_products) + 2):
            top_cell = top_products_sheet.cell(
                row=row_idx, column=top_products_revenue_col
            )
            top_cell.number_format = "0.00"

    quarantine_df = (
        pd.concat(dropped_rows, ignore_index=True)
        if dropped_rows
        else pd.DataFrame(columns=[*df.columns, "drop_reason"])
    )
    quarantine_df.to_csv(quarantine_path, index=False)

    quality_report = {
        **counters,
        "run_timestamp": datetime.now(timezone.utc).isoformat(),
    }
    with quality_report_path.open("w", encoding="utf-8") as fp:
        json.dump(quality_report, fp, indent=2)

    if publish:
        try:
            from publish_google_sheets import publish_csvs
        except ImportError as exc:
            raise RuntimeError(
                "Publishing requires google dependencies. Install requirements.txt."
            ) from exc

        sheet_id = publish_csvs(
            sheet_id=os.getenv("GOOGLE_SHEET_ID"),
            weekly_csv=weekly_path,
            top_csv=top_products_path,
        )
        LOGGER.info("[publish] Google Sheet ID: %s", sheet_id)
        LOGGER.info(
            "[publish] Google Sheet URL: https://docs.google.com/spreadsheets/d/%s",
            sheet_id,
        )


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Generate weekly reports from raw orders data")
    parser.add_argument(
        "--input",
        default="data/raw/orders_export.csv",
        help="Path to input CSV (default: data/raw/orders_export.csv)",
    )
    parser.add_argument(
        "--outdir",
        default="data/processed",
        help="Output directory (default: data/processed)",
    )
    parser.add_argument(
        "--paid-only",
        action=argparse.BooleanOptionalAction,
        default=True,
        help="Filter rows to paid status only (use --no-paid-only to disable)",
    )
    parser.add_argument(
        "--start-date",
        default=None,
        help="Optional start date (YYYY-MM-DD), inclusive",
    )
    parser.add_argument(
        "--end-date",
        default=None,
        help="Optional end date (YYYY-MM-DD), inclusive",
    )
    parser.add_argument(
        "--verbose",
        action="store_true",
        help="Enable DEBUG logging",
    )
    parser.add_argument(
        "--publish",
        action="store_true",
        help="Publish weekly_summary.csv and top_products.csv to Google Sheets",
    )
    args = parser.parse_args()
    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(levelname)s:%(name)s:%(message)s",
    )

    main(
        publish=args.publish,
        input_path=ROOT / args.input,
        outdir=ROOT / args.outdir,
        paid_only=args.paid_only,
        start_date=args.start_date,
        end_date=args.end_date,
    )
