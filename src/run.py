from __future__ import annotations

from pathlib import Path

import pandas as pd


ROOT = Path(__file__).resolve().parents[1]
RAW_PATH = ROOT / "data" / "raw" / "orders_export.csv"
PROCESSED_DIR = ROOT / "data" / "processed"
CLEAN_PATH = PROCESSED_DIR / "clean_orders.csv"
WEEKLY_PATH = PROCESSED_DIR / "weekly_summary.csv"
TOP_PRODUCTS_PATH = PROCESSED_DIR / "top_products.csv"
REPORT_PATH = PROCESSED_DIR / "weekly_report.xlsx"


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


def main() -> None:
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    df = pd.read_csv(RAW_PATH)
    df["_source_row"] = range(len(df))
    print(f"[rows] raw input: {len(df)}")

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

    df[order_date_col] = pd.to_datetime(df[order_date_col], errors="coerce")
    invalid_dates = df[order_date_col].isna().sum()
    if invalid_dates:
        print(f"[dates] dropped {invalid_dates} rows with unparseable order_date")
    df = df.dropna(subset=[order_date_col]).copy()
    print(f"[rows] after date parsing: {len(df)}")

    df[units_col] = pd.to_numeric(df[units_col], errors="coerce")
    invalid_units = df[units_col].isna().sum()
    if invalid_units:
        print(
            f"[units] coerced {invalid_units} missing/invalid values to 0 "
            "before aggregation"
        )
    df[units_col] = df[units_col].fillna(0)
    df[units_col] = df[units_col].round().astype("Int64")

    if price_col is not None:
        df[price_col] = to_numeric_currency(df[price_col])
        invalid_price = df[price_col].isna().sum()
        if invalid_price:
            print(
                f"[price] {invalid_price} missing/invalid values detected; "
                "defaulting to 0 when revenue is computed from units * price"
            )

    if revenue_col is not None:
        df[revenue_col] = to_numeric_currency(df[revenue_col])
        invalid_revenue = df[revenue_col].isna().sum()
        if invalid_revenue:
            print(
                f"[revenue] {invalid_revenue} missing/invalid values detected; "
                "backfilling with units * price where possible, otherwise 0"
            )

    df[status_col] = df[status_col].astype(str).str.strip().str.lower()
    df = df[df[status_col] == "paid"].copy()
    print(f"[rows] after paid filter: {len(df)}")

    missing_order_ids = df[order_id_col].isna().sum()
    if missing_order_ids:
        print(
            f"[order_id] excluding {missing_order_ids} paid rows with missing order_id "
            "from dedup/order metrics"
        )

    df_with_id = df[df[order_id_col].notna()].copy()
    df_missing_id = df[df[order_id_col].isna()].copy()

    # Deterministic ordering ensures deduplication keeps a reproducible "latest" row.
    df_with_id = df_with_id.sort_values(
        by=[order_date_col, "_source_row"],
        ascending=[True, True],
        kind="mergesort",
    )
    df_with_id = df_with_id.drop_duplicates(subset=[order_id_col], keep="last")
    print(f"[rows] after dedup: {len(df_with_id)}")

    # Keep rows with missing order_id for non-order metrics, but never include them in order counts.
    df = pd.concat([df_with_id, df_missing_id], ignore_index=True)

    if revenue_col is not None:
        df["_revenue"] = df[revenue_col].fillna(df[units_col].astype(float) * df.get(price_col, 0).fillna(0) if price_col else 0)
    else:
        df["_revenue"] = df[units_col].astype(float) * df[price_col].fillna(0)

    df["week"] = df[order_date_col].dt.to_period("W-MON").dt.start_time.dt.strftime("%Y-%m-%d")

    clean_df = df.copy()
    clean_df[order_date_col] = clean_df[order_date_col].dt.strftime("%Y-%m-%d")
    clean_df = clean_df.drop(columns=["_source_row"])
    clean_df.to_csv(CLEAN_PATH, index=False)

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


    weekly_summary.to_csv(WEEKLY_PATH, index=False)
    top_products.to_csv(TOP_PRODUCTS_PATH, index=False)

    with pd.ExcelWriter(REPORT_PATH, engine="openpyxl") as writer:
        weekly_summary.to_excel(writer, sheet_name="Summary", index=False)
        top_products.to_excel(writer, sheet_name="Top Products", index=False)


if __name__ == "__main__":
    main()
