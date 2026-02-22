from __future__ import annotations

import csv
import random
from datetime import datetime, timedelta
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = ROOT / "data" / "raw" / "orders_export.csv"

CHANNELS = ["web", "marketplace", "retail", "affiliate"]
PRODUCTS: list[tuple[str, float]] = [
    ("Wireless Mouse", 24.99),
    ("USB-C Cable", 9.5),
    ("Noise Cancelling Headphones", 199.0),
    ("Mechanical Keyboard", 89.99),
    ("Laptop Stand", 34.5),
    ("Webcam", 59.99),
    ("Monitor 27in", 229.0),
    ("Office Chair", 189.0),
    ("Desk Lamp", 39.95),
    ("External SSD 1TB", 119.0),
    ("Ergonomic Mouse", 49.0),
    ("HDMI Adapter", 12.0),
    ("Portable Charger", 29.99),
    ("Standing Desk", 349.0),
]
DATE_FORMATS = ["%Y-%m-%d", "%m/%d/%Y", "%b %d %Y"]
INVALID_DATES = ["2025-02-31", "not_a_date", "13/40/2025", "2025-14-10", "bad-date", "2025-00-12"]


def format_price(value: float) -> str:
    styles = [
        f"{value:.2f}",
        f"${value:.2f}",
        f"{value:,.2f}",
        f"USD {value:.2f}",
    ]
    return random.choices(styles, weights=[6, 2, 1, 1], k=1)[0]


def random_date_str(day: datetime, mixed: bool) -> str:
    if not mixed:
        return day.strftime("%Y-%m-%d")
    fmt = random.choice(DATE_FORMATS)
    return day.strftime(fmt)


def generate_rows(total_rows: int = 210) -> list[dict[str, str]]:
    random.seed(42)
    start = datetime(2025, 1, 6)
    rows: list[dict[str, str]] = []

    messy_count = 0
    all_rows = list(range(total_rows))
    parseable_mixed_rows = set(random.sample(all_rows, 6))

    invalid_date_candidates = [i for i in all_rows if i not in parseable_mixed_rows]
    invalid_date_rows = set(random.sample(invalid_date_candidates, 6))

    currency_candidates = [i for i in all_rows if i not in invalid_date_rows]
    currency_rows = set(random.sample(currency_candidates, 6))
    duplicate_rows = set(random.sample([i for i in range(20, total_rows)], 6))

    for i in range(total_rows):
        order_num = 20001 + i
        day = start + timedelta(days=random.randint(0, 62))
        product, base_price = random.choice(PRODUCTS)
        units = random.choices([1, 2, 3, 4, 5, 6, 8], weights=[25, 20, 18, 12, 10, 8, 7], k=1)[0]
        status = random.choices(["paid", "cancelled", "refunded"], weights=[82, 10, 8], k=1)[0]
        channel = random.choices(CHANNELS, weights=[40, 30, 20, 10], k=1)[0]

        if i in duplicate_rows and rows:
            reused = random.choice(rows)
            order_id = reused["order_id"]
            messy_count += 1
        else:
            order_id = f"ORD-{order_num}"

        if i in invalid_date_rows:
            order_date = random.choice(INVALID_DATES)
            messy_count += 1
        else:
            order_date = random_date_str(day, mixed=(i in parseable_mixed_rows))
            if i in parseable_mixed_rows:
                messy_count += 1

        price_multiplier = random.uniform(0.92, 1.15)
        raw_price = round(base_price * price_multiplier, 2)
        price = f"{raw_price:.2f}"
        if i in currency_rows:
            price = format_price(raw_price)
            messy_count += 1

        row = {
            "order_id": order_id,
            "order_date": order_date,
            "channel": channel,
            "product": product,
            "units": str(units),
            "price": price,
            "status": status,
            "coupon_code": random.choice(["", "SAVE10", "BULK5", "SPRING", ""]),
            "customer_note": random.choice(["", "gift", "rush", "", "repeat buyer"]),
        }
        rows.append(row)

    messy_ratio = messy_count / total_rows
    print(f"Generated {total_rows} rows ({messy_count} messy rows, {messy_ratio:.1%}).")
    return rows


def main() -> None:
    rows = generate_rows()
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp_path = OUTPUT_PATH.with_suffix(".csv.tmp")

    with tmp_path.open("w", newline="", encoding="utf-8") as fp:
        writer = csv.DictWriter(fp, fieldnames=list(rows[0].keys()))
        writer.writeheader()
        writer.writerows(rows)

    tmp_path.replace(OUTPUT_PATH)
    print(f"Wrote example export to {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
