import sys
from pathlib import Path

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / "src"))

from run import to_numeric_currency


@pytest.mark.parametrize(
    ("raw", "expected"),
    [
        ("$1,234.50", 1234.50),
        (" - ", pd.NA),
        ("", pd.NA),
        ("EUR 9.99", 9.99),
        ("abc", pd.NA),
    ],
)
def test_to_numeric_currency_parses_and_coerces_invalid_values(raw: str, expected) -> None:
    series = pd.Series([raw])

    actual = to_numeric_currency(series).iloc[0]

    if pd.isna(expected):
        assert pd.isna(actual)
    else:
        assert actual == pytest.approx(expected)


def test_order_date_parsing_coerces_invalid_dates_and_marks_quarantine_mask(
) -> None:
    order_date_col = "order_date"
    df = pd.DataFrame(
        {
            order_date_col: ["2024-01-15", "not-a-date", "2024-03-20", "", None],
            "order_id": ["A-1", "A-2", "A-3", "A-4", "A-5"],
        }
    )

    df[order_date_col] = pd.to_datetime(df[order_date_col], errors="coerce")
    invalid_date_mask = df[order_date_col].isna()

    assert df.loc[0, order_date_col] == pd.Timestamp("2024-01-15")
    assert df.loc[2, order_date_col] == pd.Timestamp("2024-03-20")
    assert pd.isna(df.loc[1, order_date_col])
    assert pd.isna(df.loc[3, order_date_col])
    assert pd.isna(df.loc[4, order_date_col])

    # Same mask logic used in main before quarantine_rows(..., "unparseable_order_date")
    assert invalid_date_mask.tolist() == [False, True, False, True, True]
    assert df.loc[invalid_date_mask, "order_id"].tolist() == ["A-2", "A-4", "A-5"]


def test_dedup_flow_keeps_last_row_by_order_date_and_source_row() -> None:
    order_id_col = "order_id"
    order_date_col = "order_date"

    df_with_id = pd.DataFrame(
        {
            order_id_col: ["ORD-1", "ORD-1", "ORD-2", "ORD-2", "ORD-3"],
            order_date_col: pd.to_datetime(
                ["2024-01-01", "2024-01-02", "2024-01-03", "2024-01-03", "2024-01-04"]
            ),
            "_source_row": [0, 1, 2, 9, 4],
            "marker": [
                "older",
                "newer",
                "same_date_lower_source",
                "same_date_higher_source",
                "unique",
            ],
        }
    )

    df_with_id = df_with_id.sort_values(
        by=[order_date_col, "_source_row"],
        ascending=[True, True],
        kind="mergesort",
    )
    dedup_mask = df_with_id.duplicated(subset=[order_id_col], keep="last")
    deduped = df_with_id.drop_duplicates(subset=[order_id_col], keep="last")

    assert dedup_mask.sum() == 2
    assert deduped[order_id_col].tolist() == ["ORD-1", "ORD-2", "ORD-3"]
    assert (
        deduped.set_index(order_id_col).loc["ORD-1", "marker"]
        == "newer"
    )
    assert (
        deduped.set_index(order_id_col).loc["ORD-2", "marker"]
        == "same_date_higher_source"
    )
