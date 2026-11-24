"""Unit tests for the loading and cleaning logic.

Run from the project folder with:  python -m unittest -v
"""

import unittest

import pandas as pd

import report_automator as ra


def order(**changes) -> dict:
    """One valid order row; override any field with keyword arguments."""
    row = {
        "order_id": "A1",
        "date": "2026-01-05",
        "product": "Desk Lamp",
        "quantity": 2,
        "unit_price": 10.0,
        "region": "North",
        "sales_rep": "Alice",
        "source_file": "test.csv",
    }
    row.update(changes)
    return row


def clean(*rows):
    return ra.clean(pd.DataFrame(rows))


class StandardizeColumnsTests(unittest.TestCase):
    def test_alias_names_are_mapped(self):
        raw = pd.DataFrame(
            {"OrderID": ["1"], "Order Date": ["2026-01-01"], "Product Name": ["A"], "Qty": [1], "Price": [2.0]}
        )
        result = ra.standardize_columns(raw, "x.csv")
        for column in ra.REQUIRED_COLUMNS:
            self.assertIn(column, result.columns)

    def test_missing_required_column_raises(self):
        raw = pd.DataFrame({"Order ID": [1], "Date": ["2026-01-01"]})
        with self.assertRaisesRegex(ValueError, "quantity"):
            ra.standardize_columns(raw, "x.csv")

    def test_optional_columns_default_to_unknown(self):
        raw = pd.DataFrame(
            {"order_id": ["1"], "date": ["2026-01-01"], "product": ["A"], "quantity": [1], "unit_price": [2.0]}
        )
        result = ra.standardize_columns(raw, "x.csv")
        self.assertEqual(result.loc[0, "region"], "Unknown")
        self.assertEqual(result.loc[0, "source_file"], "x.csv")


class CleanTests(unittest.TestCase):
    def test_duplicate_orders_are_removed(self):
        df, stats = clean(order(), order())
        self.assertEqual(len(df), 1)
        self.assertEqual(stats["duplicates_removed"], 1)

    def test_rows_with_missing_values_are_dropped(self):
        df, stats = clean(order(order_id="A1"), order(order_id="A2", quantity=None))
        self.assertEqual(len(df), 1)
        self.assertEqual(stats["rows_invalid"], 1)

    def test_text_spelling_is_unified(self):
        df, _ = clean(
            order(order_id="1", product="usb-c hub"),
            order(order_id="2", product="  USB-C Hub "),
            order(order_id="3", product="USB-C Hub"),
        )
        self.assertEqual(list(df["product"].unique()), ["USB-C Hub"])

    def test_mixed_date_formats(self):
        df, _ = clean(
            order(order_id="1", date="2026-03-05"),
            order(order_id="2", date="05/03/2026"),
        )
        self.assertTrue((df["date"] == pd.Timestamp(2026, 3, 5)).all())

    def test_revenue_and_month(self):
        df, _ = clean(order(quantity=3, unit_price=19.99))
        self.assertAlmostEqual(df.loc[0, "revenue"], 59.97)
        self.assertEqual(df.loc[0, "month"], "2026-01")


class SummaryTests(unittest.TestCase):
    def test_shares_add_up_to_one(self):
        df, _ = clean(
            order(order_id="1", product="A", quantity=1, unit_price=10),
            order(order_id="2", product="B", quantity=3, unit_price=10),
        )
        tables, kpis = ra.build_summary(df)
        self.assertAlmostEqual(tables["product"]["share"].sum(), 1.0)
        self.assertEqual(tables["product"].loc[0, "product"], "B")  # biggest first
        self.assertEqual(dict((k[0], k[1]) for k in kpis)["Total revenue"], 40)


if __name__ == "__main__":
    unittest.main()
