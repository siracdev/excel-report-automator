"""Create messy sample sales data for demoing the Excel Report Automator.

Run it once:  python generate_sample_data.py

It writes three files to ./sample_data that are deliberately inconsistent,
like the files real clients send:

  sales_january.csv   tidy column names, ISO dates (2026-01-05)
  sales_february.csv  different column names, extra spaces, lower-case regions,
                      day/month/year dates (05/02/2026)
  sales_march.xlsx    Excel file, mixed date formats, duplicated orders,
                      missing values, one product typed in lower case
"""

import random
from datetime import date
from pathlib import Path

import pandas as pd

random.seed(42)  # same data on every run, so results are reproducible

OUTPUT_DIR = Path("sample_data")
YEAR = 2026

PRODUCTS = {
    "Laptop Stand": 29.90,
    "USB-C Hub": 45.00,
    "Mechanical Keyboard": 89.50,
    "Webcam HD": 59.00,
    "Desk Lamp": 24.75,
}
REGIONS = ["North", "South", "East", "West"]
REPS = ["Alice", "Bob", "Carla", "David"]


def make_orders(month: int, count: int) -> pd.DataFrame:
    """Return `count` random orders for the given month."""
    rows = []
    for i in range(count):
        product = random.choice(list(PRODUCTS))
        rows.append(
            {
                "order_id": f"{month:02d}-{1000 + i}",
                "date": date(YEAR, month, random.randint(1, 28)),
                "product": product,
                "region": random.choice(REGIONS),
                "quantity": random.randint(1, 12),
                "unit_price": PRODUCTS[product],
                "sales_rep": random.choice(REPS),
            }
        )
    return pd.DataFrame(rows)


def main() -> None:
    OUTPUT_DIR.mkdir(exist_ok=True)

    # January: tidy file.
    jan = make_orders(1, 40)
    jan["date"] = jan["date"].map(lambda d: d.isoformat())
    jan = jan.rename(
        columns={
            "order_id": "Order ID", "date": "Date", "product": "Product", "region": "Region",
            "quantity": "Quantity", "unit_price": "Unit Price", "sales_rep": "Sales Rep",
        }
    )
    jan.to_csv(OUTPUT_DIR / "sales_january.csv", index=False)

    # February: other column names, stray spaces, lower-case text, day-first dates.
    feb = make_orders(2, 45)
    feb["date"] = feb["date"].map(lambda d: d.strftime("%d/%m/%Y"))
    feb["product"] = feb["product"].map(lambda p: f"  {p} ")
    feb["region"] = feb["region"].str.lower()
    feb = feb.rename(
        columns={
            "order_id": "OrderID", "date": "Order Date", "product": "Product Name",
            "quantity": "Qty", "unit_price": "Price", "sales_rep": "Rep",
        }
    )
    feb.to_csv(OUTPUT_DIR / "sales_february.csv", index=False)

    # March: Excel file with several problems.
    mar = make_orders(3, 50)
    mar["date"] = [
        d.strftime("%Y-%m-%d") if i % 2 == 0 else d.strftime("%d/%m/%Y")
        for i, d in enumerate(mar["date"])
    ]
    mar.loc[8, "product"] = "usb-c hub"
    mar["quantity"] = mar["quantity"].astype(float)
    mar.loc[5, "quantity"] = float("nan")      # missing quantity
    mar.loc[22, "unit_price"] = float("nan")   # missing price
    mar = pd.concat([mar, mar.iloc[[3, 17]]], ignore_index=True)  # two duplicated orders
    mar = mar.rename(
        columns={
            "order_id": "Order ID", "date": "Date", "product": "Product", "region": "Region",
            "quantity": "Quantity", "unit_price": "Unit Price", "sales_rep": "Sales Rep",
        }
    )
    mar.to_excel(OUTPUT_DIR / "sales_march.xlsx", index=False)

    print(f"Sample files written to {OUTPUT_DIR}/:")
    for path in sorted(OUTPUT_DIR.iterdir()):
        print(f"  {path.name}")


if __name__ == "__main__":
    main()
