#!/usr/bin/env python3
"""Excel Report Automator

Combine a folder of messy sales files (CSV / Excel) into one clean,
formatted Excel report with summary tables and charts.

Example:
    python report_automator.py --input sample_data --output output/sales_report.xlsx
"""

from __future__ import annotations

import argparse
import sys
from datetime import date, datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import PageSetupProperties

# ---------------------------------------------------------------------------
# 1. Configuration
# ---------------------------------------------------------------------------

SUPPORTED_EXTENSIONS = {".csv", ".xlsx", ".xlsm"}

REQUIRED_COLUMNS = ["order_id", "date", "product", "quantity", "unit_price"]
OPTIONAL_COLUMNS = ["region", "sales_rep"]

# Different clients name the same column differently.
# Add your own spellings here (lower-case, underscores written as spaces).
COLUMN_ALIASES = {
    "order_id": {"order id", "orderid", "id", "order number"},
    "date": {"date", "order date", "orderdate"},
    "product": {"product", "product name", "item"},
    "region": {"region", "area"},
    "quantity": {"quantity", "qty", "units"},
    "unit_price": {"unit price", "unitprice", "price"},
    "sales_rep": {"sales rep", "salesrep", "rep", "salesperson"},
}


def _normalize_name(name) -> str:
    return str(name).strip().lower().replace("_", " ")


# Reverse lookup: "qty" -> "quantity", "order date" -> "date", ...
_ALIAS_LOOKUP = {_normalize_name(std): std for std in COLUMN_ALIASES}
for _standard, _aliases in COLUMN_ALIASES.items():
    for _alias in _aliases:
        _ALIAS_LOOKUP[_alias] = _standard


# ---------------------------------------------------------------------------
# 2. Loading files
# ---------------------------------------------------------------------------


def find_input_files(folder: Path) -> list[Path]:
    """Return all supported files in the folder, ignoring temp/hidden files."""
    return sorted(
        p
        for p in folder.iterdir()
        if p.is_file()
        and p.suffix.lower() in SUPPORTED_EXTENSIONS
        and not p.name.startswith(("~$", "."))
    )


def read_file(path: Path) -> pd.DataFrame:
    """Read one CSV or Excel file into a DataFrame."""
    if path.suffix.lower() == ".csv":
        # sep=None lets pandas detect comma vs semicolon; we also try a few
        # common encodings because exported CSVs are not always UTF-8.
        for encoding in ("utf-8-sig", "cp1254", "latin-1"):
            try:
                return pd.read_csv(path, sep=None, engine="python", encoding=encoding)
            except UnicodeDecodeError:
                continue
        raise ValueError("could not decode the file")
    return pd.read_excel(path)


def standardize_columns(raw: pd.DataFrame, source_name: str) -> pd.DataFrame:
    """Rename columns to the standard names and keep only the ones we need."""
    renamed = raw.rename(
        columns=lambda c: _ALIAS_LOOKUP.get(_normalize_name(c), _normalize_name(c))
    )

    known = [c for c in renamed.columns if c in REQUIRED_COLUMNS + OPTIONAL_COLUMNS]
    duplicated = sorted({c for c in known if known.count(c) > 1})
    if duplicated:
        raise ValueError(f"more than one column maps to: {', '.join(duplicated)}")

    missing = [c for c in REQUIRED_COLUMNS if c not in renamed.columns]
    if missing:
        raise ValueError(
            f"missing required column(s): {', '.join(missing)} "
            f"(found: {', '.join(map(str, raw.columns))})"
        )

    for column in OPTIONAL_COLUMNS:
        if column not in renamed.columns:
            renamed[column] = "Unknown"

    result = renamed[REQUIRED_COLUMNS + OPTIONAL_COLUMNS].copy()
    result["source_file"] = source_name
    return result


# ---------------------------------------------------------------------------
# 3. Cleaning
# ---------------------------------------------------------------------------


def unify_text(series: pd.Series) -> pd.Series:
    """Trim spaces and make spelling consistent, ignoring case.

    'usb-c hub', ' USB-C Hub ' and 'USB-C Hub' all become 'USB-C Hub'
    (the most common spelling wins).
    """
    cleaned = series.astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
    cleaned = cleaned.mask(cleaned == "")
    lowered = cleaned.str.lower()
    spelling = {key: group.value_counts().idxmax() for key, group in cleaned.groupby(lowered)}
    return lowered.map(spelling)


def parse_dates(series: pd.Series) -> pd.Series:
    """Parse dates written in different formats.

    ISO dates (2026-03-05) are read year-first. Anything else, like
    05/03/2026, is read day-first (5 March). Unreadable values become NaT.
    """
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    text = series.astype("string").str.strip()
    iso = pd.to_datetime(text.str.slice(0, 10), format="%Y-%m-%d", errors="coerce")
    other = pd.to_datetime(text.where(iso.isna()), format="mixed", dayfirst=True, errors="coerce")
    return iso.fillna(other)


def clean(data: pd.DataFrame) -> tuple[pd.DataFrame, dict]:
    """Clean the combined data. Returns the clean table and a log of what changed."""
    df = data.copy()
    stats = {"rows_read": len(df)}

    order_id = df["order_id"].astype("string").str.strip().str.replace(r"\.0$", "", regex=True)
    df["order_id"] = order_id.mask(order_id == "")
    df["date"] = parse_dates(df["date"])
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
    df["unit_price"] = pd.to_numeric(df["unit_price"], errors="coerce")
    for column in ("product", "region", "sales_rep"):
        df[column] = unify_text(df[column])
    df["region"] = df["region"].fillna("Unknown")
    df["sales_rep"] = df["sales_rep"].fillna("Unknown")

    # Drop rows that are missing something we cannot work without.
    before = len(df)
    df = df.dropna(subset=REQUIRED_COLUMNS)
    stats["rows_invalid"] = before - len(df)

    # The same order appearing twice is counted once.
    before = len(df)
    df = df.drop_duplicates(subset="order_id", keep="first")
    stats["duplicates_removed"] = before - len(df)

    if (df["quantity"] % 1 == 0).all():
        df["quantity"] = df["quantity"].astype(int)
    df["revenue"] = (df["quantity"] * df["unit_price"]).round(2)
    df["month"] = df["date"].dt.strftime("%Y-%m")

    df = df.sort_values(["date", "order_id"]).reset_index(drop=True)
    stats["rows_clean"] = len(df)
    return df, stats


# ---------------------------------------------------------------------------
# 4. Analysis
# ---------------------------------------------------------------------------


def totals_by(df: pd.DataFrame, column: str) -> pd.DataFrame:
    """Revenue, units and orders for each value of `column`."""
    table = (
        df.groupby(column)
        .agg(revenue=("revenue", "sum"), units=("quantity", "sum"), orders=("order_id", "nunique"))
        .reset_index()
    )
    table["share"] = table["revenue"] / table["revenue"].sum()
    return table


def build_summary(df: pd.DataFrame, currency: str = "$") -> tuple[dict, list]:
    """Build the summary tables and the headline numbers."""
    by_product = totals_by(df, "product").sort_values("revenue", ascending=False)
    by_region = totals_by(df, "region").sort_values("revenue", ascending=False)
    by_month = totals_by(df, "month").sort_values("month").drop(columns="share")
    by_month["growth"] = by_month["revenue"].pct_change()
    tables = {
        "product": by_product.reset_index(drop=True),
        "region": by_region.reset_index(drop=True),
        "month": by_month.reset_index(drop=True),
    }

    total = df["revenue"].sum()
    orders = df["order_id"].nunique()
    top_product = tables["product"].iloc[0]
    top_region = tables["region"].iloc[0]
    best_month = tables["month"].loc[tables["month"]["revenue"].idxmax()]

    def detail(row) -> str:
        return f"{currency}{row['revenue']:,.2f} ({row['share']:.1%} of revenue)"

    # (label, value, format, detail) - format is "money", "int" or None
    kpis = [
        ("Period", f"{df['date'].min():%Y-%m-%d} to {df['date'].max():%Y-%m-%d}", None, None),
        ("Total revenue", total, "money", None),
        ("Orders", orders, "int", None),
        ("Units sold", df["quantity"].sum(), "int", None),
        ("Average order value", total / orders, "money", None),
        ("Top product", top_product["product"], None, detail(top_product)),
        ("Top region", top_region["region"], None, detail(top_region)),
        ("Best month", best_month["month"], None, f"{currency}{best_month['revenue']:,.2f}"),
    ]
    return tables, kpis


# ---------------------------------------------------------------------------
# 5. Excel output
# ---------------------------------------------------------------------------

HEADER_FILL = PatternFill("solid", start_color="1F4E78")
HEADER_FONT = Font(bold=True, color="FFFFFF")
TITLE_FONT = Font(bold=True, size=16, color="1F4E78")
SECTION_FONT = Font(bold=True, size=12, color="1F4E78")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

INTEGER = "#,##0"
PERCENT = "0.0%"
DATE_FORMAT = "yyyy-mm-dd"

HEADER_LABELS = {
    "order_id": "Order ID",
    "share": "Share of Revenue",
    "growth": "MoM Growth",
}

DATA_COLUMNS = [
    "order_id", "date", "product", "region", "quantity",
    "unit_price", "revenue", "sales_rep", "month", "source_file",
]


def write_table(ws, df, start_row=1, start_col=1, formats=None) -> int:
    """Write a DataFrame as a styled table. Returns the last row used."""
    formats = formats or {}
    for offset, column in enumerate(df.columns):
        label = HEADER_LABELS.get(column, column.replace("_", " ").title())
        cell = ws.cell(row=start_row, column=start_col + offset, value=label)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER

    for r, row in enumerate(df.itertuples(index=False), start=start_row + 1):
        for offset, value in enumerate(row):
            if pd.isna(value):
                value = None
            elif isinstance(value, pd.Timestamp):
                value = value.to_pydatetime()
            elif hasattr(value, "item"):  # numpy number -> plain Python number
                value = value.item()
            cell = ws.cell(row=r, column=start_col + offset, value=value)
            cell.border = BORDER
            number_format = formats.get(df.columns[offset])
            if number_format:
                cell.number_format = number_format
    return start_row + len(df)


def _display_length(value) -> int:
    if value is None:
        return 0
    if isinstance(value, (datetime, date)):
        return 10
    if isinstance(value, float):
        return len(f"{value:,.2f}") + 2
    return len(str(value))


def autosize(ws, min_width=10, max_width=45) -> None:
    """Set each column width to fit its longest value."""
    for cells in ws.columns:
        longest = max(_display_length(c.value) for c in cells)
        letter = get_column_letter(cells[0].column)
        ws.column_dimensions[letter].width = max(min_width, min(max_width, longest + 2))


def add_chart(ws, kind, title, rows, value_col, anchor, y_title) -> None:
    """Add a bar or line chart. Categories are in column A, values in `value_col`."""
    chart = BarChart() if kind == "bar" else LineChart()
    if kind == "bar":
        chart.type = "col"
    chart.title = title
    chart.y_axis.title = y_title
    chart.y_axis.number_format = INTEGER
    chart.x_axis.delete = False  # openpyxl 3.1 charts can hide axes in recent Excel otherwise
    chart.y_axis.delete = False
    chart.legend = None
    chart.height, chart.width = 8.5, 17

    data = Reference(ws, min_col=value_col, min_row=1, max_row=rows + 1)
    categories = Reference(ws, min_col=1, min_row=2, max_row=rows + 1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(categories)
    if kind == "line":
        chart.series[0].marker.symbol = "circle"
        chart.series[0].smooth = False
    ws.add_chart(chart, anchor)


def fit_to_page_width(ws) -> None:
    """Print/PDF friendly: landscape and one page wide."""
    ws.page_setup.orientation = "landscape"
    ws.sheet_properties.pageSetUpPr = PageSetupProperties(fitToPage=True)
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0


def write_summary_sheet(ws, title, kpis, stats, file_rows, skipped, money_fmt) -> None:
    ws["A1"] = title
    ws["A1"].font = TITLE_FONT
    ws["A2"] = f"Generated {datetime.now():%Y-%m-%d %H:%M}"
    ws["A2"].font = Font(italic=True, color="7F7F7F")

    def section(row: int, heading: str) -> int:
        ws.cell(row=row, column=1, value=heading).font = SECTION_FONT
        return row + 1

    def line(row: int, label, value, number_format=None, detail=None) -> int:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        cell = ws.cell(row=row, column=2, value=value)
        cell.alignment = Alignment(horizontal="left")
        if number_format:
            cell.number_format = number_format
        if detail:
            ws.cell(row=row, column=3, value=detail)
        return row + 1

    formats = {"money": money_fmt, "int": INTEGER, None: None}

    row = section(4, "Key figures")
    for label, value, fmt, detail in kpis:
        row = line(row, label, value.item() if hasattr(value, "item") else value, formats[fmt], detail)

    row = section(row + 1, "Data cleaning log")
    row = line(row, "Files processed", len(file_rows), INTEGER)
    row = line(row, "Files skipped", ", ".join(skipped) if skipped else "None")
    row = line(row, "Rows read", stats["rows_read"], INTEGER)
    row = line(row, "Rows removed (missing or invalid data)", stats["rows_invalid"], INTEGER)
    row = line(row, "Rows removed (duplicate orders)", stats["duplicates_removed"], INTEGER)
    row = line(row, "Rows in report", stats["rows_clean"], INTEGER)

    row = section(row + 1, "Source files")
    write_table(ws, file_rows, start_row=row, formats={"rows_read": INTEGER})

    ws.column_dimensions["A"].width = 38
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 36


def write_report(path: Path, df, tables, kpis, stats, file_rows, skipped, title, currency) -> None:
    money_fmt = f'"{currency}"#,##0.00'
    formats = {
        "revenue": money_fmt, "unit_price": money_fmt, "units": INTEGER, "orders": INTEGER,
        "share": PERCENT, "growth": PERCENT, "date": DATE_FORMAT,
    }

    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"
    write_summary_sheet(summary, title, kpis, stats, file_rows, skipped, money_fmt)

    sheets = [
        ("By Product", tables["product"], "bar", "Revenue by product"),
        ("By Region", tables["region"], "bar", "Revenue by region"),
        ("Monthly Trend", tables["month"], "line", "Revenue by month"),
    ]
    for name, table, kind, chart_title in sheets:
        ws = wb.create_sheet(name)
        last_row = write_table(ws, table, formats=formats)
        autosize(ws)
        ws.freeze_panes = "A2"
        revenue_col = list(table.columns).index("revenue") + 1
        add_chart(ws, kind, chart_title, last_row - 1, revenue_col, "G2", f"Revenue ({currency})")

    data_sheet = wb.create_sheet("Clean Data")
    write_table(data_sheet, df[DATA_COLUMNS], formats=formats)
    autosize(data_sheet)
    data_sheet.freeze_panes = "A2"
    data_sheet.auto_filter.ref = data_sheet.dimensions

    for sheet in wb.worksheets:
        fit_to_page_width(sheet)

    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ---------------------------------------------------------------------------
# 6. Command line interface
# ---------------------------------------------------------------------------


def parse_args(argv=None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Combine messy sales CSV/Excel files into one clean Excel report."
    )
    parser.add_argument("-i", "--input", default="sample_data", help="folder with CSV/Excel files (default: sample_data)")
    parser.add_argument("-o", "--output", default="output/sales_report.xlsx", help="report file to create")
    parser.add_argument("--title", default="Sales Report", help="title shown on the Summary sheet")
    parser.add_argument("--currency", default="$", help="currency symbol used in the report (default: $)")
    return parser.parse_args(argv)


def main(argv=None) -> int:
    args = parse_args(argv)
    folder = Path(args.input)
    if not folder.is_dir():
        print(f"Error: input folder '{folder}' does not exist.", file=sys.stderr)
        return 1

    files = find_input_files(folder)
    if not files:
        print(f"Error: no CSV or Excel files found in '{folder}'.", file=sys.stderr)
        return 1

    print(f"Reading {len(files)} file(s) from '{folder}'")
    frames, counts, skipped = [], [], []
    for path in files:
        try:
            frame = standardize_columns(read_file(path), path.name)
        except Exception as error:  # a bad file should not stop the whole run
            print(f"  ! Skipping {path.name}: {error}", file=sys.stderr)
            skipped.append(path.name)
            continue
        print(f"  {path.name}: {len(frame)} rows")
        frames.append(frame)
        counts.append({"source_file": path.name, "rows_read": len(frame)})

    if not frames:
        print("Error: none of the files could be used.", file=sys.stderr)
        return 1

    df, stats = clean(pd.concat(frames, ignore_index=True))
    if df.empty:
        print("Error: no valid rows left after cleaning.", file=sys.stderr)
        return 1

    tables, kpis = build_summary(df, args.currency)
    output = Path(args.output)
    write_report(output, df, tables, kpis, stats, pd.DataFrame(counts), skipped, args.title, args.currency)

    print(
        f"Cleaning: {stats['duplicates_removed']} duplicate order(s) removed, "
        f"{stats['rows_invalid']} row(s) dropped for missing/invalid data"
    )
    print(f"Report saved to {output}  ({stats['rows_clean']} orders, "
          f"total revenue {args.currency}{df['revenue'].sum():,.2f})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
