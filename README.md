# Excel Report Automator

Turn a folder of messy sales spreadsheets into one clean, formatted Excel report with a single command.

## The problem

Sales data often arrives from different people and systems. One file calls the column `Qty`, another calls it `Quantity`. Dates are written as `2026-03-05` in one file and `05/03/2026` in another. Orders appear twice, and some cells are empty. Merging everything by hand in Excel is slow and easy to get wrong.

## What it does

- Reads every CSV and Excel file (`.csv`, `.xlsx`, `.xlsm`) in a folder. CSV files can be comma or semicolon separated.
- Maps different column names to one standard (`Qty`, `Quantity` and `Units` all become `quantity`).
- Cleans the data:
  - trims extra spaces and unifies spelling, so `usb-c hub` and `USB-C Hub` become one product
  - reads mixed date formats
  - removes duplicate orders
  - drops rows with missing or invalid data
- Builds an Excel report with:
  - **Summary**: key figures (revenue, orders, average order value, top product, top region, best month) and a cleaning log
  - **By Product**, **By Region**, **Monthly Trend**: tables with charts (the monthly table includes month-over-month growth)
  - **Clean Data**: every cleaned order, with filters and a frozen header row
- Keeps a cleaning log, so nothing disappears silently. You can always see how many rows were read, removed or kept.
- Skips a file that cannot be read (for example, one with a missing column), tells you why, and carries on with the rest.

An example of the result is in [`examples/sample_report.xlsx`](examples/sample_report.xlsx).

## Quick start

Tested with Python 3.13, pandas 3.0 and openpyxl 3.1.

```bash
pip install -r requirements.txt

# optional: create three messy demo files in sample_data/
python generate_sample_data.py

# build the report
python report_automator.py --input sample_data --output output/sales_report.xlsx
```

Output of the demo run:

```text
Reading 3 file(s) from 'sample_data'
  sales_february.csv: 45 rows
  sales_january.csv: 40 rows
  sales_march.xlsx: 52 rows
Cleaning: 2 duplicate order(s) removed, 2 row(s) dropped for missing/invalid data
Report saved to output/sales_report.xlsx  (133 orders, total revenue $42,073.30)
```

## Options

| Option | Description | Default |
| --- | --- | --- |
| `-i`, `--input` | Folder that contains the CSV/Excel files | `sample_data` |
| `-o`, `--output` | Report file to create | `output/sales_report.xlsx` |
| `--title` | Title shown on the Summary sheet | `Sales Report` |
| `--currency` | Currency symbol used in the report | `$` |

Example with your own data:

```bash
python report_automator.py -i "C:/clients/acme/sales" -o acme_report.xlsx --title "ACME Sales Q1" --currency "€"
```

## Expected input

Each file needs these columns. The tool recognises several common names for each one:

| Standard name | Also recognised as | Required |
| --- | --- | --- |
| `order_id` | Order ID, OrderID, ID, Order Number | yes |
| `date` | Date, Order Date | yes |
| `product` | Product, Product Name, Item | yes |
| `quantity` | Quantity, Qty, Units | yes |
| `unit_price` | Unit Price, Price | yes |
| `region` | Region, Area | no (becomes `Unknown`) |
| `sales_rep` | Sales Rep, Rep, Salesperson | no (becomes `Unknown`) |

Need to support another column name? Add it to `COLUMN_ALIASES` at the top of `report_automator.py`.

## How it works

1. **Load**: read each file and rename its columns to the standard names.
2. **Clean**: fix text and dates, drop invalid rows, remove duplicate orders, calculate `revenue = quantity × unit_price`.
3. **Analyse**: group the clean data by product, region and month.
4. **Write**: build the Excel workbook with `openpyxl` (styling, number formats, charts).

## Project structure

```text
excel-report-automator/
├── report_automator.py         # the tool
├── generate_sample_data.py     # creates messy demo files
├── test_report_automator.py    # unit tests
├── sample_data/                # demo input: two CSV files and one Excel file
├── examples/sample_report.xlsx # example output
├── requirements.txt
└── README.md
```

## Tests

```bash
python -m unittest -v
```

## Known limitations

- Dates that are not in ISO format (`YYYY-MM-DD`) are read day-first, so `05/03/2026` means 5 March. US-style month-first files would need a small change in `parse_dates`.
- Decimal commas (`29,90`) are not converted, so rows that use them are dropped and counted in the cleaning log.
- All amounts are assumed to be in one currency.

## Built with

Python, pandas, openpyxl
