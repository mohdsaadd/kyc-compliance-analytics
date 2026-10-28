# KYC Compliance Analytics Pipeline (Excel + SQL + Python)

An end-to-end data pipeline that takes a messy KYC case export from Excel, cleans and validates it with
**pandas**, analyses it with **SQL** (SQLite), and produces a formatted **Excel report** with live formulas,
a chart and an overdue-case list.

It automates the kind of reporting I did manually as a Process Analyst (KYC verification, Excel trackers,
compliance KPIs): cleaning exports, tracking turnaround against an SLA and flagging overdue cases.

## What it does

```
Raw Excel export ──▶ pandas cleaning ──▶ SQLite + SQL analysis ──▶ Excel report
 (messy data)        (validate + log)     (5 named queries)        (formulas + chart)
```

1. **Extract:** reads a real `.xlsx` export, or generates a realistic messy sample (duplicates, inconsistent
   status wording, stray whitespace, invalid dates, completion dates before submission).
2. **Clean:** trims and standardises text, maps status variants (`Done` to `Approved`, `In Progress` to
   `Pending`), removes duplicates, parses dates, and applies logical checks. Every removed row is written
   to a `Data_Issues` sheet with the reason, so nothing disappears silently.
3. **SQL analysis:** loads data into SQLite and runs named queries: status summary, turnaround by document
   type, analyst performance, monthly trend, and overdue pending cases (SLA-based, sorted by risk).
4. **Report:** builds `kyc_compliance_report.xlsx`:
   - **Summary:** KPIs and a risk-by-status table built with `COUNTIF(S)`, `AVERAGE` and `IFERROR`
     formulas, plus a stacked chart. The as-of date and SLA are input cells, so changing them
     recalculates everything in Excel.
   - **Cases:** the cleaned data with a turnaround formula per row.
   - **SQL_* sheets:** the SQL query results.
   - **Overdue_Cases** and **Data_Issues**.

A recalculated example is in [`docs/sample_report.xlsx`](docs/sample_report.xlsx).

## Tech stack

Python 3, pandas, openpyxl, SQLite (SQL), pytest

## Getting started

```bash
git clone <your-repo-url>
cd kyc-compliance-analytics
python -m venv venv && source venv/bin/activate   # Windows: venv\Scripts\activate
pip install -r requirements.txt

python main.py                                    # sample data, SLA 5 days
python main.py --input my_export.xlsx --sla-days 7 --as-of 2026-10-27
```

Output goes to `output/`: `kyc_compliance_report.xlsx`, `kyc.sqlite`, and (for sample runs) `raw_kyc_export.xlsx`.
Your own export needs these columns: `case_id, customer_name, document_type, region, risk_level, status,
analyst, submitted_date, completed_date`.

### Example output

```
[extract]   515 raw rows
[clean]     481 rows kept, 34 removed
            - 15 x duplicate case_id
            - 10 x invalid or missing submitted_date
            - 9 x completed_date earlier than submitted_date
[sql]       35 overdue pending cases (SLA 5 days)
```

## Project structure

```
kyc-compliance-analytics/
├── main.py                # pipeline entry point (CLI)
├── sql/queries.sql        # named SQL queries
├── src/
│   ├── generate_data.py   # messy sample data
│   ├── clean.py           # cleaning and validation
│   ├── database.py        # SQLite schema, loading, query runner
│   └── report.py          # Excel report builder
├── tests/test_pipeline.py # 13 tests
└── docs/sample_report.xlsx
```

## Tests

```bash
python -m pytest -q
```

Tests cover the cleaning rules, duplicate handling, the logging of invalid rows, SQL results (turnaround
averages, SLA logic, risk ordering) and an end-to-end run that checks the report contains live formulas.

I also checked that the Excel formula results match the SQL results for the same data (total, pending,
approved, rejected and overdue counts).

## Possible improvements

- Email or Slack alerts for overdue cases
- A Power BI dashboard connected to `kyc.sqlite`
- Scheduling the pipeline daily (cron / Task Scheduler)
- Using VBA to add a "refresh" button inside the workbook
