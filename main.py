"""KYC Compliance Analytics pipeline:  Excel -> pandas cleaning -> SQLite/SQL -> Excel report."""
import argparse
from datetime import date
from pathlib import Path

import pandas as pd

from src.clean import clean_cases
from src.database import load_cases, run_query
from src.generate_data import generate_cases, save_raw_excel
from src.report import build_report


def run_pipeline(input_path=None, output_dir="output", as_of=None, sla_days=5, rows=500):
    as_of = as_of or date.today()
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)

    # 1. Extract: read a real export, or generate a messy sample one
    if input_path:
        raw = pd.read_excel(input_path)
    else:
        raw = generate_cases(rows, as_of=as_of)
        save_raw_excel(raw, out / "raw_kyc_export.xlsx")
    print(f"[extract]   {len(raw)} raw rows")

    # 2. Transform: clean and validate
    clean, issues = clean_cases(raw)
    print(f"[clean]     {len(clean)} rows kept, {len(issues)} removed")
    if len(issues):
        for reason, count in issues["reason"].value_counts().items():
            print(f"            - {count} x {reason}")

    # 3. Load + analyse with SQL
    conn = load_cases(clean, out / "kyc.sqlite")
    params = {"as_of": as_of.isoformat(), "sla_days": sla_days}
    results = {
        "SQL_Turnaround_by_Doc": run_query(conn, "turnaround_by_document"),
        "SQL_Analyst_Performance": run_query(conn, "analyst_performance"),
        "SQL_Monthly_Trend": run_query(conn, "monthly_trend"),
        "Overdue_Cases": run_query(conn, "overdue_cases", **params),
    }
    print(f"[sql]       {len(results['Overdue_Cases'])} overdue pending cases (SLA {sla_days} days)")
    print(run_query(conn, "status_summary").to_string(index=False))
    conn.close()

    # 4. Report
    report = build_report(clean, issues, results, as_of, sla_days, out / "kyc_compliance_report.xlsx")
    print(f"[report]    {report}")
    return report


if __name__ == "__main__":
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument("--input", help="Raw KYC Excel export (omit to generate sample data)")
    p.add_argument("--output-dir", default="output")
    p.add_argument("--as-of", type=date.fromisoformat, default=date.today(), help="YYYY-MM-DD")
    p.add_argument("--sla-days", type=int, default=5)
    p.add_argument("--rows", type=int, default=500, help="Rows of sample data to generate")
    a = p.parse_args()
    run_pipeline(a.input, a.output_dir, a.as_of, a.sla_days, a.rows)
