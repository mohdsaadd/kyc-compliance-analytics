"""Build the Excel report. The Summary sheet uses live formulas over the Cases sheet,
so changing the as-of date or SLA recalculates every KPI inside Excel."""
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

FONT = "Arial"
HEADER_FILL = PatternFill("solid", fgColor="1F2A44")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
RISKS = ["Low", "Medium", "High"]
STATUSES = ["Pending", "Approved", "Rejected"]


def _style_header(ws, row=1):
    for cell in ws[row]:
        if cell.value is not None:
            cell.font = Font(name=FONT, bold=True, color="FFFFFF")
            cell.fill = HEADER_FILL
            cell.alignment = Alignment(horizontal="center")


def _autosize(ws, minimum=10, maximum=38):
    for col in ws.columns:
        width = max((len(str(c.value)) for c in col if c.value is not None), default=0)
        ws.column_dimensions[get_column_letter(col[0].column)].width = \
            min(max(width + 2, minimum), maximum)


def _write_df(ws, df, date_cols=()):
    ws.append(list(df.columns))
    for rec in df.itertuples(index=False):
        row = []
        for col, val in zip(df.columns, rec):
            if hasattr(val, "to_pydatetime"):
                val = val.to_pydatetime().date()
            elif val != val:  # NaN / NaT
                val = None
            row.append(val)
        ws.append(row)
    for col_name in date_cols:
        idx = list(df.columns).index(col_name) + 1
        for r in range(2, ws.max_row + 1):
            ws.cell(r, idx).number_format = "yyyy-mm-dd"
    _style_header(ws)
    _autosize(ws)
    ws.freeze_panes = "A2"
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name=FONT, size=10)


def build_report(clean_df, issues_df, sql_results, as_of, sla_days, path):
    """sql_results: dict {sheet_name: DataFrame} written as values (outputs of SQL queries)."""
    wb = Workbook()
    summary = wb.active
    summary.title = "Summary"

    # ---- Cases sheet (cleaned data + turnaround formula) ----
    cases = wb.create_sheet("Cases")
    _write_df(cases, clean_df, date_cols=("submitted_date", "completed_date"))
    n_last = max(len(clean_df), 1) + 1
    cases["J1"] = "turnaround_days"
    cases["J1"].font = Font(name=FONT, bold=True, color="FFFFFF")
    cases["J1"].fill = HEADER_FILL
    cases["J1"].alignment = Alignment(horizontal="center")
    for r in range(2, n_last + 1):
        cases[f"J{r}"] = f'=IF(I{r}="","",I{r}-H{r})'
        cases[f"J{r}"].font = Font(name=FONT, size=10)
    cases.column_dimensions["J"].width = 17

    # ---- Summary sheet ----
    f = lambda **kw: Font(name=FONT, size=10, **kw)
    summary["A1"] = "KYC Compliance Report"
    summary["A1"].font = Font(name=FONT, size=16, bold=True)

    summary["A3"], summary["A4"], summary["A5"] = "Assumptions (edit the yellow cells)", "As-of date", "SLA (days)"
    summary["B4"], summary["B5"] = as_of, sla_days
    summary["B4"].number_format = "yyyy-mm-dd"
    for ref in ("B4", "B5"):
        summary[ref].fill = INPUT_FILL
        summary[ref].font = f(color="0000FF")  # blue = hardcoded input

    rng = lambda col: f"Cases!${col}$2:${col}${n_last}"
    summary["A7"] = "Key metrics"
    metrics = [
        ("Total cases", f"=COUNTA({rng('A')})", "0"),
        ("Pending", f'=COUNTIF({rng("F")},"Pending")', "0"),
        ("Approved", f'=COUNTIF({rng("F")},"Approved")', "0"),
        ("Rejected", f'=COUNTIF({rng("F")},"Rejected")', "0"),
        ("Approval rate (of decided cases)", "=IFERROR(B10/(B10+B11),0)", "0.0%"),
        ("Average turnaround (days)", f"=IFERROR(AVERAGE({rng('J')}),0)", "0.0"),
        ("Overdue pending cases", f'=COUNTIFS({rng("F")},"Pending",{rng("H")},"<"&($B$4-$B$5))', "0"),
    ]
    for i, (label, formula, fmt) in enumerate(metrics, start=8):
        summary[f"A{i}"], summary[f"B{i}"] = label, formula
        summary[f"B{i}"].number_format = fmt

    summary["A16"] = "Cases by risk level and status"
    for j, h in enumerate(["Risk level"] + STATUSES + ["Total"], start=1):
        summary.cell(17, j, h)
    for i, risk in enumerate(RISKS, start=18):
        summary.cell(i, 1, risk)
        for j, status in enumerate(STATUSES, start=2):
            col = get_column_letter(j)
            summary.cell(i, j, f'=COUNTIFS({rng("E")},$A{i},{rng("F")},{col}$17)')
        summary.cell(i, 5, f"=SUM(B{i}:D{i})")
    summary.cell(21, 1, "Total")
    for j in range(2, 6):
        col = get_column_letter(j)
        summary.cell(21, j, f"=SUM({col}18:{col}20)")

    notes = [
        "Notes",
        "Source: cleaned KYC case data (Cases sheet); SQL-based analysis is on the SQL_* sheets.",
        "Overdue = Pending case submitted more than SLA days before the as-of date.",
        "As-of date and SLA are inputs (blue text, yellow fill); every metric above recalculates from them.",
        "Data_Issues lists the raw rows removed during cleaning and the reason.",
    ]
    for i, text in enumerate(notes, start=23):
        summary[f"A{i}"] = text

    for row in summary.iter_rows():
        for c in row:
            if c.font.name != FONT or (c.font.size == 11 and not c.font.bold):
                c.font = f(color=c.font.color.rgb if c.font.color and isinstance(c.font.color.rgb, str) else None)
    for ref in ("A3", "A7", "A16", "A23"):
        summary[ref].font = Font(name=FONT, size=11, bold=True)
    summary["A1"].font = Font(name=FONT, size=16, bold=True)
    for j in range(1, 6):
        c = summary.cell(17, j)
        c.font = Font(name=FONT, bold=True, color="FFFFFF")
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center")
    for j in range(1, 6):
        summary.cell(21, j).font = Font(name=FONT, size=10, bold=True)
    summary.column_dimensions["A"].width = 38
    for col in "BCDE":
        summary.column_dimensions[col].width = 13

    # highlight overdue count when it is above zero
    summary.conditional_formatting.add(
        "B14", CellIsRule(operator="greaterThan", formula=["0"],
                          fill=PatternFill("solid", bgColor="F8CBAD")))

    chart = BarChart()
    chart.type, chart.grouping, chart.overlap = "col", "stacked", 100
    chart.title = "Cases by risk level and status"
    chart.y_axis.title, chart.x_axis.title = "Cases", "Risk level"
    chart.add_data(Reference(summary, min_col=2, max_col=4, min_row=17, max_row=20), titles_from_data=True)
    chart.set_categories(Reference(summary, min_col=1, min_row=18, max_row=20))
    chart.height, chart.width = 8, 14
    chart.x_axis.delete = chart.y_axis.delete = False
    summary.add_chart(chart, "G3")

    # ---- SQL result sheets, overdue list, data issues ----
    for name, df in sql_results.items():
        ws = wb.create_sheet(name[:31])
        _write_df(ws, df, date_cols=[c for c in df.columns if "date" in c])
    issues = wb.create_sheet("Data_Issues")
    _write_df(issues, issues_df)

    wb.save(path)
    return path
