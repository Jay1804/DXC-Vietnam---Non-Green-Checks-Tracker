"""
Builds the polished "DXC Vietnam - Final Report.xlsx" from the Amber/Red
selected-columns workbook.

CHECK_CLOSURE_DATE filter (dynamic, evaluated against today's date each run):
  - Normally: yesterday's date only.
  - If today is Monday: Friday, Saturday, and Sunday (to cover the weekend).

If no rows match the date filter, the report keeps the header only.

Usage:
    python build_final_report.py
"""

import sys
from datetime import datetime, timedelta

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

SOURCE_FILE = "DXC Vietnam - Amber_Red - Selected Columns.xlsx"
OUTPUT_FILE = "DXC Vietnam - Final Report.xlsx"
DATE_COLUMN = "CHECK_CLOSURE_DATE"

HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11, name="Calibri")
TITLE_FONT = Font(color="1F3864", bold=True, size=16, name="Calibri")
SUBTITLE_FONT = Font(color="595959", italic=True, size=10, name="Calibri")
BODY_FONT = Font(size=10, name="Calibri")
BAND_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
SEVERITY_FILLS = {
    "Red": PatternFill(start_color="F8CBAD", end_color="F8CBAD", fill_type="solid"),
    "Amber": PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid"),
}
THIN_BORDER = Border(
    left=Side(style="thin", color="D9D9D9"),
    right=Side(style="thin", color="D9D9D9"),
    top=Side(style="thin", color="D9D9D9"),
    bottom=Side(style="thin", color="D9D9D9"),
)


def get_target_dates():
    today = datetime.now()
    if today.weekday() == 0:  # Monday -> Fri, Sat, Sun
        return {(today - timedelta(days=d)).strftime("%Y-%m-%d") for d in (1, 2, 3)}
    return {(today - timedelta(days=1)).strftime("%Y-%m-%d")}


def load_source_rows():
    wb = openpyxl.load_workbook(SOURCE_FILE)
    ws = wb.active
    headers = [c.value for c in ws[1]]
    rows = [list(row) for row in ws.iter_rows(min_row=2, values_only=True)]
    return headers, rows


def filter_by_closure_date(headers, rows, target_dates):
    idx = headers.index(DATE_COLUMN)
    kept = []
    for row in rows:
        val = row[idx]
        date_part = str(val).split(" ")[0] if val else ""
        if date_part in target_dates:
            kept.append(row)
    return kept


def autosize_columns(ws, headers, start_row):
    for col_idx, header in enumerate(headers, start=1):
        max_len = len(str(header))
        for row in ws.iter_rows(min_row=start_row, min_col=col_idx, max_col=col_idx):
            for cell in row:
                if cell.value is not None:
                    max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 40)


def build_report(headers, rows, target_dates):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Final Report"

    n_cols = len(headers)
    last_col_letter = get_column_letter(n_cols)

    # Title block
    ws.merge_cells(f"A1:{last_col_letter}1")
    ws["A1"] = "DXC Vietnam - Non Green Checks Report"
    ws["A1"].font = TITLE_FONT
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 26

    date_label = " / ".join(sorted(target_dates)) if target_dates else "N/A"
    ws.merge_cells(f"A2:{last_col_letter}2")
    ws["A2"] = f"Check Closure Date: {date_label}    |    Severity: Amber, Red    |    Generated: {datetime.now().strftime('%Y-%m-%d %H:%M')}"
    ws["A2"].font = SUBTITLE_FONT
    ws.row_dimensions[2].height = 16

    ws.row_dimensions[3].height = 6  # spacer

    header_row = 4
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=header_row, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
    ws.row_dimensions[header_row].height = 22

    severity_idx = headers.index("check_severity") if "check_severity" in headers else None

    for r_offset, row in enumerate(rows):
        r = header_row + 1 + r_offset
        band = BAND_FILL if r_offset % 2 == 1 else None
        for col_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r, column=col_idx, value=value)
            cell.font = BODY_FONT
            cell.border = THIN_BORDER
            cell.alignment = Alignment(horizontal="left", vertical="center")
            if band:
                cell.fill = band
        if severity_idx is not None:
            sev_val = row[severity_idx]
            sev_fill = SEVERITY_FILLS.get(sev_val)
            if sev_fill:
                ws.cell(row=r, column=severity_idx + 1).fill = sev_fill

    last_row = header_row + max(len(rows), 1)
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    ws.auto_filter.ref = f"A{header_row}:{last_col_letter}{header_row + len(rows)}" if rows else f"A{header_row}:{last_col_letter}{header_row}"

    autosize_columns(ws, headers, start_row=header_row)

    if not rows:
        ws.merge_cells(f"A{header_row + 1}:{last_col_letter}{header_row + 1}")
        note = ws.cell(row=header_row + 1, column=1, value="No records found for the selected date(s).")
        note.font = Font(italic=True, color="808080", size=10)
        note.alignment = Alignment(horizontal="center")

    wb.save(OUTPUT_FILE)
    return last_row


def main():
    target_dates = get_target_dates()
    headers, rows = load_source_rows()
    filtered_rows = filter_by_closure_date(headers, rows, target_dates)
    build_report(headers, filtered_rows, target_dates)

    print(f"Target closure date(s): {sorted(target_dates)}")
    print(f"Rows matched: {len(filtered_rows)}")
    print(f"Saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        sys.exit(1)
