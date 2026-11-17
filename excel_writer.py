"""Writes the processed sheets to a formatted .xlsx (openpyxl)."""
from __future__ import annotations

import io
import re
from datetime import date, datetime

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.properties import Outline

from report_processor import Sheet, is_num

TITLE = "1F3864"
HEADER_FILL = PatternFill("solid", fgColor="1F4E79")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
BAND_FILL = PatternFill("solid", fgColor="F5F8FC")
SUMMARY_FILL = PatternFill("solid", fgColor="DDEBF7")
TOTAL_FILL = PatternFill("solid", fgColor="FCE4D6")
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(top=THIN, left=THIN, bottom=THIN, right=THIN)


def _text_len(v) -> int:
    if v is None:
        return 0
    if isinstance(v, (datetime, date)):
        return 11
    if is_num(v):
        return len(f"{v:,}")
    return max((len(line) for line in str(v).split("\n")), default=0)


def _safe_name(name: str, used: set) -> str:
    n = re.sub(r"[\\/?*\[\]:]", " ", str(name or "Sheet"))[:31].strip() or "Sheet"
    base, i = n, 2
    while n.lower() in used:
        n = f"{base[:28]} {i}"
        i += 1
    used.add(n.lower())
    return n


def build_workbook(sheets: list[Sheet]) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    used: set = set()
    for s in sheets:
        _write_sheet(wb, s, _safe_name(s.name, used))
    if not wb.worksheets:
        wb.create_sheet("Sheet1")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _write_sheet(wb: Workbook, s: Sheet, sheet_name: str) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.sheet_properties.outlinePr = Outline(summaryBelow=False, summaryRight=False)
    ws.sheet_view.showGridLines = False
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    ncol = max(len(s.headers), 1)
    r = 1
    for i, t in enumerate(s.title_rows):
        c = ws.cell(row=r, column=1, value=t)
        if ncol > 1:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        c.font = Font(bold=True, size=14 if i == 0 else 11, color=TITLE)
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[r].height = 24 if i == 0 else 18
        r += 1
    if s.title_rows:
        r += 1

    if not s.headers:
        return

    header_row = r
    for ci, h in enumerate(s.headers, start=1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[r].height = 32
    r += 1

    widths = [min(_text_len(h), 22) + 4 for h in s.headers]
    band = 0
    for row in s.rows:
        if row.type != "detail":
            band += 1
        for ci, v in enumerate(row.cells[:ncol]):
            display = v
            if row.type == "detail" and ci == s.item_col:
                display = "    " + ("" if v is None else str(v))
            if row.type == "summary" and ci == s.item_col:
                display = f"▸ {'' if v is None else v}  ({row.count})"
            c = ws.cell(row=r, column=ci + 1, value=display)
            c.border = BORDER
            num, dash, dt = is_num(v), v == "-", isinstance(v, (datetime, date))
            if num:
                c.number_format = "#,##0"
            elif dt:
                c.number_format = "dd-mmm-yyyy"
            if num or (dash and s.numeric_cols[ci]):
                horiz = "right"
            elif dash or dt:
                horiz = "center"
            else:
                horiz = "left"
            c.alignment = Alignment(horizontal=horiz, vertical="center",
                                    indent=0 if (num or dash or dt) else 1)
            if row.type == "summary":
                c.fill, c.font = SUMMARY_FILL, Font(bold=True)
            elif row.type == "detail":
                c.font = Font(color="404040", size=10)
            elif row.type == "total":
                c.fill, c.font = TOTAL_FILL, Font(bold=True)
            elif band % 2 == 0:
                c.fill = BAND_FILL
            widths[ci] = max(widths[ci], _text_len(display) + 3)
        if row.type == "summary" and s.item_col >= 0:
            ws.cell(row=r, column=s.item_col + 1).comment = Comment(
                f"Repeated item: {row.count} records. Click the [+] in the left margin to expand.", "Report Tool")
        if row.type == "detail":
            ws.row_dimensions[r].outline_level = 1
            ws.row_dimensions[r].hidden = True
        ws.row_dimensions[r].height = 18 if row.type == "detail" else 20
        r += 1

    for ci, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = max(8, min(w, 60))

    freeze_col = s.item_col + 2 if 0 <= s.item_col <= 2 else 1
    ws.freeze_panes = ws.cell(row=header_row + 1, column=freeze_col)
    if s.rows:
        ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ncol)}{header_row + len(s.rows)}"
    ws.print_title_rows = f"{header_row}:{header_row}"
