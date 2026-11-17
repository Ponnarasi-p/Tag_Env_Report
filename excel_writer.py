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

# Colours follow the original downloaded report.
TITLE_FILL = PatternFill("solid", fgColor="92D050")     # report title (green)
SECTION_FILL = PatternFill("solid", fgColor="9BC2E6")   # section title, e.g. "Master Roll" (blue)
HEADER_FILL = PatternFill("solid", fgColor="FFE676")    # column headings (yellow)
HEADER_FONT = Font(bold=True, color="000000", size=11)
BAND_FILL = PatternFill("solid", fgColor="F7F9FC")
SUMMARY_FILL = PatternFill("solid", fgColor="DDEBF7")   # grouped item row
TOTAL_FILL = PatternFill("solid", fgColor="DBDBD0")     # Grand Total (grey)
THIN = Side(style="thin", color="A6A6A6")
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

    ncol = max([len(sec.headers) for sec in s.sections] + [1])
    widths = [0] * ncol
    r = 1
    for i, t in enumerate(s.title_rows):
        c = ws.cell(row=r, column=1, value=t)
        if ncol > 1:
            ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=ncol)
        c.font = Font(bold=True, size=14 if i == 0 else 11)
        c.alignment = Alignment(horizontal="center", vertical="center")
        if i == 0:
            for ci in range(1, ncol + 1):
                ws.cell(row=r, column=ci).fill = TITLE_FILL
        ws.row_dimensions[r].height = 26 if i == 0 else 18
        r += 1

    single = len(s.sections) == 1
    first_header_row = None
    for n, sec in enumerate(s.sections):
        if not sec.headers:
            continue
        if r > 1:
            r += 1  # blank line between the title and each table
        sw = len(sec.headers)
        if sec.title:
            c = ws.cell(row=r, column=1, value=sec.title)
            if sw > 1:
                ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=sw)
            for ci in range(1, sw + 1):
                cell = ws.cell(row=r, column=ci)
                cell.fill = SECTION_FILL
                cell.border = BORDER
            c.font = Font(bold=True, size=12)
            c.alignment = Alignment(horizontal="center", vertical="center")
            ws.row_dimensions[r].height = 22
            r += 1
        header_row = r
        first_header_row = first_header_row or header_row
        r = _write_table(ws, sec, r, widths)
        if single and sec.rows:
            ws.auto_filter.ref = f"A{header_row}:{get_column_letter(sw)}{header_row + len(sec.rows)}"

    for ci, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(ci)].width = max(8, min(w or 10, 60))

    if first_header_row:
        if single:
            item_col = s.sections[0].item_col
            freeze_col = item_col + 2 if 0 <= item_col <= 2 else 1
            ws.freeze_panes = ws.cell(row=first_header_row + 1, column=freeze_col)
            ws.print_title_rows = f"{first_header_row}:{first_header_row}"
        elif s.title_rows:
            ws.freeze_panes = ws.cell(row=len(s.title_rows) + 1, column=1)


def _write_table(ws, sec, r: int, widths: list) -> int:
    """Writes one heading row plus its data rows starting at row r; returns the next free row."""
    ncol = len(sec.headers)
    for ci, h in enumerate(sec.headers, start=1):
        c = ws.cell(row=r, column=ci, value=h)
        c.font = HEADER_FONT
        c.fill = HEADER_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
        widths[ci - 1] = max(widths[ci - 1], min(_text_len(h), 22) + 4)
    ws.row_dimensions[r].height = 32
    r += 1

    band = 0
    for row in sec.rows:
        if row.type != "detail":
            band += 1
        for ci, v in enumerate(row.cells[:ncol]):
            display = v
            if row.type == "detail" and ci == sec.item_col:
                display = "    " + ("" if v is None else str(v))
            if row.type == "summary" and ci == sec.item_col:
                display = f"▸ {'' if v is None else v}  ({row.count})"
            c = ws.cell(row=r, column=ci + 1, value=display)
            c.border = BORDER
            num, dash, dt = is_num(v), v == "-", isinstance(v, (datetime, date))
            if num:
                c.number_format = "#,##0"
            elif dt:
                c.number_format = "dd-mmm-yyyy"
            if num or (dash and sec.numeric_cols[ci]):
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
        if row.type == "summary" and sec.item_col >= 0:
            ws.cell(row=r, column=sec.item_col + 1).comment = Comment(
                f"Repeated item: {row.count} records. Click the [+] in the left margin to expand.", "Report Tool")
        if row.type == "detail":
            ws.row_dimensions[r].outline_level = 1
            ws.row_dimensions[r].hidden = True
        ws.row_dimensions[r].height = 18 if row.type == "detail" else 20
        r += 1
    return r
