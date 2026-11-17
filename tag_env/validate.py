"""Structure and formatting checks of a generated report against the template."""
from __future__ import annotations

import io
from collections import Counter

import openpyxl
import pandas as pd


def _load(src):
    from .writer import TEMPLATE_PATH
    if src is None:
        src = TEMPLATE_PATH
    if isinstance(src, (bytes, bytearray)):
        src = io.BytesIO(src)
    return openpyxl.load_workbook(src)


def _style_sig(c):
    return (c.font.name, c.font.sz, bool(c.font.b), c.fill.fill_type,
            c.fill.fgColor.rgb if c.fill.fill_type else None, c.number_format,
            c.border.left.style, c.border.right.style, c.border.top.style, c.border.bottom.style,
            c.alignment.horizontal)


def compare_with_template(generated: bytes, template_src, cfg: dict) -> pd.DataFrame:
    g, t = _load(generated), _load(template_src)
    checks = []

    def add(area, check, ok, detail=""):
        checks.append({"Area": area, "Check": check, "Result": "PASS" if ok else "FAIL", "Detail": detail})

    add("Workbook", "Sheet names and order", g.sheetnames == t.sheetnames,
        f"{g.sheetnames}" if g.sheetnames != t.sheetnames else "")
    vis_g = [ws.sheet_state for ws in g.worksheets]
    vis_t = [ws.sheet_state for ws in t.worksheets]
    add("Workbook", "Hidden / visible sheets", vis_g == vis_t)
    add("Workbook", "No links to other workbooks", not g._external_links)

    header_rows = {cfg["main_sheet"]["name"]: cfg["main_sheet"]["header_rows"],
                   cfg["stock_sheets"]["names"]["short"]: cfg["stock_sheets"]["header_rows"],
                   cfg["stock_sheets"]["names"]["full"]: cfg["stock_sheets"]["header_rows"]}
    for name in t.sheetnames:
        if name not in g.sheetnames:
            continue
        gw, tw = g[name], t[name]
        widths_ok = all(abs((gw.column_dimensions[k].width or 0) - (v.width or 0)) < 0.01
                        for k, v in tw.column_dimensions.items())
        add(name, "Column widths", widths_ok)
        h = header_rows.get(name)
        if h is None:  # export sheet: whole layout is fixed
            same_labels = all(gw.cell(r, 1).value == tw.cell(r, 1).value for r in range(1, tw.max_row + 1))
            add(name, "Row labels and layout unchanged", same_labels)
            add(name, "Merged cells unchanged",
                sorted(map(str, gw.merged_cells.ranges)) == sorted(map(str, tw.merged_cells.ranges)))
            styles = all(_style_sig(gw.cell(r, c)) == _style_sig(tw.cell(r, c))
                         for r in range(1, tw.max_row + 1) for c in range(1, tw.max_column + 1))
            add(name, "Cell formatting unchanged", styles)
            continue
        hdr_vals = all(gw.cell(r, c).value == tw.cell(r, c).value
                       for r in range(2, h + 1) for c in range(1, tw.max_column + 1)
                       if not (name == cfg["main_sheet"]["name"] and r == cfg["main_sheet"]["conversion_row"]))
        add(name, "Header titles and column names", hdr_vals)
        hdr_sty = all(_style_sig(gw.cell(r, c)) == _style_sig(tw.cell(r, c))
                      for r in range(1, h + 1) for c in range(1, tw.max_column + 1))
        add(name, "Header formatting (fonts, fills, borders, number formats)", hdr_sty)
        tm = sorted(str(m) for m in tw.merged_cells.ranges if m.min_row <= h)
        gm = sorted(str(m) for m in gw.merged_cells.ranges if m.min_row <= h)
        add(name, "Header merged cells", tm == gm)
        # data-cell formatting: each generated cell uses a format that the same
        # column already has in the template (cells inside merged blocks skipped)
        ncol = tw.max_column
        t_col = {c: set() for c in range(1, ncol + 1)}
        r = h + 1
        while tw.cell(r, 1).value not in (None, "") and str(tw.cell(r, 1).value).upper() not in ("TOTAL", "GRAND TOTAL"):
            for c in range(1, ncol + 1):
                t_col[c].add(_style_sig(tw.cell(r, c)))
            r += 1
        in_merge = set()
        for mr in gw.merged_cells.ranges:
            if mr.min_row > h:
                in_merge |= {(rr, cc) for rr in range(mr.min_row, mr.max_row + 1)
                             for cc in range(mr.min_col, mr.max_col + 1)}
        bad_rows = []
        r = h + 1
        n = 0
        while gw.cell(r, 1).value not in (None, "") and str(gw.cell(r, 1).value).upper() not in ("TOTAL", "GRAND TOTAL"):
            if any((r, c) not in in_merge and _style_sig(gw.cell(r, c)) not in t_col[c] for c in range(1, ncol + 1)):
                bad_rows.append(r)
            n += 1
            r += 1
        add(name, "Data cells use the template's formatting", not bad_rows,
            f"{n} rows" + (f"; rows {bad_rows[:8]} differ" if bad_rows else ""))
        add(name, "Row height", all(gw.row_dimensions[x].height == tw.row_dimensions[h + 1].height
                                    for x in range(h + 1, h + 1 + n)))
    return pd.DataFrame(checks)
