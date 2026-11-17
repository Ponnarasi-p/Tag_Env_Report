"""Read the reference Tag Env Stock workbook and write a new report into a copy
of it, keeping every sheet, header, style, width and merge, and replacing only
the data rows, conversion factors, dates and totals.
"""
from __future__ import annotations

import copy
import io
import re
from collections import Counter
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.utils import column_index_from_string as col_idx, get_column_letter

from .engine import Result, Row
from .readers import clean_name

DEN_RE = re.compile(r"/\s*(\d+(?:\.\d+)?)\s*$")
TEMPLATE_PATH = Path(__file__).resolve().parents[1] / "templates" / "Tag_Env_Stock_template.xlsx"


def _num(s: str):
    f = float(s)
    return int(f) if f.is_integer() else f


def _load(src) -> openpyxl.Workbook:
    if src is None:
        src = TEMPLATE_PATH
    if isinstance(src, (bytes, bytearray)):
        src = io.BytesIO(src)
    return openpyxl.load_workbook(src)


def _data_rows(ws, first: int) -> list[int]:
    rows = []
    for r in range(first, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if v is None or str(v).strip() == "" or str(v).strip().upper() in ("TOTAL", "GRAND TOTAL"):
            break
        rows.append(r)
    return rows


# --------------------------------------------------------------------------
# template reading
# --------------------------------------------------------------------------
def read_template(src, cfg: dict) -> dict:
    wb = _load(src)
    main = cfg["main_sheet"]
    ws = wb[main["name"]]
    first = main["header_rows"] + 1
    rows_idx = _data_rows(ws, first)
    cases_cols = {spec["cases_col"] for spec in main["stock_columns"].values()}

    ss = cfg["stock_sheets"]
    short_ws = wb[ss["names"]["short"]] if ss["names"]["short"] in wb.sheetnames else None
    full_ws = wb[ss["names"]["full"]] if ss["names"]["full"] in wb.sheetnames else None
    full_by_short, stock_remark = {}, {}
    if short_ws is not None:
        s_first = ss["header_rows"] + 1
        for r in _data_rows(short_ws, s_first):
            k = clean_name(short_ws.cell(r, 1).value).upper()
            if full_ws is not None:
                full_by_short.setdefault(k, full_ws.cell(r, 1).value)
            stock_remark.setdefault(k, short_ws[f"{ss['remarks_col']}{r}"].value)

    rows = []
    label_at = {}
    for r in rows_idx:
        label = ws.cell(r, 1).value
        k = clean_name(label).upper()
        label_at[r] = k
        denoms = {}
        for c in cases_cols:
            m = DEN_RE.search(str(ws[f"{c}{r}"].value or ""))
            if m:
                denoms[c] = float(m.group(1))
        rows.append({"label": label, "full_label": full_by_short.get(k), "denoms": denoms,
                     "remark": ws[f"{main['remarks_col']}{r}"].value, "stock_remark": stock_remark.get(k)})

    merges = {}
    for mr in ws.merged_cells.ranges:
        if mr.min_row < first or mr.min_row not in label_at:
            continue
        col = get_column_letter(mr.min_col)
        labels = [label_at[r] for r in range(mr.min_row, mr.max_row + 1) if r in label_at]
        merges[(label_at[mr.min_row], col)] = labels

    ex = wb[cfg["export_sheet"]["name"]]
    export_labels = []
    for r in range(1, ex.max_row + 1):
        v = clean_name(ex.cell(r, 1).value)
        if v and v.upper() not in ("BRAND", "TOTAL", "GRAND TOTAL") and "STOCK" not in v.upper():
            export_labels.append(v)
    return {"rows": rows, "merges": merges, "export_labels": export_labels}


# --------------------------------------------------------------------------
# helpers
# --------------------------------------------------------------------------
def _row_styles(ws, r: int, ncols: int):
    return [copy.copy(ws.cell(r, c)._style) for c in range(1, ncols + 1)]


def _apply_styles(ws, r: int, styles):
    for c, st in enumerate(styles, start=1):
        ws.cell(r, c)._style = copy.copy(st)


def _shift_formula(f, src_row: int, dst_row: int):
    if not isinstance(f, str) or not f.startswith("="):
        return f
    return re.sub(rf"(?<![A-Z$])(\$?[A-Z]{{1,3}}){src_row}(?!\d)", rf"\g<1>{dst_row}", f)


def _resize_ranges(f, first: int, last: int):
    if not isinstance(f, str) or not f.startswith("="):
        return f
    return re.sub(r"(\$?[A-Z]{1,3})(\d+):(\$?[A-Z]{1,3})(\d+)",
                  lambda m: f"{m.group(1)}{first}:{m.group(3)}{last}", f)


def _unmerge_below(ws, first: int):
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= first:
            ws.unmerge_cells(str(mr))


def _signature(styles) -> tuple:
    return tuple(styles)


def _fmt_den(d):
    return int(d) if float(d).is_integer() else round(float(d), 4)


def _strip_external(wb) -> int:
    """Replace any formula pointing to another workbook with its value 0."""
    n = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value.startswith("=") and "[" in c.value:
                    c.value = 0
                    n += 1
    wb._external_links = []
    return n


# --------------------------------------------------------------------------
# writing
# --------------------------------------------------------------------------
def write_report(res: Result, cfg: dict, report_date: datetime | None, template_src=None) -> tuple[bytes, list[str]]:
    wb = _load(template_src)
    notes: list[str] = []
    main = cfg["main_sheet"]
    stock_cols = main["stock_columns"]
    ws = wb[main["name"]]
    first = main["header_rows"] + 1
    t_rows = _data_rows(ws, first)
    ncols = max(ws.max_column, col_idx(main["remarks_col"]) + 1)

    # ---- capture template styles / formulas --------------------------------
    style_by_key = {clean_name(ws.cell(r, 1).value).upper(): _row_styles(ws, r, ncols) for r in t_rows}
    sig_count = Counter(_signature(s) for s in style_by_key.values())
    default_style = list(sig_count.most_common(1)[0][0]) if sig_count else None
    t_first = t_rows[0] if t_rows else first
    formula_cols = list(main["max_columns"]) + [main["total_can_make"]["col"]]
    formula_tpl = {c: ws[f"{c}{t_first}"].value for c in formula_cols}
    # "=S5" passes an "N/A" text on and breaks the total; "=MAX(S5)" (as the template
    # already uses for Constanta) gives the same number and ignores text.
    for c, f in formula_tpl.items():
        if isinstance(f, str) and re.fullmatch(r"=\$?[A-Z]{1,3}\$?\d+", f):
            formula_tpl[c] = f"=MAX({f[1:]})"
    total_r = (t_rows[-1] + 1) if t_rows else first
    total_vals = [ws.cell(total_r, c).value for c in range(1, ncols + 1)]
    total_style = _row_styles(ws, total_r, ncols)
    blank_style = _row_styles(ws, ws.max_row + 1, ncols)
    height = ws.row_dimensions[t_first].height
    old_last = ws.max_row
    uniform = {c: copy.copy(ws[f"{c}{t_first}"]._style) for c in main.get("uniform_style_columns", [])}

    # conversion factors in row 4
    for cc, f in res.conversion.items():
        if f is not None:
            ws[f"{cc}{main['conversion_row']}"].value = f
    if report_date:
        ws["A1"].value = report_date

    _unmerge_below(ws, first)

    # merged (non-anchor) cells: move any stock there onto the anchor row
    covered: dict[tuple[str, str], str] = {}
    for (anchor, col), keys in res.merges.items():
        for k in keys[1:]:
            covered[(k, col)] = anchor
    rows_by_key = {r.key: r for r in res.rows}
    for (k, col), anchor in covered.items():
        r = rows_by_key.get(k)
        if r and r.qty.get(col):
            a = rows_by_key[anchor]
            a.qty[col] += r.qty[col]
            a.items[col] += r.items[col]
            notes.append(f"{', '.join(r.items[col])} ({r.qty[col]:,.2f} kg) sits inside the shared-tag block "
                         f"of '{a.label.strip()}' and was added to that block.")
            r.qty[col] = 0

    # ---- data rows ----------------------------------------------------------
    cur = first
    row_of: dict[str, int] = {}
    for row in res.rows:
        styles = style_by_key.get(row.key) or default_style
        if styles:
            _apply_styles(ws, cur, styles)
        if height:
            ws.row_dimensions[cur].height = height
        for c in range(1, ncols + 1):
            ws.cell(cur, c).value = None
        ws.cell(cur, 1).value = row.label
        for col, spec in stock_cols.items():
            if (row.key, col) in covered:
                continue
            q = round(row.qty.get(col, 0.0), 2)
            ws[f"{col}{cur}"].value = q
            cc = spec["cases_col"]
            d = row.denom.get(col)
            cell = ws[f"{cc}{cur}"]
            if q and d and res.conversion.get(cc):
                cell.value = f"=({col}{cur}*${cc}${main['conversion_row']})/{_fmt_den(d)}"
            elif q:
                cell.value = "N/A"          # not available in the uploaded BOM
            else:
                cell.value = 0              # no stock -> no cases (shown as "-")
        for c, f in formula_tpl.items():
            ws[f"{c}{cur}"].value = _shift_formula(f, t_first, cur) if f else None
        ws[f"{main['remarks_col']}{cur}"].value = row.remark
        for c, st in uniform.items():
            ws[f"{c}{cur}"]._style = copy.copy(st)
        row_of[row.key] = cur
        cur += 1
    last = cur - 1

    # ---- total row ----------------------------------------------------------
    _apply_styles(ws, cur, total_style)
    if height:
        ws.row_dimensions[cur].height = ws.row_dimensions[total_r].height or height
    for c, v in enumerate(total_vals, start=1):
        ws.cell(cur, c).value = _resize_ranges(v, first, last)
    end = cur
    for r in range(end + 1, max(old_last, end) + 1):
        _apply_styles(ws, r, blank_style)
        for c in range(1, ncols + 1):
            ws.cell(r, c).value = None

    # ---- merges (column by column, exactly as the template had them) --------
    for (anchor, col), keys in res.merges.items():
        rs = [row_of[k] for k in keys if k in row_of]
        if len(rs) < 2 or rs != list(range(rs[0], rs[0] + len(rs))):
            continue
        ws.merge_cells(f"{col}{rs[0]}:{col}{rs[-1]}")

    fr = ws.freeze_panes
    if fr and int(re.sub(r"\D", "", fr) or 0) > first:
        ws.freeze_panes = f"B{first}"
        notes.append(f"Freeze panes in the template was at {fr} (inside the data); set to B{first}.")

    # ---- hidden stock sheets -----------------------------------------------
    # both sheets share row order; the short sheet's labels match the main sheet
    short_name = cfg["stock_sheets"]["names"]["short"]
    short_keys = []
    if short_name in wb.sheetnames:
        sw = wb[short_name]
        short_keys = [clean_name(sw.cell(r, 1).value).upper()
                      for r in _data_rows(sw, cfg["stock_sheets"]["header_rows"] + 1)]
    for which in ("short", "full"):
        name = cfg["stock_sheets"]["names"][which]
        if name in wb.sheetnames:
            _write_stock_sheet(wb[name], res, cfg, which, report_date, short_keys)

    # ---- export sheet -------------------------------------------------------
    _write_export_sheet(wb[cfg["export_sheet"]["name"]], res, cfg)

    n_ext = _strip_external(wb)
    if n_ext:
        notes.append(f"{n_ext} leftover link(s) to an external workbook were replaced with 0.")
    wb.calculation.fullCalcOnLoad = True
    out = io.BytesIO()
    wb.save(out)
    return out.getvalue(), notes


def _write_stock_sheet(ws, res: Result, cfg: dict, which: str, report_date, short_keys: list[str]):
    ss = cfg["stock_sheets"]
    first = ss["header_rows"] + 1
    t_rows = _data_rows(ws, first)
    ncols = max(ws.max_column, col_idx(ss["remarks_col"]))
    styles_list = [_row_styles(ws, r, ncols) for r in t_rows]
    # match template rows by the short label (same row order in both sheets)
    style_by_main = {}
    for i, k in enumerate(short_keys):
        if i < len(styles_list):
            style_by_main.setdefault(k, styles_list[i])
    default = Counter(tuple(s) for s in styles_list).most_common(1)[0][0] if styles_list else None
    tot_r = (t_rows[-1] + 1) if t_rows else first
    tot_rows = []
    r = tot_r
    while r <= ws.max_row and str(ws.cell(r, 1).value or "").strip():
        tot_rows.append(([ws.cell(r, c).value for c in range(1, ncols + 1)], _row_styles(ws, r, ncols),
                         ws.row_dimensions[r].height))
        r += 1
    blank = _row_styles(ws, ws.max_row + 1, ncols)
    height = ws.row_dimensions[first].height
    old_last = max(ws.max_row, r)
    if report_date:
        ws["A1"].value = report_date
    _unmerge_below(ws, first)

    cmap = ss["column_map"]
    covered = {(k, col) for (a, col), keys in res.merges.items() for k in keys[1:]}
    cur = first
    row_of = {}
    for row in res.rows:
        st = style_by_main.get(row.key) or default
        if st:
            _apply_styles(ws, cur, st)
        if height:
            ws.row_dimensions[cur].height = height
        for c in range(1, ncols + 1):
            ws.cell(cur, c).value = None
        ws.cell(cur, 1).value = row.label if which == "short" else (row.full_label or row.label)
        for mcol, hcol in cmap.items():
            if (row.key, mcol) in covered:
                continue
            ws[f"{hcol}{cur}"].value = round(row.qty.get(mcol, 0.0), 2)
        ws[f"{ss['remarks_col']}{cur}"].value = row.stock_remark
        row_of[row.key] = cur
        cur += 1
    last = cur - 1
    new_tot_r = cur  # the TOTAL row moves here; Grand Total refers to it
    for vals, st, h in tot_rows:
        _apply_styles(ws, cur, st)
        if h:
            ws.row_dimensions[cur].height = h
        for c, v in enumerate(vals, start=1):
            if isinstance(v, str) and v.startswith("="):
                if ":" in v:
                    v = _resize_ranges(v, first, last)   # SUBTOTAL over the full data range
                else:
                    v = _shift_formula(v, tot_r, new_tot_r)
            ws.cell(cur, c).value = v
        cur += 1
    for rr in range(cur, old_last + 1):
        _apply_styles(ws, rr, blank)
        for c in range(1, ncols + 1):
            ws.cell(rr, c).value = None
    for (anchor, mcol), keys in res.merges.items():
        hcol = cmap.get(mcol)
        rs = [row_of[k] for k in keys if k in row_of]
        if hcol and len(rs) > 1 and rs == list(range(rs[0], rs[0] + len(rs))):
            ws.merge_cells(f"{hcol}{rs[0]}:{hcol}{rs[-1]}")


def _write_export_sheet(ws, res: Result, cfg: dict):
    ec, tc = cfg["export_sheet"]["env_col"], cfg["export_sheet"]["tag_col"]
    for r in range(1, ws.max_row + 1):
        k = clean_name(ws.cell(r, 1).value).upper()
        er = res.export_rows.get(k)
        if er is None:
            continue
        ws[f"{ec}{r}"].value = round(er.env, 2)
        tag_cell = ws[f"{tc}{r}"]
        if tag_cell.value is not None or er.tag:
            tag_cell.value = round(er.tag, 2)
