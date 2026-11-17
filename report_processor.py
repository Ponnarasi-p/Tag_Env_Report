"""
Report correction rules.

Pure data transformation: takes the sheets of the uploaded workbook as lists
of rows and returns a cleaned model plus a log of what changed.
"""
from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from datetime import date, datetime
from typing import Any

# Header names are compared after normalising: lower-case, text in brackets
# removed, everything except letters/digits removed.
# e.g. "Width (mm)" -> "width", "Lot No." -> "lotno".
CONFIG = {
    "item_name": ["itemname", "item", "itemdescription", "itemdesc", "productname", "materialname", "papername"],
    "serial_no": ["sno", "slno", "srno", "serialno", "no", "sn", "sl"],
    "raw_aging_remove": ["lotcode", "lotno", "lotnumber", "supplierlotno", "supplierlotnumber",
                         "category", "gsm", "qty", "quantity", "machineline"],
    "tagenv_remove": ["lotcode", "lotno", "lotnumber", "supplierlotno", "supplierlotnumber",
                      "width", "gsm", "sourcetype"],
    "qtn": ["qtn", "qtnkg", "qtnmt", "qtnton", "qtntons"],
    # Quantity column renamed to "QTY (kg)" on the Raw Aging / TagEnv sheets, in this order of preference.
    "qty_kg": ["qtn", "qty", "quantity"],
    "holding_days": ["holdingdays", "holdingday", "holddays"],
    "stock": ["stock", "stockkg", "closingstock", "currentstock", "stockqty", "balancestock",
              "availablestock", "stockinhand", "totalstock", "stockinkg"],
    "null_tokens": ["", "null", "nil", "n/a", "na", "#n/a", "none", "undefined", "nan"],
    "total_tokens": ["total", "grandtotal", "subtotal", "overalltotal"],
}


def norm(v: Any) -> str:
    if v is None:
        return ""
    s = str(v).lower()
    s = re.sub(r"\([^)]*\)", "", s)
    s = re.sub(r"\[[^\]]*\]", "", s)
    return re.sub(r"[^a-z0-9]", "", s)


def is_num(v: Any) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool) and not (isinstance(v, float) and math.isnan(v))


def is_empty(v: Any) -> bool:
    if v is None:
        return True
    if isinstance(v, (datetime, date)):
        return False
    if isinstance(v, float):
        return math.isnan(v)
    if isinstance(v, (int, bool)):
        return False
    return str(v).strip().lower() in CONFIG["null_tokens"]


def find_col(headers: list, aliases: list) -> int:
    for i, h in enumerate(headers):
        if norm(h) in aliases:
            return i
    return -1


def sheet_kind(name: str) -> str:
    n = norm(name)
    if "tagenv" in n or ("all" in n and "tag" in n):
        return "tagenv"
    if "raw" in n and ("aging" in n or "ageing" in n):
        return "raw"
    return "other"


_NUM_RE = re.compile(r"^-?(\d{1,3}(,\d{3})+|\d+)?(\.\d+)?$")


def to_number_if_numeric(v: Any) -> Any:
    """Numbers stored as text ("1,250.50") become numbers; codes like "00123" stay text."""
    if not isinstance(v, str):
        return v
    s = v.strip()
    if s in ("", "-") or not _NUM_RE.match(s) or re.match(r"^-?0\d", s):
        return v
    try:
        n = float(s.replace(",", ""))
    except ValueError:
        return v
    return int(n) if n.is_integer() else n


def round_half_away(n: float) -> int:
    r = int(math.floor(abs(n) + 0.5))
    return -r if n < 0 and r else r


def detect_header_row(rows: list) -> int:
    limit = min(len(rows), 25)
    for r in range(limit):
        if find_col(rows[r] or [], CONFIG["item_name"]) != -1:
            return r
    best, best_count = 0, -1
    for i in range(limit):
        c = sum(1 for v in (rows[i] or []) if isinstance(v, str) and v.strip())
        if c > best_count:
            best, best_count = i, c
    return best


@dataclass
class Row:
    type: str            # data | summary | detail | total
    cells: list
    group_id: int = 0
    count: int = 0


@dataclass
class Sheet:
    name: str
    kind: str
    title_rows: list = field(default_factory=list)
    headers: list = field(default_factory=list)
    rows: list = field(default_factory=list)
    numeric_cols: list = field(default_factory=list)
    item_col: int = -1


def process_workbook(sheets: list[tuple[str, list[list]]]) -> tuple[list[Sheet], list[dict]]:
    log: list[dict] = []
    out = [process_sheet(name, rows, log) for name, rows in sheets]
    if not any(s.kind == "raw" for s in out):
        log.append({"sheet": "-", "level": "warn", "message": 'No "Raw Aging Stock" sheet found; its column removals were skipped.'})
    if not any(s.kind == "tagenv" for s in out):
        log.append({"sheet": "-", "level": "warn", "message": 'No "All TagEnv Aging Stock" sheet found; grouping was skipped.'})
    return out, log


def process_sheet(name: str, raw_rows: list[list], log: list[dict]) -> Sheet:
    kind = sheet_kind(name)

    def note(msg: str, level: str = "info"):
        log.append({"sheet": name, "message": msg, "level": level})

    rows = [list(r or []) for r in raw_rows]
    while rows and all(is_empty(v) for v in rows[-1]):
        rows.pop()
    if not rows:
        note("Sheet is empty; copied as-is.")
        return Sheet(name=name, kind=kind)

    h = detect_header_row(rows)
    title_rows = ["   ".join(str(v) for v in r if not is_empty(v))
                  for r in rows[:h] if not all(is_empty(v) for v in r)]
    headers = ["" if is_empty(v) else str(v).strip() for v in rows[h]]
    width = max([len(headers)] + [len(r) for r in rows])
    headers += [""] * (width - len(headers))
    data = [[to_number_if_numeric(v) for v in (r + [None] * (width - len(r)))] for r in rows[h + 1:]]

    def drop_cols(pred) -> list[str]:
        nonlocal headers, data
        idx = {ci for ci, hd in enumerate(headers) if pred(hd, ci)}
        if not idx:
            return []
        removed = [headers[i] for i in sorted(idx)]
        headers = [hd for ci, hd in enumerate(headers) if ci not in idx]
        data = [[v for ci, v in enumerate(r) if ci not in idx] for r in data]
        return removed

    # Columns with neither a header nor any data.
    keep = [hd != "" or any(not is_empty(r[ci]) for r in data) for ci, hd in enumerate(headers)]
    drop_cols(lambda _hd, ci: not keep[ci])

    # ---- Rule 3 / 4: column removal
    if kind in ("raw", "tagenv"):
        lst = CONFIG["raw_aging_remove"] if kind == "raw" else CONFIG["tagenv_remove"]
        removed = drop_cols(lambda hd, _ci: norm(hd) in lst)
        if removed:
            note("Removed columns: " + ", ".join(removed) + ".")
        else:
            note("None of the columns to remove were present.", "warn")

    item_col = find_col(headers, CONFIG["item_name"])
    if item_col == -1:
        note('No "Item Name" column found; empty-row removal and grouping skipped for this sheet.', "warn")

    # ---- Rule 2: Day Consumption tons -> kg, recompute No. of Days
    day_col = days_col = -1
    for ci, hd in enumerate(headers):
        n = norm(hd)
        if "consumption" not in n:
            continue
        if "noofdays" in n or "days" in n:
            if days_col == -1:
                days_col = ci
        elif day_col == -1:
            day_col = ci
    if day_col != -1:
        stock_col = find_col(headers, CONFIG["stock"])
        converted = 0
        for r in data:
            t = r[day_col]
            if is_num(t):
                r[day_col] = t * 1000
                converted += 1
            if days_col != -1:
                kg = r[day_col]
                if stock_col != -1 and is_num(r[stock_col]) and is_num(kg):
                    r[days_col] = None if kg == 0 else r[stock_col] / kg
                elif is_num(r[days_col]) and is_num(t):
                    # Old value was stock / tons; using kg means dividing by 1000.
                    r[days_col] = r[days_col] / 1000
        old = headers[day_col]
        unit_re = re.compile(r"\b(tons?|mt|tonnes?)\b", re.I)
        headers[day_col] = unit_re.sub("Kg", old) if unit_re.search(old) else old + " (Kg)"
        note(f'Converted {converted} "{old}" values from tons to kg.')
        if days_col != -1:
            if stock_col != -1:
                note(f'"{headers[days_col]}" recalculated as "{headers[stock_col]}" ÷ Day Consumption (Kg).')
            else:
                note(f'"{headers[days_col]}" recalculated using Day Consumption in kg '
                     "(no stock column found, so the existing value was scaled by 1/1000).")

    # ---- Rule 1: remove rows that only have an item name
    serial_cols = [ci for ci, hd in enumerate(headers) if norm(hd) in CONFIG["serial_no"]]
    before = len(data)
    data = [r for r in data
            if any(not is_empty(v) for ci, v in enumerate(r) if ci != item_col and ci not in serial_cols)]
    if before - len(data):
        note(f"Removed {before - len(data)} row(s) with no details apart from the item name (or fully blank).")
    # ---- Raw Aging Stock: Holding Days, highest first (Total rows stay last)
    if kind == "raw":
        hd = find_col(headers, CONFIG["holding_days"])
        if hd == -1:
            note('No "Holding Days" column found, so rows were not sorted.', "warn")
        else:
            def raw_total(r):
                return item_col != -1 and norm(r[item_col]) in CONFIG["total_tokens"]
            body = [r for r in data if not raw_total(r)]
            body.sort(key=lambda r: (0, -r[hd]) if is_num(r[hd]) else (1, 0))  # stable for ties
            data = body + [r for r in data if raw_total(r)]
            note("Sorted by Holding Days (highest first).")

    for sc in serial_cols:  # renumber S.No
        n = 0
        for r in data:
            if is_num(r[sc]):
                n += 1
                r[sc] = n

    numeric_cols = []
    for ci in range(len(headers)):
        nums = sum(1 for r in data if is_num(r[ci]))
        texts = sum(1 for r in data if not is_num(r[ci]) and not is_empty(r[ci])
                    and not isinstance(r[ci], (datetime, date)))
        numeric_cols.append(nums > 0 and nums >= texts)

    def is_total(r):
        return item_col != -1 and norm(r[item_col]) in CONFIG["total_tokens"]

    # ---- Rule 4: grouping on TagEnv
    if kind == "tagenv" and item_col != -1:
        out_rows = group_rows(data, headers, item_col, is_total, note)
    else:
        out_rows = [Row("total" if is_total(r) else "data", r) for r in data]

    # ---- Rule 6: round decimals (after grouping so averages use exact values)
    rounded = 0
    for row in out_rows:
        for ci, v in enumerate(row.cells):
            if isinstance(v, float) and not math.isnan(v):
                if not v.is_integer():
                    rounded += 1
                row.cells[ci] = round_half_away(v)
    if rounded:
        note(f"Rounded {rounded} decimal value(s) to whole numbers.")

    # ---- Rule 5 / 7: zero -> "-"
    zeros = 0
    zero_re = re.compile(r"^\s*-?0+(\.0+)?\s*$")
    for row in out_rows:
        for ci, v in enumerate(row.cells):
            if (is_num(v) and v == 0) or (isinstance(v, str) and zero_re.match(v)):
                row.cells[ci] = "-"
                zeros += 1
    if zeros:
        note(f'Replaced {zeros} zero value(s) with "-".')

    # ---- Quantity heading becomes "QTY (kg)" (QTN preferred, then Qty, then Quantity)
    if kind in ("raw", "tagenv"):
        for alias in CONFIG["qty_kg"]:
            ci = next((i for i, hd in enumerate(headers) if norm(hd) == alias), -1)
            if ci != -1:
                if headers[ci] != "QTY (kg)":
                    note(f'Renamed heading "{headers[ci]}" to "QTY (kg)".')
                    headers[ci] = "QTY (kg)"
                break

    return Sheet(name=name, kind=kind, title_rows=title_rows, headers=headers, rows=out_rows,
                 numeric_cols=numeric_cols, item_col=item_col)


def group_rows(data, headers, item_col, is_total, note) -> list[Row]:
    # Every quantity column (QTN, Qty, Quantity...) is totalled on grouped rows.
    qty_aliases = set(CONFIG["qtn"]) | set(CONFIG["qty_kg"])
    sum_cols = {ci for ci, h in enumerate(headers) if norm(h) in qty_aliases}
    qtn_col = min(sum_cols) if sum_cols else -1
    hd_col = find_col(headers, CONFIG["holding_days"])
    if qtn_col == -1:
        note('No QTN / Qty column found for subtotals.', "warn")
    if hd_col == -1:
        note('No "Holding Days" column found for averages.', "warn")

    groups: dict[str, list] = {}
    totals = []
    for r in data:
        if is_total(r):
            totals.append(r)
            continue
        key = "\0blank" if is_empty(r[item_col]) else str(r[item_col]).strip().lower()
        groups.setdefault(key, []).append(r)

    def hd_key(cells):
        # Descending by Holding Days; rows without a value go last.
        v = cells[hd_col] if hd_col != -1 else None
        return (0, -v) if is_num(v) else (1, 0)

    blocks: list[list[Row]] = []
    gid = multi = 0
    for recs in groups.values():
        if len(recs) == 1:
            blocks.append([Row("data", recs[0])])
            continue
        multi += 1
        gid += 1
        summary = []
        for ci in range(len(headers)):
            vals = [r[ci] for r in recs]
            if ci == item_col:
                summary.append(recs[0][item_col])
            elif ci in sum_cols:
                nums = [v for v in vals if is_num(v)]
                summary.append(float(sum(nums)) if nums else None)
            elif ci == hd_col:
                nums = [v for v in vals if is_num(v)]
                summary.append(sum(nums) / len(nums) if nums else None)
            else:
                # Show a value on the summary row only when every record agrees.
                first = vals[0]
                same = all(str(v) == str(first) for v in vals)
                summary.append(first if same and not isinstance(first, (datetime, date)) else None)
        details = sorted(recs, key=hd_key)
        blocks.append([Row("summary", summary, gid, len(recs))] + [Row("detail", r, gid) for r in details])

    if hd_col != -1:
        blocks.sort(key=lambda b: hd_key(b[0].cells))  # stable: ties keep original order
    out = [row for b in blocks for row in b]
    out.extend(Row("total", r) for r in totals)
    note(f"Grouped {multi} repeated item name(s): quantity totalled, Holding Days averaged. "
         f"{len(groups)} unique items shown"
         + (", sorted by Holding Days (highest first)." if hd_col != -1 else "."))
    return out
