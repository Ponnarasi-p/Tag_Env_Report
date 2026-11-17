"""Learn which report row each stock item belongs to, from a manually prepared
Tag Env Stock report and the Stock Summary it was made from.

For every stock item, the quantity is looked up in the column its machine line
maps to; when exactly one cell holds that value, the item -> row link is
recorded. Existing entries in the map are kept unless --replace is given.

Usage:
    python tools/learn_item_row_map.py MANUAL_REPORT.xlsx STOCK_SUMMARY.xls [--replace]
"""
from __future__ import annotations

import argparse
import csv
import json
import sys
from collections import defaultdict
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from tag_env.readers import read_stock_summary, clean_name  # noqa: E402

MAP_FILE = ROOT / "config" / "item_row_map.csv"


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("manual_report")
    ap.add_argument("stock_summary")
    ap.add_argument("--replace", action="store_true")
    a = ap.parse_args()

    cfg = json.loads((ROOT / "config" / "settings.json").read_text())
    ml2col = cfg["machine_line_to_column"]
    wb = openpyxl.load_workbook(a.manual_report, data_only=False)
    ws = wb[cfg["main_sheet"]["name"]]
    ex = wb[cfg["export_sheet"]["name"]]

    cells: dict[tuple[str, float], list[tuple[str, str]]] = defaultdict(list)
    for r in range(cfg["main_sheet"]["header_rows"] + 1, ws.max_row + 1):
        label = clean_name(ws.cell(r, 1).value)
        if not label:
            continue
        for col in cfg["main_sheet"]["stock_columns"]:
            v = ws[f"{col}{r}"].value
            if isinstance(v, (int, float)) and v:
                cells[(col, round(float(v), 2))].append(("main", label))
    for r in range(1, ex.max_row + 1):
        label = clean_name(ex.cell(r, 1).value)
        for col in (cfg["export_sheet"]["env_col"], cfg["export_sheet"]["tag_col"]):
            v = ex[f"{col}{r}"].value
            if label and isinstance(v, (int, float)) and v:
                cells[("EXPORT", round(float(v), 2))].append(("export", label))

    stock, _, _ = read_stock_summary(a.stock_summary)
    existing = {}
    if MAP_FILE.exists() and not a.replace:
        with MAP_FILE.open(newline="", encoding="utf-8") as f:
            for row in csv.DictReader(f):
                existing[row["item"]] = row
    learned = 0
    for _, s in stock.iterrows():
        col = ml2col.get(s["machine_line"])
        if not col or not s["qty"]:
            continue
        key = round(float(s["qty"]), 2)
        hits = cells.get((col, key), []) + cells.get(("EXPORT", key), [])
        if len(hits) == 1 and s["item"] not in existing:
            sheet, label = hits[0]
            existing[s["item"]] = {"item": s["item"], "sheet": sheet, "row_label": label}
            learned += 1
    with MAP_FILE.open("w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=["item", "sheet", "row_label"])
        w.writeheader()
        for k in sorted(existing):
            w.writerow(existing[k])
    print(f"learned {learned} new item(s); map now has {len(existing)} entries -> {MAP_FILE}")


if __name__ == "__main__":
    main()
