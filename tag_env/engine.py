"""Business logic: stock items -> report rows/columns, Pcs-per-CFC lookup and
Make Cases calculation.

Make Cases = (Available KG x Conversion factor of the column's machine line)
             / Pcs per CFC

Pcs per CFC comes from the Exploded BOM: the total ENV (or TAG) pieces in one
CFC of the finished good (FG) the item is used in. A variety pack holding
three flavours of 80 envelopes therefore gives 240, matching the manual report.

Which FG is used when an item is in several:
  1. an FG with a pending order (largest pending quantity wins);
  2. otherwise the most common Pcs per CFC across the item's FGs (ties: smaller).
Everything comes from the uploaded files. If an item is not in the uploaded
BOM, Make Cases shows N/A and the user is asked to submit a new BOM; no value
is taken from an older report.
"""
from __future__ import annotations

import csv
import json
import re
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from .readers import clean_name

ROOT = Path(__file__).resolve().parents[1]
CONFIG_DIR = ROOT / "config"


# --------------------------------------------------------------------------
# configuration
# --------------------------------------------------------------------------
def load_settings(path: Path | None = None) -> dict:
    return json.loads((path or CONFIG_DIR / "settings.json").read_text(encoding="utf-8"))


def load_item_row_map(path: Path | None = None) -> dict[str, tuple[str, str]]:
    p = path or CONFIG_DIR / "item_row_map.csv"
    out: dict[str, tuple[str, str]] = {}
    if p.exists():
        with p.open(newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                out[clean_name(r["item"])] = (r["sheet"], clean_name(r["row_label"]))
    return out


def load_mjp_m4_map(path: Path | None = None) -> dict[str, str]:
    p = path or CONFIG_DIR / "mjp_m4_item_mapping.csv"
    out: dict[str, str] = {}
    if p.exists():
        with p.open(newline="", encoding="utf-8") as f:
            for r in csv.DictReader(f):
                out[clean_name(r["MJP Item Name"])] = clean_name(r["M4 Item Name"])
    return out


# --------------------------------------------------------------------------
# name normalisation used to match a new item to an existing row
# --------------------------------------------------------------------------
_DROP_TOKENS = {"DC", "SC", "SF", "HS", "C250", "ENV", "TBGS", "TBG", "PENV", "PTAG", "PK45",
                "TAG", "FT", "MD-20", "POLY"}
_DROP_PHRASES = [r"\(\s*70\s*GSM\s*\)", r"\(\s*WO\s*GLUE\s*\)", r"\(\s*V\d\s*\)", r"\(\s*3 MIN\s*\)"]


def row_key(name: str) -> str:
    s = clean_name(name).upper()
    s = re.sub(r"^(ENV|TAG)\s+", "", s)
    for p in _DROP_PHRASES:
        s = re.sub(p, " ", s)
    s = s.replace("-FT", " ").replace("TBGS-", "TBGS ")
    toks = [t for t in re.split(r"\s+", s) if t and t not in _DROP_TOKENS and t != "-"]
    return " ".join(toks)


def item_kind(item: str, bom_type: str | None = None) -> str:
    if bom_type in ("ENV", "TAG"):
        return bom_type
    return "TAG" if clean_name(item).upper().startswith("TAG") else "ENV"


# --------------------------------------------------------------------------
# data structures
# --------------------------------------------------------------------------
@dataclass
class Row:
    label: str                     # as shown in column A (template text kept as is)
    full_label: str = ""           # label for the "(2)" sheet
    is_new: bool = False
    template_index: int | None = None
    remark: str | None = None
    stock_remark: str | None = None
    items: dict[str, list[str]] = field(default_factory=lambda: defaultdict(list))  # col -> items
    qty: dict[str, float] = field(default_factory=lambda: defaultdict(float))        # col -> kg
    denom: dict[str, float | None] = field(default_factory=dict)                     # col -> pcs/CFC
    denom_source: dict[str, str] = field(default_factory=dict)

    @property
    def key(self) -> str:
        return clean_name(self.label).upper()


@dataclass
class ExportRow:
    label: str
    env: float = 0.0
    tag: float = 0.0
    found: bool = False


@dataclass
class Result:
    rows: list[Row]
    export_rows: dict[str, ExportRow]
    audit: pd.DataFrame
    warnings: list[str]
    summary: dict
    merges: dict[tuple[str, str], list[str]]      # (anchor label key, stock col) -> labels spanned
    conversion: dict[str, float]                  # cases col -> factor


# --------------------------------------------------------------------------
# BOM index
# --------------------------------------------------------------------------
class BomIndex:
    def __init__(self, bom: pd.DataFrame, fg_uom: str = "CFC"):
        b = bom[bom["fg_uom"] == fg_uom.upper()]
        b = b[b["item_type"].isin(["ENV", "TAG"])]
        # only piece counts; a few BOM lines carry tags/envelopes in KGS
        b = b[b["uom"].astype(str).str.strip().str.upper().isin(["PCS", "NOS", "PC", "NO"])]
        self.items = set(b["item"])
        self.total = b.groupby(["fg", "item_type"])["qty"].sum().to_dict()
        self.by_item: dict[str, list[tuple[str, str, float]]] = defaultdict(list)
        self.by_fg: dict[str, set[str]] = defaultdict(set)
        self.type_of: dict[str, str] = {}
        for fg, item, t, q in zip(b["fg"], b["item"], b["item_type"], b["qty"]):
            self.by_item[item].append((fg, t, q))
            self.by_fg[fg].add(item)
            self.type_of[item] = t
        self.upper = {i.upper(): i for i in self.items}
        self.anomalies: set[str] = set()

    def resolve(self, item: str, mjp_m4: dict[str, str]) -> str | None:
        if item in self.items:
            return item
        m4 = mjp_m4.get(item)
        if m4 and m4 in self.items:
            return m4
        return self.upper.get(item.upper())

    def candidates(self, bom_item: str) -> dict[str, float]:
        """FG -> total pcs of the same component type per CFC.

        Totals below 1 piece are BOM data errors (e.g. 0.0072 PCS) and are skipped.
        """
        out = {}
        for fg, t, q in self.by_item.get(bom_item, []):
            tot = float(self.total.get((fg, t), q))
            if tot < 1:
                self.anomalies.add(f"{fg} / {bom_item} = {q:g}")
                continue
            out[fg] = tot
        return out


# --------------------------------------------------------------------------
# main entry
# --------------------------------------------------------------------------
def build(stock: pd.DataFrame, orders: pd.DataFrame, bom: pd.DataFrame, template_info: dict,
          settings: dict | None = None, item_row_map: dict | None = None,
          mjp_m4: dict | None = None) -> Result:
    cfg = settings or load_settings()
    item_row_map = item_row_map if item_row_map is not None else load_item_row_map()
    mjp_m4 = mjp_m4 if mjp_m4 is not None else load_mjp_m4_map()
    main = cfg["main_sheet"]
    stock_cols = main["stock_columns"]
    ml2col = cfg["machine_line_to_column"]
    markers = [m.upper() for m in cfg["export_sheet"]["markers"]]
    warnings: list[str] = []

    # conversion factor per cases column
    conversion: dict[str, float] = {}
    for col, spec in stock_cols.items():
        f = cfg["conversion_factors"].get(spec["conversion"])
        if f is None:
            warnings.append(f"No conversion factor configured for '{spec['conversion']}' "
                            f"(column {spec['cases_col']}); Make Cases left empty for that column.")
        conversion[spec["cases_col"]] = f

    # ---- rows from template -------------------------------------------------
    rows: list[Row] = []
    by_key: dict[str, Row] = {}
    by_rowkey: dict[str, Row] = {}
    for i, t in enumerate(template_info["rows"]):
        # Remarks are left blank: they are manual notes about a past date, not data from the inputs.
        r = Row(label=t["label"], full_label=t.get("full_label") or t["label"], template_index=i)
        rows.append(r)
        by_key.setdefault(r.key, r)
        by_rowkey.setdefault(row_key(r.label), r)
    export_rows = {clean_name(l).upper(): ExportRow(label=l) for l in template_info["export_labels"]}

    def new_row(label: str, full: str) -> Row:
        r = Row(label=label, full_label=full, is_new=True)
        rows.append(r)
        by_key[r.key] = r
        by_rowkey.setdefault(row_key(label), r)
        return r

    bidx = BomIndex(bom, cfg["denominator"]["fg_uom"])
    pending = orders.groupby("product")["pending_qty"].sum().to_dict() if len(orders) else {}
    pending_products = set(pending)

    audit = []
    unmatched_items, skipped_lines, new_rows = [], Counter(), []

    # ---- place every stock item --------------------------------------------
    for _, s in stock.iterrows():
        item, ml, qty = s["item"], s["machine_line"], float(s["qty"])
        col = ml2col.get(ml)
        if col is None:
            skipped_lines[ml or "(blank)"] += 1
            continue
        up = item.upper()
        mapped = item_row_map.get(item)
        # export (Ghana / Abidjan / Angola) items
        if (mapped and mapped[0] == "export") or any(m in up for m in markers):
            key = (mapped[1] if mapped else item).upper()
            er = export_rows.get(clean_name(key).upper())
            kind = item_kind(item, bidx.type_of.get(bidx.resolve(item, mjp_m4) or ""))
            if er is None:
                warnings.append(f"Export item '{item}' ({qty:,.2f} kg) has no row on the "
                                f"'{cfg['export_sheet']['name']}' sheet; add it to the template.")
                continue
            er.found = True
            if kind == "TAG":
                er.tag += qty
            else:
                er.env += qty
            audit.append({"Item": item, "Machine Line": ml, "Stock Kg": qty, "Sheet": cfg["export_sheet"]["name"],
                          "Row": er.label, "Column": cfg["export_sheet"]["tag_col" if kind == "TAG" else "env_col"],
                          "Matched By": "export sheet"})
            continue

        # main sheet row
        row, how = None, ""
        if mapped and mapped[0] == "main":
            row = by_key.get(mapped[1].upper())
            how = "item map"
            if row is None:  # mapped to a label the template does not have yet
                row = new_row(mapped[1] + " ", mapped[1] + " ")
                new_rows.append(row.label.strip())
                how = "item map (new row)"
        if row is None:
            row = by_rowkey.get(row_key(item))
            how = "name match" if row else ""
        if row is None:
            kind = item_kind(item)
            if kind == "ENV":
                label = "ENV " + row_key(item) + " "
            else:
                label = item
            row = new_row(label, item)
            new_rows.append(label.strip())
            unmatched_items.append(item)
            how = "new row"
        row.items[col].append(item)
        row.qty[col] += qty
        audit.append({"Item": item, "Machine Line": ml, "Stock Kg": qty, "Sheet": main["name"],
                      "Row": row.label.strip(), "Column": col, "Matched By": how})

    # ---- denominators -------------------------------------------------------
    item_denoms: dict[str, tuple[float | None, str, str]] = {}
    for row in rows:
        for col, items in row.items.items():
            if not row.qty.get(col):
                continue
            # use the item with the most stock in the cell
            items_sorted = sorted(items, key=lambda it: -float(stock.loc[stock["item"] == it, "qty"].sum()))
            chosen = None
            for it in items_sorted:
                d, src, fg = _denominator(it, bidx, mjp_m4, pending, pending_products)
                item_denoms[it] = (d, src, fg)
                if chosen is None or (chosen[0] is None and d is not None):
                    chosen = (d, src, fg)
            row.denom[col] = chosen[0]
            row.denom_source[col] = chosen[1] + (f" [{chosen[2]}]" if chosen[2] else "")
    # cells without stock have no item, so no Pcs per CFC (Make Cases is 0)

    # ---- audit enrichment ---------------------------------------------------
    for a in audit:
        if a["Sheet"] != main["name"]:
            continue
        col = a["Column"]
        d, src, fg = item_denoms.get(a["Item"], (None, "", ""))
        spec = stock_cols[col]
        f = conversion.get(spec["cases_col"])
        a.update({"Group": spec["group"], "Type": spec["kind"], "Conversion Machine Line": spec["conversion"],
                  "Conversion (Pcs/Kg)": f, "Pcs per CFC (BOM)": d, "Pcs/CFC Source": src, "FG Used": fg,
                  "Pending Order Qty (FG)": pending.get(fg, 0) if fg else 0,
                  "Make Cases": (a["Stock Kg"] * f / d) if (d and f) else None})
    audit_df = pd.DataFrame(audit)

    # ---- new rows go in alphabetically, like the manual report ------------
    rows = _order_rows(rows)

    # ---- merges for shared tags -------------------------------------------
    merges = _plan_merges(rows, template_info.get("merges", {}), stock_cols)

    # ---- warnings -----------------------------------------------------------
    if len(audit_df):
        miss = audit_df[(audit_df["Sheet"] == main["name"]) & audit_df["Pcs per CFC (BOM)"].isna()
                        & (audit_df["Stock Kg"] != 0)]
        missing_bom_items = miss[["Item", "Machine Line", "Stock Kg", "Row", "Pcs/CFC Source"]].reset_index(drop=True)
    else:
        missing_bom_items = pd.DataFrame(columns=["Item", "Machine Line", "Stock Kg", "Row", "Pcs/CFC Source"])
    if len(missing_bom_items):
        warnings.append(f"{len(missing_bom_items)} item(s) are not available in the Exploded BOM, so their "
                        "Make Cases shows N/A. Please submit a new Exploded BOM that includes them.")
    if new_rows:
        warnings.append(f"{len(new_rows)} new row(s) added for items not in the template: "
                        + ", ".join(new_rows[:10]) + ("..." if len(new_rows) > 10 else ""))
    if skipped_lines:
        warnings.append("Ignored stock lines that are not Tag/Envelope machine lines: "
                        + ", ".join(f"{k} ({v})" for k, v in skipped_lines.most_common()))
    missing_fg = sorted(p for p in pending_products if p not in bidx.by_fg)
    if missing_fg:
        warnings.append(f"{len(missing_fg)} pending-order product(s) are not in the Exploded BOM "
                        "(their orders cannot steer Pcs per CFC): " + ", ".join(missing_fg[:6])
                        + ("..." if len(missing_fg) > 6 else ""))
    if bidx.anomalies:
        warnings.append(f"Exploded BOM has {len(bidx.anomalies)} tag/envelope line(s) under 1 piece per CFC "
                        "(ignored as data errors): " + "; ".join(sorted(bidx.anomalies)[:5]))
    unused_export = [e.label for e in export_rows.values() if not e.found]

    summary = {
        "stock_lines": len(stock),
        "placed_main": int((audit_df["Sheet"] == main["name"]).sum()) if len(audit_df) else 0,
        "placed_export": int((audit_df["Sheet"] != main["name"]).sum()) if len(audit_df) else 0,
        "rows": len(rows),
        "new_rows": len(new_rows),
        "pending_products": len(pending_products),
        "pending_products_in_bom": len(pending_products) - len(missing_fg),
        "bom_fgs": len(bidx.by_fg),
        "cells_from_orders": int(audit_df["Pcs/CFC Source"].astype(str).str.startswith("Pending").sum()) if len(audit_df) else 0,
        "export_rows_without_stock": unused_export,
        "missing_bom_items": missing_bom_items,
    }
    return Result(rows, export_rows, audit_df, warnings, summary, merges, conversion)


NOT_IN_BOM = "Not available in BOM"


def _denominator(item, bidx: BomIndex, mjp_m4, pending, pending_products):
    """Pcs per CFC from the uploaded BOM only. Nothing is taken from older reports."""
    bom_item = bidx.resolve(item, mjp_m4)
    if bom_item is None:
        return None, NOT_IN_BOM, ""
    cands = bidx.candidates(bom_item)
    if not cands:
        return None, NOT_IN_BOM + " (invalid quantity)", ""
    pend = {fg: d for fg, d in cands.items() if fg in pending_products}
    if pend:
        fg = max(pend, key=lambda f: (pending[f], pend[f]))
        return pend[fg], "Pending order", fg
    cnt = Counter(cands.values())
    best = max(cnt.items(), key=lambda kv: (kv[1], -kv[0]))[0]   # most common; tie -> smaller pack
    fg = next(f for f, d in cands.items() if d == best)
    return best, "BOM (most common pack)", fg


def _order_rows(rows: list[Row]) -> list[Row]:
    """Template rows keep their order. Each new row is placed before the first
    row of the same block (ENV... / TAG...) whose label sorts after it."""
    ordered = [r for r in rows if not r.is_new]
    for r in sorted((r for r in rows if r.is_new), key=lambda r: r.key):
        prefix = r.key.split(" ", 1)[0]
        block = [i for i, x in enumerate(ordered) if x.key.split(" ", 1)[0] == prefix]
        at = next((i for i in block if ordered[i].key > r.key), None)
        if at is None:
            at = (block[-1] + 1) if block else len(ordered)
        ordered.insert(at, r)
    return ordered


def _plan_merges(rows: list[Row], template_merges: dict, stock_cols: dict) -> dict:
    """Shared tags are written once and merged down over the rows that use them.

    Template merges are kept by row label. New rows that land between the
    block's rows are included so the block is not split.
    """
    pos = {r.key: i for i, r in enumerate(rows)}
    out: dict[tuple[str, str], list[str]] = {}
    allowed = set(stock_cols) | {s["cases_col"] for s in stock_cols.values()}
    for (anchor, col), labels in template_merges.items():
        if col not in allowed or anchor not in pos:
            continue
        span_keys = set(labels)
        span = [anchor]
        for r in rows[pos[anchor] + 1:]:
            if r.key in span_keys or r.is_new:
                span.append(r.key)
            else:
                break
        while len(span) > 1 and span[-1] not in span_keys:
            span.pop()
        if len(span) > 1:
            out[(anchor, col)] = span
    return out
