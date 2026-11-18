"""Business logic: stock items -> report rows/columns, Pcs-per-CFC lookup and
Make Cases calculation.

Make Cases = (Available KG x Conversion factor of the column's machine line)
             / Pcs per CFC

Pcs per CFC is the pack size in the name of the finished good (FG) the item is
used in: 'GV MINT 48/25 DC ENV TBGS' -> 48 x 25 = 1200. The FG is found through
the Exploded BOM, the MJP vs M4 item-name mapping, or by name against the BOM
and Order Status FG names (see FgLinker).

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

from .notes import Note
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
        self.anomalies: set[tuple[str, str, float]] = set()

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
                self.anomalies.add((fg, bom_item, float(q)))
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
    no_conv = []
    for col, spec in stock_cols.items():
        f = cfg["conversion_factors"].get(spec["conversion"])
        if f is None:
            no_conv.append((f"{spec['group']} {'Tag' if spec['kind'] == 'TAG' else 'Envelope'}",
                            spec["cases_col"], spec["conversion"]))
        conversion[spec["cases_col"]] = f
    if no_conv:
        warnings.append(Note("No conversion factor is set for these report columns; their Make Cases is left empty. "
                             "Add the factor in Settings or config/settings.json",
                             no_conv, ["Report column", "Excel column", "Machine line needed"]))

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
    unmatched_items, skipped, new_rows, export_missing = [], [], {}, []

    # ---- place every stock item --------------------------------------------
    for _, s in stock.iterrows():
        item, ml, qty = s["item"], s["machine_line"], float(s["qty"])
        col = ml2col.get(ml)
        if col is None:
            skipped.append((item, ml or "(blank)", qty))
            continue
        up = item.upper()
        mapped = item_row_map.get(item)
        # export (Ghana / Abidjan / Angola) items
        if (mapped and mapped[0] == "export") or any(m in up for m in markers):
            key = (mapped[1] if mapped else item).upper()
            er = export_rows.get(clean_name(key).upper())
            kind = item_kind(item, bidx.type_of.get(bidx.resolve(item, mjp_m4) or ""))
            if er is None:
                export_missing.append((item, ml, qty))
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
                new_rows[row.key] = row.label.strip()
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
            new_rows[row.key] = label.strip()
            unmatched_items.append(item)
            how = "new row"
        row.items[col].append(item)
        row.qty[col] += qty
        audit.append({"Item": item, "Machine Line": ml, "Stock Kg": qty, "Sheet": main["name"],
                      "Row": row.label.strip(), "Column": col, "Matched By": how})

    # ---- denominators -------------------------------------------------------
    linker = FgLinker(bidx, mjp_m4, pending_products)
    item_denoms: dict[str, tuple[float | None, str, str]] = {}
    for row in rows:
        for col, items in row.items.items():
            if not row.qty.get(col):
                continue
            # use the item with the most stock in the cell
            items_sorted = sorted(items, key=lambda it: -float(stock.loc[stock["item"] == it, "qty"].sum()))
            chosen = None
            for it in items_sorted:
                d, src, fg = _denominator(it, row.label, linker, pending, pending_products)
                item_denoms[it] = (d, src, fg)
                if chosen is None or (chosen[0] is None and d is not None):
                    chosen = (d, src, fg)
            row.denom[col] = chosen[0]
            row.denom_source[col] = chosen[1] + (f" [{chosen[2]}]" if chosen[2] else "")
        # an item with no FG of its own takes the pack size of the same row's other
        # envelope/tag (same brand and flavour, e.g. JUMBO MANGO on MD-20 vs C250)
        known = {c: (row.denom[c], row.denom_source.get(c, "")) for c in row.denom if row.denom.get(c)}
        for col in [c for c in row.denom if not row.denom.get(c)]:
            same_kind = [c for c in known if stock_cols[c]["kind"] == stock_cols[col]["kind"]]
            pick = (same_kind or list(known))[:1]
            if pick:
                d, srcd = known[pick[0]]
                fg = srcd[srcd.find("[") + 1:-1] if "[" in srcd else ""
                row.denom[col] = d
                row.denom_source[col] = f"Same row ({stock_cols[pick[0]]['group']}) [{fg}]"
                for it in row.items[col]:
                    item_denoms[it] = (d, f"Same row as {stock_cols[pick[0]]['group']} "
                                          f"{'tag' if stock_cols[pick[0]]['kind'] == 'TAG' else 'envelope'}", fg)
    # cells without stock have no item, so no Pcs per CFC (Make Cases is 0)

    # ---- remarks: Order Received / No Order --------------------------------
    order_by_row = _order_remarks(rows, bidx, mjp_m4, pending_products)

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
        rk = clean_name(a["Row"]).upper()
        a["Remark"], a["Ordered Product(s)"] = order_by_row.get(rk, (REMARK_NO_ORDER, ""))
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
        warnings.append(Note(
            f"{len(missing_bom_items)} item(s) are not available in the Exploded BOM, so their Make Cases "
            "shows N/A. Please submit a new Exploded BOM that includes them",
            missing_bom_items.rename(columns={"Row": "Report Row", "Pcs/CFC Source": "Reason"})))
    if export_missing:
        warnings.append(Note(
            f"{len(export_missing)} export item(s) have no row on the '{cfg['export_sheet']['name']}' sheet and "
            "are not on the report; add them to the template",
            export_missing, ["Item", "Machine Line", "Stock Kg"]))
    if new_rows and len(audit_df):
        nr = audit_df[audit_df["Row"].isin(set(new_rows.values()))]
        warnings.append(Note(
            f"{len(new_rows)} new row(s) were added for items that are not in the template",
            [(r["Row"], r["Item"], r["Machine Line"], r["Stock Kg"]) for _, r in nr.iterrows()],
            ["New Row", "Item", "Machine Line", "Stock Kg"]))
    if skipped:
        warnings.append(Note(
            f"{len(skipped)} stock line(s) were ignored because their machine line is not a Tag/Envelope line",
            sorted(skipped, key=lambda x: (x[1], x[0])), ["Item", "Machine Line", "Stock Kg"]))
    missing_fg = sorted(p for p in pending_products if p not in bidx.by_fg)
    if missing_fg:
        og = orders[orders["product"].isin(missing_fg)]
        agg = og.groupby("product").agg(pending=("pending_qty", "sum"),
                                        buyers=("buyer", lambda s: ", ".join(sorted(set(map(str, s.dropna())))))
                                        if "buyer" in og else ("pending_qty", "size"))
        warnings.append(Note(
            f"{len(missing_fg)} ordered product(s) are not in the Exploded BOM, so their orders cannot be "
            "linked to envelopes/tags (Pcs per CFC and remarks)",
            [(p, float(agg.loc[p, "pending"]), agg.loc[p, "buyers"]) for p in missing_fg],
            ["Ordered Product", "Pending Qty (CFC)", "Buyer(s)"]))
    if bidx.anomalies:
        warnings.append(Note(
            f"Exploded BOM has {len(bidx.anomalies)} tag/envelope line(s) under 1 piece per CFC; "
            "they were ignored as data errors",
            sorted(bidx.anomalies), ["FG Name", "Item Name", "Qty per CFC in BOM"]))
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
REMARK_ORDER = "Order Received"
REMARK_NO_ORDER = "No Order"
_PACK_RE = re.compile(r"\s+\d+\s*/\s*\d+.*$")   # " 24/25 DC SF ENV TBGS ..." after the brand


def _product_base(product: str) -> str:
    """'AFRICA CHOICE 24/25 DC ENV TBGS' -> 'AFRICA CHOICE' (normalised like row labels)."""
    return row_key(_PACK_RE.sub("", clean_name(product)))


def _order_remarks(rows: list[Row], bidx: BomIndex, mjp_m4, pending_products) -> dict:
    """Row key -> (remark, ordered products that gave it).

    A row is 'Order Received' when the Order Status report has a pending order for
    a finished good that (a) uses one of the row's envelopes/tags in the BOM, or
    (b) has the same brand name as the row (name before the pack size).
    Otherwise 'No Order'.
    """
    by_base: dict[str, list[str]] = defaultdict(list)
    for p in pending_products:
        by_base[_product_base(p)].append(p)
    out = {}
    for r in rows:
        hits = set()
        all_items = [it for its in r.items.values() for it in its]
        # envelopes identify the product; tags are often shared by many flavours,
        # so they are used only on rows that hold tags alone
        env_items = [it for it in all_items if item_kind(it, bidx.type_of.get(bidx.resolve(it, mjp_m4) or "")) == "ENV"]
        for it in (env_items or all_items):
            b = bidx.resolve(it, mjp_m4)
            if b:
                hits |= {fg for fg, _, _ in bidx.by_item.get(b, []) if fg in pending_products}
        hits |= set(by_base.get(row_key(r.label), []))
        r.remark = REMARK_ORDER if hits else REMARK_NO_ORDER
        r.stock_remark = r.remark
        out[r.key] = (r.remark, "; ".join(sorted(hits)[:5]) + (" ..." if len(hits) > 5 else ""))
    return out


_PACK_NUM_RE = re.compile(r"(?<![\d.])(\d+)\s*/\s*(\d+)(?:\s*/\s*(\d+))?(?![\d.])")


def pack_from_name(fg: str) -> float | None:
    """Pcs per CFC written in the FG name: 'GV MINT 48/25 DC ENV TBGS' -> 48 x 25 = 1200."""
    m = _PACK_NUM_RE.search(fg or "")
    if not m:
        return None
    v = 1
    for g in m.groups():
        if g:
            v *= int(g)
    return float(v)


class FgLinker:
    """Finds the finished goods (FG names) an envelope/tag belongs to.

    1. the item is in the BOM (directly, or via the MJP vs M4 name mapping);
    2. a BOM item with the same name once pack/machine words are removed
       (e.g. 'TAG PACIFICO DC SF TBGS' = BOM 'TAG PACIFICO DC TBGS');
    3. FG names (BOM + Order Status) whose brand/flavour before the pack size
       equals the item's (e.g. ENV FESTIVA GINGER -> 'FESTIVA GINGER 24/25 ...');
    4. the longest FG brand (2+ words) that the item name starts with
       (e.g. ALDI VRUCHTENTHEE RASPBERRY -> 'ALDI VRUCHTENTHEE 24/20 ...');
    5. FG names that start with the item's brand (ENV ELANI -> 'ELANI GREEN TEA 24/25 ...').
    """

    def __init__(self, bidx: BomIndex, mjp_m4: dict, order_products):
        self.bidx, self.mjp = bidx, mjp_m4
        self.by_key: dict[str, set[str]] = defaultdict(set)
        for it in bidx.items:
            self.by_key[row_key(it)].add(it)
        self.by_base: dict[str, set[str]] = defaultdict(set)
        for fg in set(bidx.by_fg) | set(order_products):
            self.by_base[_product_base(fg)].add(fg)

    def fgs(self, item: str, row_label: str) -> tuple[set[str], str]:
        b = self.bidx.resolve(item, self.mjp)
        if b:
            return {fg for fg, _, _ in self.bidx.by_item.get(b, [])}, "BOM"
        for key in filter(None, [row_key(item), row_key(self.mjp.get(item, ""))]):
            its = self.by_key.get(key)
            if its:
                return {fg for it in its for fg, _, _ in self.bidx.by_item.get(it, [])}, "BOM (name match)"
        for key in filter(None, [row_key(item), row_key(self.mjp.get(item, "")), row_key(row_label)]):
            if self.by_base.get(key):
                return set(self.by_base[key]), "FG name match"
        words = row_key(item).split()
        for n in range(len(words) - 1, 1, -1):
            if self.by_base.get(" ".join(words[:n])):
                return set(self.by_base[" ".join(words[:n])]), "FG name match (brand)"
        # 5. FG names that start with the item's brand (e.g. ELANI -> 'ELANI GREEN TEA 24/25 ...')
        key = row_key(item)
        if key:
            hits = {fg for base, fgs in self.by_base.items() if base.startswith(key + " ") for fg in fgs}
            if hits:
                return hits, "FG name match (starts with brand)"
        return set(), ""


def _denominator(item, row_label, linker: FgLinker, pending, pending_products):
    """Pcs per CFC = pack size in the chosen FG's name (n/m -> n x m).

    Only the uploaded BOM and Order Status are used. If the FG name has no
    n/m, the BOM's ENV/TAG pieces per CFC is used instead.
    """
    fgs, how = linker.fgs(item, row_label)
    cands = {}
    for fg in fgs:
        p = pack_from_name(fg)
        if p is None:
            t = linker.bidx.type_of.get(linker.bidx.resolve(item, linker.mjp) or "") or item_kind(item)
            tot = linker.bidx.total.get((fg, t))
            p = float(tot) if tot and tot >= 1 else None
        if p:
            cands[fg] = p
    if not cands:
        return None, NOT_IN_BOM if not fgs else NOT_IN_BOM + " (no pack size in FG name)", ""
    pend = {fg: d for fg, d in cands.items() if fg in pending_products}
    if pend:
        fg = max(pend, key=lambda f: (pending[f], pend[f]))
        return pend[fg], f"Pending order ({how})", fg
    cnt = Counter(cands.values())
    best = max(cnt.items(), key=lambda kv: (kv[1], -kv[0]))[0]   # most common; tie -> smaller pack
    fg = sorted(f for f, d in cands.items() if d == best)[0]
    return best, f"{how}, most common pack", fg


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
