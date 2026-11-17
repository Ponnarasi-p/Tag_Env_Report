"""Readers for the three input reports.

Each reader accepts bytes (from an upload) or a path, finds its header row by
content instead of by fixed position, validates the required columns and
returns a clean DataFrame. Problems are raised as InputError with a message a
user can act on.
"""
from __future__ import annotations

import io
import re
from datetime import datetime
from pathlib import Path

import pandas as pd

from .notes import Note


class InputError(ValueError):
    """Raised when an uploaded file cannot be used."""


def _as_bytes(src) -> bytes:
    if isinstance(src, (bytes, bytearray)):
        return bytes(src)
    if isinstance(src, (str, Path)):
        return Path(src).read_bytes()
    if hasattr(src, "getvalue"):
        return src.getvalue()
    return src.read()


def _engine(data: bytes) -> str | None:
    if data[:4] == b"PK\x03\x04":
        return "calamine_or_openpyxl"
    if data[:8] == bytes.fromhex("D0CF11E0A1B11AE1"):
        return "xlrd"
    return None


def _read_excel(data: bytes, label: str, **kw) -> dict[str, pd.DataFrame]:
    kind = _engine(data)
    if kind is None:
        # Some ERP ".xls" exports are HTML tables.
        try:
            tables = pd.read_html(io.BytesIO(data))
            return {f"Table{i+1}": t for i, t in enumerate(tables)}
        except Exception as exc:  # noqa: BLE001
            raise InputError(f"{label}: the file is not a readable Excel workbook.") from exc
    engines = ["xlrd"] if kind == "xlrd" else ["calamine", "openpyxl"]
    last = None
    for eng in engines:
        try:
            return pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, engine=eng, **kw)
        except ImportError as exc:
            last = exc
            continue
        except Exception as exc:  # noqa: BLE001
            last = exc
            continue
    raise InputError(f"{label}: could not open the workbook ({last}).")


def _find_header(df: pd.DataFrame, required: list[str], max_rows: int = 15) -> int | None:
    req = [r.lower() for r in required]
    for i in range(min(max_rows, len(df))):
        row = [str(v).strip().lower() for v in df.iloc[i].tolist()]
        if all(r in row for r in req):
            return i
    return None


def _frame_from(df: pd.DataFrame, header_idx: int) -> pd.DataFrame:
    cols = [str(c).strip() for c in df.iloc[header_idx].tolist()]
    # make duplicate column names unique (BOM has two "ITEM ID" columns)
    seen: dict[str, int] = {}
    uniq = []
    for c in cols:
        if c in seen:
            seen[c] += 1
            uniq.append(f"{c}.{seen[c]}")
        else:
            seen[c] = 0
            uniq.append(c)
    out = df.iloc[header_idx + 1:].copy()
    out.columns = uniq
    return out.reset_index(drop=True)


def clean_name(v) -> str:
    """Item names: trim, collapse inner whitespace/newlines, keep case."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    return re.sub(r"\s+", " ", str(v)).strip()


# --------------------------------------------------------------------------
# Stock Summary Report
# --------------------------------------------------------------------------
STOCK_REQUIRED = ["Item Name", "Machine Line Name", "Qty"]


def read_stock_summary(src) -> tuple[pd.DataFrame, datetime | None, list[str]]:
    """Returns (stock df[item, machine_line, qty], report date, warnings)."""
    data = _as_bytes(src)
    sheets = _read_excel(data, "Stock Summary Report")
    warnings: list[str] = []
    for name, raw in sheets.items():
        h = _find_header(raw, STOCK_REQUIRED)
        if h is None:
            continue
        report_date = None
        for i in range(h):
            for v in raw.iloc[i].tolist():
                m = re.search(r"(\d{4}-\d{2}-\d{2})", str(v))
                if m:
                    report_date = datetime.strptime(m.group(1), "%Y-%m-%d")
        df = _frame_from(raw, h)
        df = df.rename(columns={"Item Name": "item", "Machine Line Name": "machine_line", "Qty": "qty"})
        df["item"] = df["item"].map(clean_name)
        df["machine_line"] = df["machine_line"].map(clean_name)
        df = df[df["item"] != ""]
        bad_qty = pd.to_numeric(df["qty"], errors="coerce")
        bad = df[bad_qty.isna()]
        if len(bad):
            warnings.append(Note(
                f"Stock Summary: {len(bad)} item(s) have an empty or non-numeric Qty, treated as 0",
                [(r.item, r.machine_line, "" if pd.isna(r.qty) else str(r.qty)) for r in bad.itertuples()],
                ["Item", "Machine Line", "Qty in file"]))
        df["qty"] = bad_qty.fillna(0.0).astype(float)
        neg = df[df["qty"] < 0]
        if len(neg):
            warnings.append(Note(f"Stock Summary: {len(neg)} item(s) have negative stock",
                                 neg[["item", "machine_line", "qty"]].set_axis(["Item", "Machine Line", "Stock Kg"], axis=1)))
        dup = df[df.duplicated("item", keep=False)]
        if len(dup):
            g = dup.groupby(["item", "machine_line"], sort=False)["qty"].agg(times="size", total="sum").reset_index()
            warnings.append(Note(
                f"Stock Summary: {dup['item'].nunique()} item(s) appear more than once; their quantities were added together",
                [(it, ml, int(n), float(t)) for it, ml, n, t in g[["item", "machine_line", "times", "total"]].values],
                ["Item", "Machine Line", "Times listed", "Total Kg used"]))
            df = df.groupby(["item", "machine_line"], as_index=False, sort=False)["qty"].sum()
        if report_date is None:
            warnings.append(Note("Stock Summary: report date not found in the file header; today's date was used"))
        return df[["item", "machine_line", "qty"]].reset_index(drop=True), report_date, warnings
    raise InputError("Stock Summary Report: could not find the header row with columns "
                     + ", ".join(STOCK_REQUIRED) + ". Is this the right file?")


# --------------------------------------------------------------------------
# Order Status Report
# --------------------------------------------------------------------------
ORDER_SHEETS = ["Order Status-By Buyer", "Order Status-By shipment Date"]


def read_order_status(src) -> tuple[pd.DataFrame, list[str]]:
    """Returns (orders df[product, pending_qty, ...], warnings). One row per order line."""
    data = _as_bytes(src)
    sheets = _read_excel(data, "Order Status Report")
    warnings: list[str] = []
    ordered = [s for s in ORDER_SHEETS if s in sheets] + [s for s in sheets if s not in ORDER_SHEETS]
    for name in ordered:
        raw = sheets[name]
        h = _find_header(raw, ["Prod Name", "Pending Ord Qty"], max_rows=20)
        if h is None:
            continue
        df = _frame_from(raw, h)
        keep = {"Prod Name": "product", "Pending Ord Qty": "pending_qty", "Prod Id": "prod_id",
                "Ord Id": "order_id", "Buyer Name": "buyer", "Order Status": "order_status"}
        df = df.rename(columns={k: v for k, v in keep.items() if k in df.columns})
        df["order_id"] = df.get("order_id", pd.Series(index=df.index)).replace(r"^\s*$", pd.NA, regex=True).ffill()
        if "buyer" in df:
            df["buyer"] = df["buyer"].replace(r"^\s*$", pd.NA, regex=True).ffill()
        df["product"] = df["product"].map(clean_name)
        df = df[df["product"] != ""]
        df["pending_qty"] = pd.to_numeric(df["pending_qty"], errors="coerce").fillna(0)
        df = df[df["pending_qty"] > 0]
        if df.empty:
            warnings.append(Note(f"Order Status (sheet '{name}'): no order lines have a pending quantity, "
                                 "so no remark can be 'Order Received' and Pcs per CFC comes from the BOM only"))
        cols = [c for c in ["order_id", "buyer", "prod_id", "product", "pending_qty", "order_status"] if c in df]
        return df[cols].reset_index(drop=True), warnings
    raise InputError("Order Status Report: no sheet has the columns 'Prod Name' and 'Pending Ord Qty'. "
                     "Expected a sheet named '" + ORDER_SHEETS[0] + "'.")


# --------------------------------------------------------------------------
# Exploded BOM
# --------------------------------------------------------------------------
BOM_REQUIRED = ["FG Name", "Item Name", "ITEM TYPE", "Qty", "FG Uom"]


def read_exploded_bom(src) -> tuple[pd.DataFrame, list[str]]:
    """Returns (bom df[fg, item, item_type, qty, uom, fg_uom], warnings)."""
    data = _as_bytes(src)
    warnings: list[str] = []
    sheets = _read_excel(data, "Exploded BOM", nrows=None)
    for name, raw in sheets.items():
        h = _find_header(raw, BOM_REQUIRED)
        if h is None:
            continue
        df = _frame_from(raw, h)
        df = df.rename(columns={"FG Name": "fg", "Item Name": "item", "ITEM TYPE": "item_type",
                                "Qty": "qty", "Uom": "uom", "FG Uom": "fg_uom"})
        for c in ["fg", "item", "item_type", "fg_uom"]:
            df[c] = df[c].map(clean_name)
        df["item_type"] = df["item_type"].str.upper()
        df["fg_uom"] = df["fg_uom"].str.upper()
        raw_qty = df["qty"]
        df["qty"] = pd.to_numeric(df["qty"], errors="coerce")
        df = df[(df["fg"] != "") & (df["item"] != "")]
        bad = df[df["qty"].isna()]
        if len(bad):
            warnings.append(Note(
                f"Exploded BOM: {len(bad)} line(s) have no numeric Qty and were skipped",
                [(r.fg, r.item, r.item_type, str(raw_qty.get(i, ""))) for i, r in bad.iterrows()],
                ["FG Name", "Item Name", "Item Type", "Qty in file"]))
        df = df.dropna(subset=["qty"])
        dup = df.duplicated(["fg", "item"], keep="first")
        if dup.any():
            kept = df[~dup].set_index(["fg", "item"])["qty"]
            warnings.append(Note(
                f"Exploded BOM (sheet '{name}'): {int(dup.sum())} FG/item line(s) are listed twice; "
                "the first line was used and the repeat ignored",
                [(r.fg, r.item, r.item_type, float(kept.get((r.fg, r.item))), float(r.qty))
                 for r in df[dup].itertuples()],
                ["FG Name", "Item Name", "Item Type", "Qty used", "Repeated Qty ignored"]))
            df = df[~dup]
        return df[["fg", "item", "item_type", "qty", "uom", "fg_uom"]].reset_index(drop=True), warnings
    raise InputError("Exploded BOM: could not find the header row with columns " + ", ".join(BOM_REQUIRED) + ".")


# --------------------------------------------------------------------------
# Optional MJP vs M4 item-name mapping
# --------------------------------------------------------------------------
def read_item_mapping(src) -> dict[str, str]:
    """MJP (stock) item name -> M4 (BOM) item name."""
    data = _as_bytes(src)
    sheets = _read_excel(data, "Item mapping")
    for raw in sheets.values():
        h = _find_header(raw, ["M4 Item Name", "MJP Item Name"])
        if h is None:
            continue
        df = _frame_from(raw, h)
        out = {}
        for m4, mjp in zip(df["M4 Item Name"], df["MJP Item Name"]):
            m4, mjp = clean_name(m4), clean_name(mjp)
            if m4 and mjp:
                out.setdefault(mjp, m4)
        return out
    raise InputError("Item mapping: expected columns 'M4 Item Name' and 'MJP Item Name'.")
