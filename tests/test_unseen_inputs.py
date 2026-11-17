"""Robustness: run the generator on inputs it has never seen.

Builds variations of a real Stock Summary / Order Status / BOM and checks the
report is still produced correctly. Needs real sample files:
    TAG_ENV_STOCK, TAG_ENV_ORDERS, TAG_ENV_BOM  (environment variables)
"""
import io
import os
import re
import sys
from pathlib import Path

import openpyxl
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from tag_env import pipeline, readers  # noqa: E402

STOCK, ORDERS, BOM = (os.environ[k] for k in ("TAG_ENV_STOCK", "TAG_ENV_ORDERS", "TAG_ENV_BOM"))
BOM_DF, _ = readers.read_exploded_bom(BOM)
results = []


def xlsx(rows, title="Stock SummaryNov 20,2026", pre=None):
    wb = openpyxl.Workbook(); ws = wb.active; ws.title = title
    for r in (pre or []):
        ws.append(r)
    for r in rows:
        ws.append(r)
    b = io.BytesIO(); wb.save(b); return b.getvalue()


def stock_rows():
    raw = pd.read_excel(STOCK, header=None)
    return raw.values.tolist()


def main_cells(xbytes):
    ws = openpyxl.load_workbook(io.BytesIO(xbytes))["Env & Tag Stock"]
    out = {}
    for r in range(5, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a:
            out[str(a).strip()] = {c: ws[f"{c}{r}"].value for c in
                                   ["B", "C", "D", "E", "H", "I", "J", "K", "N", "O", "R", "S", "X", "Y", "AB", "AC", "AD", "AE", "AH", "AI"]}
    return out, ws


def case(name, fn):
    try:
        ok, detail = fn()
    except Exception as e:  # noqa: BLE001
        ok, detail = False, f"crashed: {type(e).__name__}: {e}"
    results.append((name, "PASS" if ok else "FAIL", detail))


# 1. brand-new brand: envelope + tag on C250, envelope on Constanta, tag must join the C250 row
def new_brand():
    rows = stock_rows()
    extra = [["ENV ZEBRA DC SF C250 TBGS", "CONSTANTA-250", 500.0, 1],
             ["TAG ZEBRA DC SF C250 TBGS", "Constanta Tag", 80.0, 1],
             ["ENV ZEBRA MINT DC ENV TBGS", "Constanta Envelope", 120.5, 1]]
    data = xlsx(rows[:2] + extra + rows[2:])
    rep = pipeline.generate(data, ORDERS, None, bom_df=BOM_DF)
    cells, _ = main_cells(rep.xlsx)
    z, zm = cells.get("ENV ZEBRA"), cells.get("ENV ZEBRA MINT")
    ok = bool(z and zm and z["D"] == 500 and z["B"] == 80 and zm["N"] == 120.5)
    ws = openpyxl.load_workbook(io.BytesIO(rep.xlsx))["Env & Tag Stock"]
    no_pcs = (z is not None and z["C"] == "N/A" and z["E"] == "N/A" and zm["O"] == "N/A"
              and rep.summary["missing_bom_items"]["Item"].str.contains("ZEBRA").sum() == 3)
    return ok and no_pcs, f"rows created: ZEBRA={bool(z)}, ZEBRA MINT={bool(zm)}; tag joined envelope row={bool(z and z['B']==80)}; Make Cases N/A + listed for new BOM={no_pcs}"


# 2. header row moved down, extra blank/junk rows, columns reordered
def moved_header():
    rows = stock_rows()
    hdr_i = next(i for i, r in enumerate(rows) if "Item Name" in [str(x) for x in r])
    body = [[r[2], r[0], r[1], r[3]] for r in rows[hdr_i:]]  # Qty, Item, ML, Amount
    data = xlsx(body, pre=[["MJIL-COIMBATORE-EOU", None, "Stock Register: [Report on: 2026-11-20]"], [], ["some note"]])
    rep = pipeline.generate(data, ORDERS, None, bom_df=BOM_DF)
    base = pipeline.generate(STOCK, ORDERS, None, bom_df=BOM_DF)
    a, _ = main_cells(rep.xlsx); b, _ = main_cells(base.xlsx)
    same = all(a[k][c] == b[k][c] for k in b for c in ("B", "D", "H", "J", "N", "R", "X", "AB", "AD", "AH"))
    return same and rep.report_date.day == 20, f"same stock as original layout={same}; date read={rep.report_date:%d-%m-%Y}"


# 3. no pending orders at all -> Pcs/CFC from BOM only, no crash
def no_orders():
    wb = openpyxl.load_workbook(ORDERS)
    ws = wb["Order Status-By Buyer"]
    hdr = next(r for r in range(1, 20) if any(str(c.value).strip() == "Pending Ord Qty" for c in ws[r]))
    col = next(c.column for c in ws[hdr] if str(c.value).strip() == "Pending Ord Qty")
    for r in range(hdr + 1, ws.max_row + 1):
        if isinstance(ws.cell(r, col).value, (int, float)):
            ws.cell(r, col).value = 0
    b = io.BytesIO(); wb.save(b)
    rep = pipeline.generate(STOCK, b.getvalue(), None, bom_df=BOM_DF)
    src = rep.audit["Pcs/CFC Source"].astype(str)
    return (not src.str.startswith("Pending").any() and (rep.checks["Result"] == "PASS").all(),
            f"cells from orders={int(src.str.startswith('Pending').sum())}; format checks pass={(rep.checks['Result']=='PASS').all()}")


# 4. items removed from stock -> their old values must not survive from the template
def stale_values():
    rows = stock_rows()
    keep = [r for r in rows if not str(r[0]).startswith("ENV KHAIMA")]
    rep = pipeline.generate(xlsx(keep), ORDERS, None, bom_df=BOM_DF)
    cells, _ = main_cells(rep.xlsx)
    left = {k: v["N"] for k, v in cells.items() if k.startswith("ENV KHAIMA") and v["N"]}
    return not left, f"KHAIMA rows with leftover stock: {left or 'none'}"


# 5. duplicate stock lines are added together
def duplicates():
    rows = stock_rows()
    i = next(i for i, r in enumerate(rows) if str(r[0]) == "ENV GV MINT DC TBGS")
    rep = pipeline.generate(xlsx(rows + [rows[i]]), ORDERS, None, bom_df=BOM_DF)
    cells, _ = main_cells(rep.xlsx)
    exp = round(2 * float(rows[i][2]), 2)
    return abs(cells["ENV GV MINT"]["N"] - exp) < 0.01 and any("more than once" in w for w in rep.input_warnings), \
        f"GV MINT = {cells['ENV GV MINT']['N']} (expected {exp}); duplicate warning shown"


# 6. wrong files in each slot give clear errors, not crashes
def wrong_files():
    msgs = []
    for args in [(ORDERS, ORDERS, BOM), (STOCK, STOCK, BOM), (STOCK, ORDERS, STOCK)]:
        try:
            pipeline.generate(*args)
            msgs.append("no error")
        except readers.InputError as e:
            msgs.append(str(e).split(":")[0])
    return msgs == ["Stock Summary Report", "Order Status Report", "Exploded BOM"], " / ".join(msgs)


# 7. Make Cases formula = Kg x own conversion cell / Pcs per CFC on every stocked cell
def formulas():
    rep = pipeline.generate(STOCK, ORDERS, None, bom_df=BOM_DF)
    _, ws = main_cells(rep.xlsx)
    pairs = {"B": "C", "D": "E", "H": "I", "J": "K", "N": "O", "R": "S", "V": "W", "X": "Y", "AB": "AC", "AD": "AE", "AH": "AI"}
    bad = []
    for r in range(5, ws.max_row):
        for s, c in pairs.items():
            f = ws[f"{c}{r}"].value
            if isinstance(f, str) and f.startswith("="):
                if not re.fullmatch(rf"=\({s}{r}\*\${c}\$4\)/\d+(\.\d+)?", f):
                    bad.append(f"{c}{r}: {f}")
    return not bad, f"{len(bad)} bad formulas" + (f": {bad[:3]}" if bad else "")


for n, f in [("New brand never seen before", new_brand), ("Stock file with moved header / reordered columns", moved_header),
             ("Order file with no pending orders", no_orders), ("Items gone from stock leave no old values", stale_values),
             ("Duplicate stock lines", duplicates), ("Wrong file uploaded in each slot", wrong_files),
             ("Every Make Cases formula uses its own conversion cell", formulas)]:
    case(n, f)
for r in results:
    print(f"{r[1]:4} | {r[0]} | {r[2]}")
