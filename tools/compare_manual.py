"""Compare a generated report with a manually prepared one: stock per row/column
and the Pcs-per-CFC denominator of every non-zero cell."""
import sys, re, json
from pathlib import Path
import openpyxl
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from tag_env.readers import clean_name
cfg = json.loads((Path(__file__).resolve().parents[1] / "config/settings.json").read_text())
sc = cfg["main_sheet"]["stock_columns"]
def grab(p):
    ws = openpyxl.load_workbook(p)[cfg["main_sheet"]["name"]]
    out = {}
    for r in range(5, ws.max_row + 1):
        k = clean_name(ws.cell(r, 1).value).upper()
        if not k: continue
        for col, s in sc.items():
            v = ws[f"{col}{r}"].value
            m = re.search(r"/\s*(\d+(?:\.\d+)?)\s*$", str(ws[f"{s['cases_col']}{r}"].value or ""))
            out[(k, col)] = (v if isinstance(v, (int, float)) else 0, float(m.group(1)) if m else None)
    return out
gen, man = grab(sys.argv[1]), grab(sys.argv[2])
qty_ok = qty_bad = den_ok = den_bad = 0
for key, (mv, md) in man.items():
    gv, gd = gen.get(key, (0, None))
    if abs((gv or 0) - (mv or 0)) > 0.01:
        qty_bad += 1; print("QTY ", key, "generated", gv, "manual", mv)
    elif mv:
        qty_ok += 1
        if gd == md: den_ok += 1
        else: den_bad += 1; print("DEN ", key, "generated", gd, "manual", md)
extra = [(k, v) for k, v in gen.items() if k not in man and v[0]]
for k, v in extra: print("ONLY IN GENERATED", k, v)
print(f"stock cells matching: {qty_ok}, different: {qty_bad}; denominators matching: {den_ok}, different: {den_bad}")
