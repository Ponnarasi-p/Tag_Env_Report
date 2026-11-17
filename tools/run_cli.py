"""Command-line run: python tools/run_cli.py STOCK ORDERS BOM OUT.xlsx [TEMPLATE]"""
import sys, time, pickle
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from tag_env import pipeline, readers

stock, orders, bom, out = sys.argv[1:5]
tmpl = Path(sys.argv[5]).read_bytes() if len(sys.argv) > 5 else None
t = time.time()
cache = Path(bom + ".pkl")
bom_df = None
if cache.exists():
    bom_df = pickle.loads(cache.read_bytes())
rep = pipeline.generate(stock, orders, bom, tmpl, bom_df=bom_df)
Path(out).write_bytes(rep.xlsx)
print(f"done in {time.time()-t:.1f}s -> {out}")
print("SUMMARY", {k: v for k, v in rep.summary.items() if k != "missing_bom_items"})
for w in rep.input_warnings + rep.warnings: print("WARN:", w)
print(rep.checks.to_string())
rep.audit.to_csv(out.replace(".xlsx", "_audit.csv"), index=False)
