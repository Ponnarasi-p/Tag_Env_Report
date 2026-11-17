"""One call that turns the three uploads into the finished report."""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd

from . import engine, readers, writer
from .validate import compare_with_template


@dataclass
class Report:
    xlsx: bytes
    file_name: str
    preview: pd.DataFrame
    audit: pd.DataFrame
    warnings: list[str]
    summary: dict
    checks: pd.DataFrame
    report_date: datetime | None = None
    input_warnings: list[str] = field(default_factory=list)


def preview_frame(res: engine.Result, cfg: dict) -> pd.DataFrame:
    """Same numbers as the Excel formulas, for showing on screen."""
    main = cfg["main_sheet"]
    recs = []
    for r in res.rows:
        rec = {"Brand": r.label.strip()}
        cases = {}
        not_in_bom = []
        for col, spec in main["stock_columns"].items():
            q = r.qty.get(col, 0.0)
            d = r.denom.get(col)
            f = res.conversion.get(spec["cases_col"])
            cases[spec["cases_col"]] = (q * f / d) if (q and d and f) else 0.0
            name = f"{spec['group']} {'Tag' if spec['kind'] == 'TAG' else 'Env'}"
            if q and not d:
                not_in_bom.append(name)
            rec[f"{name} Kg"] = round(q, 2)
            rec[f"{name} Pcs/CFC"] = d if q else None
            rec[f"{name} Cases"] = round(cases[spec["cases_col"]], 1) if (q and d) else None
        mx = {m: max(cases.get(c, 0.0) for c in src) for m, src in main["max_columns"].items()}
        rec["Total Can Make"] = round(sum(mx[c] for c in main["total_can_make"]["sum_of"]), 1)
        rec["Remarks"] = r.remark
        rec["N/A (not in BOM)"] = ", ".join(not_in_bom)
        rec["New Row"] = "Yes" if r.is_new else ""
        recs.append(rec)
    df = pd.DataFrame(recs)
    # drop columns that are empty for every row so the preview stays readable
    keep = [c for c in df.columns if c in ("Brand", "Total Can Make", "Remarks", "N/A (not in BOM)", "New Row")
            or df[c].fillna(0).astype(float).abs().sum() > 0]
    return df[keep]


def generate(stock_src, order_src, bom_src, template_src=None, bom_df: pd.DataFrame | None = None,
             settings: dict | None = None) -> Report:
    cfg = settings or engine.load_settings()
    stock, report_date, w1 = readers.read_stock_summary(stock_src)
    orders, w2 = readers.read_order_status(order_src)
    if bom_df is None:
        bom_df, w3 = readers.read_exploded_bom(bom_src)
    else:
        w3 = []
    tinfo = writer.read_template(template_src, cfg)
    res = engine.build(stock, orders, bom_df, tinfo, cfg)
    xlsx, notes = writer.write_report(res, cfg, report_date, template_src)
    checks = compare_with_template(xlsx, template_src, cfg)
    date_txt = (report_date or datetime.now()).strftime("%d-%m-%y")
    return Report(
        xlsx=xlsx,
        file_name=f"Tag Env Stock as on {date_txt}.xlsx",
        preview=preview_frame(res, cfg),
        audit=res.audit,
        warnings=res.warnings + notes,
        summary=res.summary,
        checks=checks,
        report_date=report_date,
        input_warnings=w1 + w2 + w3,
    )
