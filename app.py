"""Tag Env Stock Report Generator (Streamlit)."""
from __future__ import annotations

import hashlib
import io
import json
import traceback

import pandas as pd
import streamlit as st

from tag_env import engine, pipeline, readers
from tag_env.notes import as_note

st.set_page_config(page_title="Tag Env Stock Report", page_icon="📦", layout="wide")


@st.cache_data(show_spinner=False, max_entries=3)
def _bom(data: bytes, _digest: str):
    return readers.read_exploded_bom(data)


def _digest(b: bytes) -> str:
    return hashlib.md5(b).hexdigest()


def _check(label: str, upload, reader):
    """Validate one upload and show a short status line."""
    if upload is None:
        st.caption("Waiting for file")
        return None
    data = upload.getvalue()
    try:
        out = reader(data)
        if label == "stock":
            df, date, _ = out
            st.success(f"{len(df)} stock lines" + (f" · report date {date:%d-%m-%Y}" if date else ""))
        elif label == "orders":
            df, _ = out
            st.success(f"{len(df)} pending order lines · {df['product'].nunique()} products")
        else:
            df, _ = out
            st.success(f"{df['fg'].nunique()} finished goods · {len(df)} BOM lines")
        return data
    except readers.InputError as e:
        st.error(str(e))
    except Exception as e:  # noqa: BLE001
        st.error(f"Could not read this file: {e}")
    return None


st.title("Tag Env Stock Report")
st.write("Upload the three reports and generate the Tag & Envelope stock report with Make Cases, "
         "in the same layout as the manual Excel report.")

c1, c2, c3 = st.columns(3)
with c1:
    st.subheader("1. Stock Summary")
    up_stock = st.file_uploader("Stock Summary Report (.xls / .xlsx)", type=["xls", "xlsx"], key="stock")
    stock_bytes = _check("stock", up_stock, readers.read_stock_summary)
with c2:
    st.subheader("2. Order Status")
    up_orders = st.file_uploader("Order Status Report (.xlsx)", type=["xlsx", "xls"], key="orders")
    order_bytes = _check("orders", up_orders, readers.read_order_status)
with c3:
    st.subheader("3. Exploded BOM")
    up_bom = st.file_uploader("Exploded BOM Report (.xlsx)", type=["xlsx", "xls"], key="bom")
    bom_bytes = None
    if up_bom is not None:
        data = up_bom.getvalue()
        try:
            with st.spinner("Reading BOM..."):
                bdf, _ = _bom(data, _digest(data))
            st.success(f"{bdf['fg'].nunique()} finished goods · {len(bdf)} BOM lines")
            bom_bytes = data
        except readers.InputError as e:
            st.error(str(e))
        except Exception as e:  # noqa: BLE001
            st.error(f"Could not read this file: {e}")
    else:
        st.caption("Waiting for file")

with st.expander("Settings: conversion factors and report template"):
    cfg = engine.load_settings()
    st.write("Conversion factors (Pcs per Kg) used for each Make Cases column. "
             "Edit `config/settings.json` to change them permanently.")
    conv_rows = [{"Report column": f"{s['group']} {'Tag' if s['kind'] == 'TAG' else 'Envelope'}",
                  "Machine line": s["conversion"],
                  "Pcs per Kg": cfg["conversion_factors"].get(s["conversion"])}
                 for s in cfg["main_sheet"]["stock_columns"].values()]
    edited = st.data_editor(pd.DataFrame(conv_rows), disabled=["Report column", "Machine line"],
                            hide_index=True, width="stretch", key="conv")
    for _, r in edited.iterrows():
        if pd.notna(r["Pcs per Kg"]):
            cfg["conversion_factors"][r["Machine line"]] = float(r["Pcs per Kg"])
    up_tmpl = st.file_uploader("Reference report to use as template (optional; default: built-in "
                               "'Tag Env Stock as on 28-10-26')", type=["xlsx"], key="tmpl")

ready = all(x is not None for x in (stock_bytes, order_bytes, bom_bytes))
go = st.button("Generate Report", type="primary", disabled=not ready, width="stretch")
if not ready:
    st.info("Upload all three files to enable **Generate Report**.")

if go:
    try:
        with st.status("Generating report...", expanded=True) as status:
            st.write("Reading inputs and matching items to report rows")
            bdf, bw = _bom(bom_bytes, _digest(bom_bytes))
            tmpl = up_tmpl.getvalue() if up_tmpl is not None else None
            st.write("Calculating Pcs per CFC and Make Cases")
            rep = pipeline.generate(stock_bytes, order_bytes, None, tmpl, bom_df=bdf, settings=cfg)
            rep.input_warnings = rep.input_warnings + bw
            st.write("Writing Excel in the template layout and checking the format")
            status.update(label="Report ready", state="complete", expanded=False)
        st.session_state["report"] = rep
    except readers.InputError as e:
        st.error(str(e))
    except Exception as e:  # noqa: BLE001
        st.error(f"Report generation failed: {e}")
        with st.expander("Technical details"):
            st.code(traceback.format_exc())

rep: pipeline.Report | None = st.session_state.get("report")
if rep is not None:
    st.divider()
    left, right = st.columns([3, 1])
    with left:
        st.subheader(rep.file_name.removesuffix(".xlsx"))
    with right:
        st.download_button("Download Excel report", rep.xlsx, file_name=rep.file_name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           type="primary", width="stretch")

    s = rep.summary
    m = st.columns(5)
    m[0].metric("Stock lines", s["stock_lines"])
    m[1].metric("Placed on report", s["placed_main"] + s["placed_export"])
    m[2].metric("Report rows", s["rows"], delta=f"{s['new_rows']} new" if s["new_rows"] else None,
                delta_color="off")
    m[3].metric("Pcs/CFC from pending orders", s["cells_from_orders"])
    m[4].metric("Ordered products found in BOM", f"{s['pending_products_in_bom']} / {s['pending_products']}")

    fails = rep.checks[rep.checks["Result"] == "FAIL"]
    if fails.empty:
        st.success(f"Format check passed: all {len(rep.checks)} structure and formatting checks "
                   "match the reference template.")
    else:
        st.error(f"Format check: {len(fails)} of {len(rep.checks)} checks differ from the template.")

    missing = s.get("missing_bom_items")
    if missing is not None and len(missing):
        st.error(f"**{len(missing)} item(s) are not available in the Exploded BOM.** Their Make Cases "
                 "shows **N/A** in the report. Please submit a new Exploded BOM that includes these "
                 "items and generate the report again.")
        st.dataframe(missing, width="stretch", hide_index=True,
                     height=min(600, 38 + 35 * len(missing)))

    # every warning: title + the full list of affected items
    all_w = [as_note(w) for w in rep.input_warnings + rep.warnings]
    other_w = [w for w in all_w if "not available in the Exploded BOM" not in w.title]
    if other_w:
        st.subheader(f"Warnings and processing notes ({len(other_w)})")
        for i, w in enumerate(other_w, start=1):
            n = len(w.table)
            with st.expander(f"⚠️ {i}. {w.title}" + (f"  · {n} line{'s' if n != 1 else ''} below" if n else ""),
                             expanded=False):
                if n:
                    st.dataframe(w.table, width="stretch", hide_index=True,
                                 height=min(600, 38 + 35 * n))
                else:
                    st.write(w.title)
    if all_w:
        wbuf = io.BytesIO()
        with pd.ExcelWriter(wbuf, engine="openpyxl") as xw:
            pd.DataFrame({"#": range(1, len(all_w) + 1), "Warning": [w.title for w in all_w],
                          "Items": [len(w.table) for w in all_w]}).to_excel(xw, sheet_name="Summary", index=False)
            for i, w in enumerate(all_w, start=1):
                if len(w.table):
                    w.table.to_excel(xw, sheet_name=f"Warning {i}", index=False)
        st.download_button("Download all warnings (Excel)", wbuf.getvalue(),
                           file_name=rep.file_name.replace(".xlsx", " - warnings.xlsx"))

    t1, t2, t3 = st.tabs(["Report preview", "Item calculation detail", "Format check"])
    with t1:
        st.caption("Main sheet 'Env & Tag Stock'. Columns with no stock are hidden here; "
                   "the Excel file keeps the full layout.")
        st.dataframe(rep.preview, width="stretch", hide_index=True, height=560)
    with t2:
        st.caption("One line per stock item: where it was placed, which FG and pending order set its "
                   "Pcs per CFC, and the resulting Make Cases.")
        st.dataframe(rep.audit, width="stretch", hide_index=True, height=560)
        buf = io.BytesIO()
        rep.audit.to_excel(buf, index=False, sheet_name="Item Detail")
        st.download_button("Download calculation detail", buf.getvalue(),
                           file_name=rep.file_name.replace(".xlsx", " - calculation detail.xlsx"))
    with t3:
        st.dataframe(rep.checks, width="stretch", hide_index=True)
