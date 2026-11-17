"""Streamlit app: upload the aging stock report, download the corrected Excel."""
from __future__ import annotations

import hashlib
import re
from pathlib import Path
from datetime import date, datetime

import pandas as pd
import streamlit as st

from excel_reader import read_sheets
from excel_writer import build_workbook
from report_processor import Sheet, is_num, process_workbook

st.set_page_config(page_title="Report Correction Tool", page_icon="📊", layout="wide")

RULES = """
1. Rows that have an item name but no other details are removed. Rows with at least one value are kept.
2. **Day Consumption** is converted from tons to kg, and **No. of Days Consumption** is recalculated using kg.
3. **Raw Aging Stock**: Lot Code, Lot No, Supplier Lot No, Category, GSM, Qty and Machine Line are removed,
   rows are sorted by Holding Days (highest first) and the quantity heading shows "(kg)".
4. **All TagEnv Aging Stock**: Lot Code, Lot No, Supplier Lot No, Width, GSM and Source Type are removed.
   Repeated items are grouped with total QTN and average Holding Days, and the sheet is sorted by Holding Days (highest first). Click a grouped item in the
   preview (or the **[+]** in Excel's left margin) to see each record.
5. Decimal numbers are rounded to whole numbers.
6. Zero values are shown as "-" on every sheet.
7. Headers, alignment, column widths, borders and frozen header rows are applied to every sheet.
"""


def _rules_version() -> str:
    """Hash of the rule/IO modules, so cached results are dropped whenever the rules change."""
    h = hashlib.sha1()
    for mod in ("report_processor.py", "excel_reader.py", "excel_writer.py"):
        h.update((Path(__file__).parent / mod).read_bytes())
    return h.hexdigest()[:10]


RULES_VERSION = _rules_version()


@st.cache_data(show_spinner=False, max_entries=5)
def run(name: str, data: bytes, rules_version: str):
    sheets, log = process_workbook(read_sheets(name, data))
    return sheets, log, build_workbook(sheets)


def fmt(v) -> str:
    if v is None:
        return ""
    if isinstance(v, (datetime, date)):
        return v.strftime("%d-%b-%Y")
    if is_num(v):
        return f"{v:,}"
    return str(v)


def to_frame(s: Sheet, rows) -> pd.DataFrame:
    # Unique column labels are required by st.dataframe.
    seen: dict = {}
    cols = []
    for h in s.headers:
        label = h or " "
        seen[label] = seen.get(label, 0) + 1
        cols.append(label if seen[label] == 1 else f"{label} ({seen[label]})")
    return pd.DataFrame([[fmt(v) for v in r.cells[:len(cols)]] for r in rows], columns=cols)


def show_sheet(s: Sheet, idx: int) -> None:
    if not s.headers:
        st.info("Empty sheet")
        return
    main_rows = [r for r in s.rows if r.type != "detail"]
    has_groups = any(r.type == "summary" for r in main_rows)
    if not has_groups:
        st.caption(f"{len(s.rows)} row(s)")
        st.dataframe(to_frame(s, main_rows), hide_index=True, width="stretch")
        return

    n_items = sum(1 for r in main_rows if r.type != "total")
    st.caption(f"{n_items} unique item(s). Rows marked ▸ are repeated items: "
               "click one to see its individual records below.")
    df = to_frame(s, main_rows)
    for i, r in enumerate(main_rows):
        if r.type == "summary":
            df.iat[i, s.item_col] = f"▸ {df.iat[i, s.item_col]}  ({r.count})"
    event = st.dataframe(df, hide_index=True, width="stretch",
                         on_select="rerun", selection_mode="single-row", key=f"grid_{idx}")
    sel = event.selection.rows if event and event.selection else []
    if sel:
        picked = main_rows[sel[0]]
        if picked.type == "summary":
            details = [r for r in s.rows if r.type == "detail" and r.group_id == picked.group_id]
            st.markdown(f"**{fmt(picked.cells[s.item_col])}**: {len(details)} records")
            st.dataframe(to_frame(s, details), hide_index=True, width="stretch")
        else:
            st.caption("This item appears only once, so it has no other records to show.")


# ------------------------------------------------------------------ page
st.title("📊 Report Correction Tool")
st.write("Upload the downloaded aging stock report and get a cleaned, formatted Excel file back.")

uploaded = st.file_uploader("Excel report", type=["xlsx", "xlsm", "xls", "csv"], label_visibility="collapsed")

if uploaded is None:
    with st.container(border=True):
        st.subheader("Corrections applied automatically")
        st.markdown(RULES)
    st.stop()

try:
    with st.spinner(f"Processing {uploaded.name}…"):
        sheets, log, xlsx = run(uploaded.name, uploaded.getvalue(), RULES_VERSION)
except Exception as e:  # show a friendly message instead of a stack trace
    st.error(f"Could not process this file: {e}")
    st.stop()

out_name = re.sub(r"\.[^.]+$", "", uploaded.name) + " - Corrected.xlsx"
left, right = st.columns([3, 1], vertical_alignment="center")
left.success(f"Corrected report ready: {len(sheets)} sheet(s).")
right.download_button("⬇️ Download corrected Excel", data=xlsx, file_name=out_name, type="primary",
                      mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                      width="stretch")

with st.expander(f"What was changed ({len(log)} corrections)", expanded=False):
    for item in log:
        line = f"**{item['sheet']}**: {item['message']}"
        if item["level"] == "warn":
            st.warning(line, icon="⚠️")
        else:
            st.markdown("- " + line)

st.subheader("Preview")
tabs = st.tabs([s.name for s in sheets])
for i, (tab, s) in enumerate(zip(tabs, sheets)):
    with tab:
        show_sheet(s, i)

with st.expander("Corrections applied"):
    st.markdown(RULES)
