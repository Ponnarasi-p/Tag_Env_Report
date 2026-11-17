"""Reads an uploaded .xlsx / .xls / .csv into (sheet name, rows) pairs."""
from __future__ import annotations

import io
import math
import re

import pandas as pd


def _clean(v):
    if v is None or v is pd.NaT:
        return None
    if isinstance(v, float) and math.isnan(v):
        return None
    if isinstance(v, pd.Timestamp):
        return None if pd.isna(v) else v.to_pydatetime()
    return v


def read_sheets(name: str, data: bytes) -> list[tuple[str, list[list]]]:
    lower = name.lower()
    if lower.endswith(".csv"):
        frames = {re.sub(r"\.csv$", "", name, flags=re.I): pd.read_csv(io.BytesIO(data), header=None, dtype=object)}
    else:
        engine = "xlrd" if lower.endswith(".xls") else "openpyxl"
        frames = pd.read_excel(io.BytesIO(data), sheet_name=None, header=None, dtype=object, engine=engine)
    return [(sheet, [[_clean(v) for v in row] for row in df.itertuples(index=False, name=None)])
            for sheet, df in frames.items()]
