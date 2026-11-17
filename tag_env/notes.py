"""Warnings that carry the full list of affected items, not just a count."""
from __future__ import annotations

import pandas as pd


class Note(str):
    """A warning: a one-line title plus a table naming every affected item.

    It is a str (title followed by one line per item), so it prints in full
    anywhere a plain message is expected; `.title` and `.table` are used by the
    app to show it as a heading and a table.
    """

    title: str
    table: pd.DataFrame

    def __new__(cls, title: str, rows=None, columns: list[str] | None = None):
        if rows is None:
            table = pd.DataFrame()
        elif isinstance(rows, pd.DataFrame):
            table = rows.reset_index(drop=True)
        else:
            rows = list(rows)
            if rows and not isinstance(rows[0], (list, tuple, dict)):
                rows = [(r,) for r in rows]
            table = pd.DataFrame(rows, columns=columns)
        lines = [title]
        for rec in table.itertuples(index=False):
            parts = []
            for col, v in zip(table.columns, rec):
                if v is None or (isinstance(v, float) and pd.isna(v)) or v == "":
                    continue
                if isinstance(v, float):
                    v = f"{v:,.2f}"
                parts.append(f"{col}: {v}" if len(table.columns) > 1 else str(v))
            lines.append("  - " + " | ".join(parts))
        obj = super().__new__(cls, "\n".join(lines))
        obj.title = title
        obj.table = table
        return obj


def as_note(w) -> Note:
    return w if isinstance(w, Note) else Note(str(w))
