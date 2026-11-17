"""End-to-end test of the Streamlit app: upload the three files, generate, check output.

Set TAG_ENV_STOCK, TAG_ENV_ORDERS, TAG_ENV_BOM to sample files, then:
    python tests/test_app.py
"""
import io
import os
import sys
from pathlib import Path

import openpyxl
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
X = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def main():
    files = [os.environ[k] for k in ("TAG_ENV_STOCK", "TAG_ENV_ORDERS", "TAG_ENV_BOM")]
    at = AppTest.from_file(str(ROOT / "app.py"), default_timeout=240)
    at.run()
    assert at.button[0].disabled, "Generate must be disabled before uploads"
    for up, f in zip(at.file_uploader[:3], files):
        up.set_value((Path(f).name, Path(f).read_bytes(), X))
    at.run()
    assert not at.exception, at.exception
    print("after upload:", [s.value for s in at.success])
    assert len(at.success) == 3, [e.value for e in at.error]
    assert not at.button[0].disabled
    at.button[0].click()
    at.run()
    assert not at.exception, at.exception
    rep = at.session_state["report"]
    print("file:", rep.file_name, "| rows:", rep.summary["rows"], "| checks failed:",
          int((rep.checks["Result"] == "FAIL").sum()))
    print("success boxes:", [s.value for s in at.success])
    print("warnings shown:", len(at.warning))
    wb = openpyxl.load_workbook(io.BytesIO(rep.xlsx))
    print("sheets:", wb.sheetnames, [ws.sheet_state for ws in wb.worksheets])
    assert (rep.checks["Result"] == "PASS").all()

    # wrong file in a slot -> clear error, no crash
    at.file_uploader[0].set_value(("bom.xlsx", Path(files[2]).read_bytes()[:0] + Path(files[1]).read_bytes(), X))
    at.run()
    print("wrong-file error:", [e.value for e in at.error][:1])
    assert at.error and at.button[0].disabled
    print("APP TEST PASSED")


if __name__ == "__main__":
    main()
