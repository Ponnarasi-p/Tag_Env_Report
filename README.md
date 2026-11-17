# Printing Daily Reports – Report Correction Tool

A small website that takes the aging-stock Excel report downloaded from the existing system,
fixes its data and formatting, and gives back a corrected `.xlsx`.

There are two versions of the same tool in this repo:

* **Streamlit app** (`app.py`), the main one to deploy. See *Deploying on Streamlit* below.
* **Static web page** (`index.html`) runs entirely in the browser, for hosts like GitHub Pages.

Both apply exactly the same rules.

## Deploying on Streamlit

1. Go to https://share.streamlit.io and sign in with GitHub.
2. Click **Create app → Deploy a public app from GitHub**.
3. Pick this repository, branch `main`, main file `app.py`, then click **Deploy**.

Streamlit installs everything listed in `requirements.txt` automatically.

To run it on your own PC:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Using it

1. Open the deployed Streamlit app (or the static page).
2. Upload the downloaded report (`.xlsx` or `.xls`).
3. Look over the list of changes and the preview, then click **Download corrected Excel**.

## What it corrects

| # | Rule | Where |
|---|------|-------|
| 1 | Removes rows that have an item name but no other values (empty or `NULL`). Rows with at least one value are kept. S.No is renumbered. | All sheets |
| 2 | Converts **Day Consumption** from tons to kg (×1000) and recalculates **No. of Days Consumption** as `Stock ÷ Day Consumption (Kg)`. With no Stock column, the old value is divided by 1000, which gives the same result. | Any sheet with these columns |
| 3 | Removes Lot Code, Lot No, Supplier Lot No, Category, GSM, Qty and Machine Line. Rows are sorted by Holding Days, highest first. The quantity column (QTN) is renamed `QTY (kg)`. | Raw Aging Stock |
| 4 | Removes Lot Code, Lot No, Supplier Lot No, Width, GSM and Source Type. Groups repeated item names into one summary row with total QTN and average Holding Days. Rows are sorted by Holding Days, highest first (a grouped item uses its average; its records are sorted too). The quantity column (QTN) is renamed `QTY (kg)`. The individual records sit underneath: click the item in the web preview, or the **[+]** in Excel's left margin, to expand. | All TagEnv Aging Stock |
| 5–7 | Zero values become `-` | All sheets, including the 4th |
| 6 | Decimals are rounded to the nearest whole number. | All sheets |
| 8 | Styled header row, borders, banded rows, alignment (numbers right, text left), auto column widths, frozen header, filters, landscape print setup. | All sheets |

Sheet order and names stay the same as the original. Any title lines above the header row are kept as a merged title.

### How columns and sheets are recognised

Names are matched loosely: case, spaces, punctuation and anything in brackets are ignored, so
`Width (mm)`, `WIDTH` and `width` all match. Sheets are found by name:
* *Raw Aging Stock* = a name containing "raw" and "aging"/"ageing"
* *All TagEnv Aging Stock* = a name containing "tagenv"

If the source system uses a different header name, add it to the matching list in `CONFIG` at the top of
[`report_processor.py`](report_processor.py) (Streamlit) and [`js/processor.js`](js/processor.js) (static page).

## Hosting the static page

The static page needs no server, so any static host works. With **GitHub Pages**: repo **Settings → Pages →
Deploy from branch → `main` / root**. The site will then be at
`https://<user>.github.io/printing_daily_reports/`.

To try it locally:

```bash
python -m http.server 8000
```

then open http://localhost:8000.

## Files

Streamlit app:
* `app.py`: upload, preview (click a grouped item to see its records) and download
* `report_processor.py`: all correction rules (edit `CONFIG` at the top to add header names)
* `excel_reader.py` / `excel_writer.py`: read the upload, write the formatted workbook (openpyxl)
* `requirements.txt`

Static page:

* `index.html`, `css/style.css`: the page
* `js/processor.js`: all correction rules (pure functions, can also run in Node)
* `js/excel-writer.js`: writes the formatted workbook (ExcelJS)
* `js/app.js`: upload, preview and download

Libraries (loaded from cdnjs): SheetJS 0.18.5 for reading `.xls`/`.xlsx`, and ExcelJS 4.4.0 for writing styled `.xlsx`.
