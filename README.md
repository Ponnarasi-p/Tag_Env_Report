# Printing Daily Reports – Report Correction Tool

A small website that takes the aging-stock Excel report downloaded from the existing system,
fixes its data and formatting, and gives back a corrected `.xlsx`.

Everything runs in the browser, so files are never uploaded to a server. There is no backend to run.

## Using it

1. Open the site (see *Hosting* below).
2. Drop the downloaded report (`.xlsx` or `.xls`) onto the page.
3. Look over the list of changes and the preview, then click **Download corrected Excel**.

## What it corrects

| # | Rule | Where |
|---|------|-------|
| 1 | Removes rows that have an item name but no other values (empty or `NULL`). Rows with at least one value are kept. S.No is renumbered. | All sheets |
| 2 | Converts **Day Consumption** from tons to kg (×1000) and recalculates **No. of Days Consumption** as `Stock ÷ Day Consumption (Kg)`. With no Stock column, the old value is divided by 1000, which gives the same result. | Any sheet with these columns |
| 3 | Removes Lot Code, Lot No, Supplier Lot No, Category, GSM, Qty and Machine Line. | Raw Aging Stock |
| 4 | Removes Lot Code, Lot No, Supplier Lot No, Width, GSM and Source Type. Groups repeated item names into one summary row with total QTN and average Holding Days. The individual records sit underneath: click the item in the web preview, or the **[+]** in Excel's left margin, to expand. | All TagEnv Aging Stock |
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
[`js/processor.js`](js/processor.js).

## Hosting

It is a static site, so any static host works. With **GitHub Pages**: repo **Settings → Pages →
Deploy from branch → `main` / root**. The site will then be at
`https://<user>.github.io/printing_daily_reports/`.

To try it locally:

```bash
python -m http.server 8000
```

then open http://localhost:8000.

## Files

* `index.html`, `css/style.css`: the page
* `js/processor.js`: all correction rules (pure functions, can also run in Node)
* `js/excel-writer.js`: writes the formatted workbook (ExcelJS)
* `js/app.js`: upload, preview and download

Libraries (loaded from cdnjs): SheetJS 0.18.5 for reading `.xls`/`.xlsx`, and ExcelJS 4.4.0 for writing styled `.xlsx`.
