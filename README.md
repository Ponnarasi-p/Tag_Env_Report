# Tag Env Stock Report Generator

Streamlit app that builds the **Tag Env Stock** Excel report from three uploads:

1. Stock Summary Report (`StockSummaryReport-COIMBATORE-EOU-*.xls`)
2. Order Status Report (`OrderStatusReport-*.xlsx`)
3. Exploded BOM (`Exploded_bom-*.xlsx`)

The output is written into a copy of the reference report
(`templates/Tag_Env_Stock_template.xlsx`, from *Tag Env Stock as on 28-10-26*), so sheet
names, hidden sheets, headers, merged cells, fonts, fills, borders, number formats, column
widths and row heights are the template's own. Only data rows, the date, the conversion
factors and the totals change.

## Run

```bash
pip install -r requirements.txt
streamlit run app.py
```

Command line (no UI):

```bash
python tools/run_cli.py STOCK.xls ORDERS.xlsx BOM.xlsx OUT.xlsx [TEMPLATE.xlsx]
```

## Business logic

**Placing stock.** Each Stock Summary line goes to a column by its *Machine Line Name*
(`config/settings.json` → `machine_line_to_column`) and to a row by item:
`config/item_row_map.csv` first, then by name (prefix `ENV`/`TAG` and pack suffixes such as
`DC SF C250 TBGS` removed). Items with no row get a new row, inserted alphabetically like the
manual report. `(GHANA)`, `(ABIDJAN)`, `(ANGOLA)` items go to the *Ghana & AB & AG Stock*
sheet. Lines on non tag/envelope machine lines (paper, ink, divider) are ignored.

**Make Cases** (written as Excel formulas, as in the manual report):

```
Make Cases = (Available Kg × Conversion factor) / Pcs per CFC
```

- *Conversion factor*: the column's machine line from the conversion reference table, kept
  in row 4 of the main sheet and in `config/settings.json`.
- *Pcs per CFC*: from the Exploded BOM. The total ENV (or TAG) pieces in one CFC of the
  finished good the item is used in. A variety pack with three flavours of 80 envelopes gives
  240, which is what the manual report uses.
- When an item is in several finished goods:
  1. a finished good with a **pending order** wins (largest pending quantity);
  2. else the value used in the template report, if the BOM still has it;
  3. else the most common pack size in the BOM.
- Item not in the BOM: the template report's value is carried over and listed as a warning.
  Nothing at all: Make Cases is left empty and listed as a warning.
- BOM lines under 1 piece per CFC (data errors, e.g. `0.0072`) are ignored and reported.

**Shared tags** (e.g. one GV FLAVOUR tag for all GV envelopes) are written once and merged
down over the rows that share them, using the template's merged blocks.

**Remarks** are manual notes and are carried over from the template by row.

## Validation done

With the 06-11 Stock Summary and Order Status files and the 28-10 report as template, the
generated report was compared with the manually prepared *Tag_Env_Stock_as_on_06-11-26 (3)*:

| Check | Result |
|---|---|
| Rows and row order (207 rows) | identical |
| Merged cells | identical |
| Stock Kg per row/column (270 cells) | 270 / 270 match |
| Pcs per CFC denominators (270 cells) | 264 / 270 match |
| Excel recalculation (3,613 formulas) | 0 errors |
| Structure/format checks vs template | 25 / 25 pass |

The 6 denominator differences are items with no pending order where the manual report's
choice changed between reports (e.g. KHAIMA CINNAMON 1125 → 600), or where a pending order
exists but the manual report used another pack size.

Compare any generated report with a manual one:

```bash
python tools/compare_manual.py GENERATED.xlsx MANUAL.xlsx
```

## Differences from the template, on purpose

- MD20 HS, Maisa and Perfecta/PK45 Make Cases formulas in the template point at another
  column's conversion cell (e.g. Maisa Tag used `$I$4` = 12600). Each column now uses its
  own conversion cell.
- MD20 HS Envelope uses *MD HS Envelope Paper* = 1425 from the conversion table (template
  row 4 had 1200). Change it in `config/settings.json` or in the app's Settings panel.
- Links to an external workbook (`[1]MD-20 HS`, `[1]Stock Summary...`) are replaced by values.
- Hidden-sheet totals use the full data range (the template had partial ranges).
- Freeze panes was at B172 in the template; it is set to B5.

## Maintaining

- **Conversion factors / column mapping**: `config/settings.json`.
- **Item → row links**: `config/item_row_map.csv`. To learn links from a newer manual report:
  `python tools/learn_item_row_map.py MANUAL_REPORT.xlsx STOCK_SUMMARY.xls`
- **MJP ↔ M4 item names** (for stock names that differ from BOM names):
  `config/mjp_m4_item_mapping.csv`.
- **Template**: replace `templates/Tag_Env_Stock_template.xlsx`, or upload a different
  reference report in the app's Settings panel.

## Tests

```bash
TAG_ENV_STOCK=... TAG_ENV_ORDERS=... TAG_ENV_BOM=... python tests/test_app.py
```
