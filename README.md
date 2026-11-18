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
- *Pcs per CFC*: the pack size in the name of the finished good (FG) the envelope/tag is used
  in: `GV MINT 48/25 DC ENV TBGS` → 48 × 25 = **1200**. (If an FG name has no `n/m`, the BOM's
  pieces per CFC is used.)
- How the FG is found for an item, in order:
  1. the item is in the Exploded BOM, directly or through the MJP vs M4 item-name mapping;
  2. a BOM item with the same name once pack/machine words are removed
     (`TAG PACIFICO DC SF TBGS` = BOM `TAG PACIFICO DC TBGS`; MD-20 and C250 versions of one flavour);
  3. FG names in the BOM or Order Status whose brand/flavour before the pack size equals the item's
     (`ENV FESTIVA GINGER DC ENV TBGS` → `FESTIVA GINGER 24/25 DC ENV TBGS`);
  4. the longest FG brand the item name starts with
     (`ALDI VRUCHTENTHEE RASPBERRY` → `ALDI VRUCHTENTHEE 24/20 ...`);
  5. FG names that start with the item's brand (`ENV ELANI` → `ELANI GREEN TEA 24/25 ...`);
  6. otherwise the pack size of the same row's other envelope/tag.
- When an item is in several FGs: an FG with a **pending order** wins (largest pending quantity);
  else the most common pack size (on a tie, the smaller pack).
- **No FG found at all: Make Cases shows `N/A`**, and the app lists the item and asks for a new
  Exploded BOM. No value is taken from an older report.
- The calculation detail shows, for every item, the FG name used and how it was found.

Every number comes from the three uploaded files. The template supplies only the layout
(sheets, headers, formatting, row order, shared-tag blocks).

**Shared tags** (e.g. one GV FLAVOUR tag for all GV envelopes) are written once and merged
down over the rows that share them, using the template's merged blocks.

**Remarks**: "Order Received" when the Order Status report has a pending order for a finished
good that uses the row's envelope (from the BOM) or has the same brand name as the row (name
before the pack size, e.g. "AFRICA CHOICE 24/25 ..."); otherwise "No Order".

## Validation done

With the 06-11 Stock Summary and Order Status files and the 28-10 report as template, the
generated report was compared with the manually prepared *Tag_Env_Stock_as_on_06-11-26 (3)*:

| Check | Result |
|---|---|
| Rows and row order (207 rows) | identical |
| Merged cells | identical |
| Stock Kg per row/column (270 cells) | 270 / 270 match |
| Pcs per CFC denominators (270 cells) | 259 same, 11 different, 0 `N/A` |
| Excel recalculation | 0 errors, 0 `N/A` |
| Structure/format checks vs template | 25 / 25 pass |

The 11 differences are cells where the manual report used a different pack size from the FG
names available (e.g. TAG PACIFICO: FG `PACIFICO 12/20` = 240, manual 1200; KHAIMA DC SF: FG
`KHAIMA 12/106` = 1272, manual 1125/1200), or items with several pack sizes and no pending order.

`tests/test_unseen_inputs.py` runs the generator on inputs it has never seen (new brand,
moved header, no orders, removed items, duplicates, wrong files).

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
