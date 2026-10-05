# Sheets: spreadsheets

Sheets is the Excel-like editor; calculation runs in a separate Rust engine process (a crash there never takes the app down). Opens and saves genuine .xlsx; .csv and .tsv open as tables.

## The interface

- **Ribbon**: eight tabs, listed one by one below.
- **Formula bar**: shows and edits the active cell's formula; common functions supported.
- **Sheet tabs** (bottom): add / rename / delete / move sheets.
- **Cell editing**: double-click or just type; Enter confirms and moves down, Tab moves right, Escape cancels (Excel habits).
- **Shortcuts**: aligned with the Excel family (ctrl+C/V/X, ctrl+Z/Y, ctrl+F, ...).

## Ribbon tabs

- **Home**: font, fill, borders, number formats (currency/percent/thousands, decimal up/down), alignment, merge, row/column insert and size, conditional formatting, format-as-table, cell styles, clipboard & format painter, sort & filter.
- **Insert**: shapes, icons, symbols, equation, screenshot and more.
- **Page Layout**: theme colors & fonts, print gridlines/headings toggles, page-break preview.
- **Formulas**: AutoSum and function insert, define names (also from selection), trace precedents/dependents, the Watch Window, recalculate sheet/workbook.
- **Data**: sort & filter (incl. advanced filter, clear filter), text-to-columns, merge workbooks, refresh all.
- **Review**: browse comments (show, previous/next), translate.
- **View**: gridlines & headings toggles, zoom, Normal / page-break preview.
- **Chart Design**: appears with a chart selected — chart type, styles and colors, edit the data range.

The Data tab, button by button (left to right in the picture):

![The Data tab](img/sheets-data.png)

- **Pivot table**: build a pivot from the current range; drag fields to aggregate.
- **Refresh**: recalculate the current pivot's data.
- **From text/CSV**: import a .csv/.txt as a new sheet, split by delimiter.
- **Merge workbooks**: pull sheets from other .xlsx files into this one.
- **Refresh all**: recalculate every pivot and external dataset.
- **Sort** (dropdown): ascending / descending / custom sort (multi-column rules).
- **Filter**: adds ▼ dropdowns to the header row; check the values to keep.
- The small stacked buttons next to it: **Clear** (bring every row back), **Reapply** (run the current filter again), **Advanced** (filter with a criteria range).
- **Text to columns** (dropdown): split one column into several by delimiter or fixed width.
- **Quick fill**: give one example and the rest of the column fills by example (ctrl+E).
- **Remove duplicates**: drop duplicated rows by the selected columns.
- **Data validation** (dropdown): input rules for the selection (dropdown lists, number ranges...).
- **Consolidate**: aggregate several ranges into one place by category.
- **What-if** (dropdown): goal seek / data tables.
- **Group / Ungroup** (dropdown): row/column groups with collapse and expand.
- **Subtotal**: insert subtotal rows per category.

The Formulas tab, button by button:

![The Formulas tab](img/sheets-formulas.png)

- **Insert function** (fx): search functions with an argument wizard.
- **AutoSum** (dropdown): one-click SUM, plus average/count/max/min.
- **Recently used / Financial / Logical / Text / Date & time / Lookup & reference / Math & trig / More**: browse and insert functions by category.
- **Name manager**: view, create and delete named ranges.
- **Define name** (dropdown): name the selection; **Use in formula** inserts an existing name; **Create from selection** names ranges from their header row/column.
- **Trace precedents / Trace dependents**: blue arrows showing where a formula's data comes from and where it feeds; **Remove arrows** clears them.
- **Show formulas**: cells show the formula itself instead of the result.
- **Error checking**: locate and explain formula errors.
- **Watch window**: pin cells you care about and watch their live values.
- **Calculation options** (dropdown): automatic vs manual recalculation; in manual mode **Calculate now / Calculate sheet** trigger it by hand.

## Numbers and formatting

- Number formats: general, number, currency, percent, date/time, fraction, scientific and more.
- Alignment, wrap, merged cells, borders and fills.
- Row heights and column widths by dragging; double-click a boundary to auto-fit.

## Data

**Sort & filter** (descending by one column, for example):

1. Click **any cell in that column** (no need to select the whole column).
2. Home tab ▸ **Sort & Filter** ▸ **Descending**; whole rows reorder together (the area sorts as one).
3. For custom rules (multiple columns, by color): the same path, **Custom Sort**.
4. Filter: select the header row and click **Sort & Filter ▸ Filter** — each header gets a ▼ dropdown where you check the values to keep; clear the filter to bring everything back.

- Sort and filter.
- Frozen panes.
- .csv / .tsv: open directly as a table (tab-delimited tsv parsed as one); saving writes the original format back.

## Right-click menus

- **In the grid**: the editor's (Univer's) own menu — cut/copy/paste, insert and delete rows/columns, hide, merge cells, freeze panes and other everyday items.
- **On the bottom status strip**: choose which statistics the status bar shows (average / count / sum, ...); the choice sticks.
- **On a sheet tab at the bottom**: add / rename / delete / color / hide sheets (Univer's tab menu).
- The top tab strip's context menu is covered in [Tabs and window management](help://tabs).

## The script editor (advanced)

Sheets ships a **script editor** with a Google-Apps-Script-shaped API for bulk work.

- Open: Tools ▸ Script editor (or the developer menu, per version).
- UI: script library on the left (new/delete), code editor + output pane on the right; **Run / Stop** buttons.
- The API is async:

```js
const sheet = await SpreadsheetApp.getActiveSpreadsheet()
const active = await sheet.getActiveSheet()
const range = await active.getRange('A1:C10')
const values = await range.getValues() // 2D array
await range.setValues(values.map((row) => row.map((v) => v * 2)))
Logger.log('done')
```

- `SpreadsheetApp` (entry), `Sheet` (getName/getRange/getLastRow...), `Range` (getValue(s)/setValue(s)/clear...), `Logger.log`, `Utilities.sleep`.
- Scripts run in a **sandboxed worker**: no network, no file system, no DOM — the API above is all they can touch, so a buggy or hostile script cannot reach anything else.
- The output pane caps its line count so huge logs cannot wedge the UI.
- **The AI can run scripts too**: the assistant's `run_script` tool executes the same sandboxed API — ideal for bulk, rule-based transforms.

## AI

- The side AI panel: select a range and instruct in plain language (reformat, generate data, write formulas).
- AI edits can be rolled back from the panel.

## Saving and export

- Saves .xlsx (formulas and formats preserved); Save As; Export to PDF follows print pagination.
- Autosave follows the global rule (on after the first manual save).

## Stability

- The Rust calculation sidecar is process-isolated from the UI: if extreme data kills it, you get a message and a session recovery attempt — not an app crash.
