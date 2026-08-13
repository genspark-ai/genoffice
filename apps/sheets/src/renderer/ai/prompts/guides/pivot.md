# Pivot table guide (pivot)

## When to pivot vs. formula aggregation

- The user asks for a "pivot table", or wants the result interactive in Excel (drag fields, refresh) → **add_pivot** (real pivot table).
- They just want "summarize by dimension" and will keep computing/charting from the summary values → a formula aggregation table also works (see the fallback plan): the result is real cell formulas that stay computable.

## Creating a pivot table (add_pivot)

`{op:"add_pivot", sheetId, sourceRange:"A1:F1000", targetCell:"H1", targetSheetId?, rowFields, columnField?, values:[{field, agg, showDataAs?}], name?}`

- **The first row of sourceRange must be headers** (non-empty and mutually unique); rowFields/columnField/values.field all use the header text.
- rowFields: row dimensions, 1–8 (string or array, **outermost first**); with multiple levels, rows group hierarchically and per-level subtotal rows are added automatically. columnField: column dimension (optional, at most 1; **when columnField is given, values may only have 1 entry**).
- values: `{field:"Amount", agg:"sum"|"count"|"average"|"max"|"min", showDataAs?}`, 1–8 entries; a values field cannot also be a row/column dimension.
- showDataAs (value display mode, optional): `"percentOfTotal"` (% of grand total) | `"percentOfRow"` (% of row total) | `"percentOfCol"` (% of column total); displayed with the 0.00% format by default.
- targetCell: top-left of the output region, **must not overlap the source region** (put it in blank space right of the source, or use targetSheetId to place it on a summary sheet).
- How it takes effect: on apply, aggregated results (including subtotal and Grand Total rows/columns) are written directly to the target cells and are immediately visible; on save, native pivot parts are written (refreshOnLoad), so **the file opens in Excel as a live interactive pivot table**.
- Limits: source region ≤ 10,000 data rows × 200 columns; each row-dimension level ≤ 10,000 distinct members; the same batch cannot also do row/column insertion/deletion or sheet management on the source/target sheet (save first).
- The output region is protected after saving (not directly editable) — don't stack set_cell on top of it.

```json
{
  "op": "add_pivot",
  "sheetId": "s1",
  "sourceRange": "A1:D500",
  "targetCell": "F1",
  "rowFields": ["Region", "Product"],
  "values": [
    { "field": "Sales", "agg": "sum" },
    { "field": "Sales", "agg": "sum", "showDataAs": "percentOfTotal" }
  ],
  "name": "RegionProductSummary"
}
```

## Refreshing existing pivot tables (refresh_pivot)

`{op:"refresh_pivot", sheetId}` — recomputes the data areas of all pivot tables on the sheet and writes them back. Use it after the source data has changed to bring pivot results up to date.

- When **new categories** appear in the source data, the layout grows automatically: new members are appended at the end of their level (with new subtotal rows for multi-level layouts) and the output region expands — but the area the growth occupies must be empty, otherwise it errors and asks you to clear it first.
- Cases that still fail with explicit errors: renamed headers / moved data sources, calculated fields, grouped fields, value filters, and growth of compact (non-tabular) layouts. Other edge cases are in the data guide's "Pivot tables" section.

## Fallback plan: formula aggregation table (for irregular data or when live formulas are needed)

Compute the summary grid directly with `SUMIFS`/`COUNTIFS`/`AVERAGEIFS`:

**Single-dimension summary** (sales by region):

```json
[
  { "op": "set_range", "sheetId": "s2", "range": "A1:B1", "values": [["Region", "Sales"]] },
  {
    "op": "set_range",
    "sheetId": "s2",
    "range": "A2:A5",
    "values": [["East"], ["South"], ["North"], ["West"]]
  },
  {
    "op": "set_formula",
    "sheetId": "s2",
    "address": "B2",
    "formula": "=SUMIFS(Sheet1!$C:$C,Sheet1!$A:$A,A2)"
  }
]
```

(Fill the B2 formula down to B5; point the criteria ranges at the source table's region/amount columns.)

**Two-dimension cross-tab** (region × month): row headers hold regions, column headers hold months, and each intersection uses a double-criteria `=SUMIFS(src!amount, src!region, $rowHeader, src!month, colHeader$)` (note the `$` locking directions for filling).

**Deduplicating dimension values**: when the source's dimension values are unknown, first `read_range` the source column, write the deduplicated dimension values into the row/column headers in your proposal, then attach the SUMIFS.

## Companion steps

- Once the summary table is computed, add a chart per the `charts` guide (category comparison → column, share → pie).
- Apply `financial-formatting` currency formats to amount columns.
- For formula aggregation tables, the total row uses `=SUM(...)`, bold + light gray fill (add_pivot ships its own Grand Total).
