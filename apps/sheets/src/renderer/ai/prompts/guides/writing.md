# Content writing guide (writing)

## Operation definitions

- `{op:"set_cell", sheetId, address:"B2", value}` — writes a single value (string/number/boolean/null). Optional `expectedValue` for concurrency protection.
- `{op:"set_formula", sheetId, address, formula:"=SUM(A1:A10)"}` — writes a formula, must start with =. Optional `expectedFormula`.
- `{op:"clear_cell", sheetId, address}` — clears a single cell.
- `{op:"set_range", sheetId, start:"B2", values:[[row1], [row2], …]}` — bulk write: a 2D array laid out by rows, spreading right and down from start. **values must be rectangular — every row the same length** (use null for a cell to clear; a shorter row would silently leave stale content behind). Instead of start you may pass `range:"B2:D5"`, whose size must exactly match values. **Strings starting with = are written as formulas.** Always use set_range for contiguous regions; never split them into many set_cell ops.
- **When rewriting rows of an existing table, cover the full row span**: values rows should span all of the table's columns (or the write should be paired with clear_range for the leftover cells) so no stale fragments from the old content survive next to the new one.
- `{op:"clear_range", sheetId, range:"A1:C10"}` — clears a rectangular region.

Limit: at most 2000 expanded cell changes per batch.

## Formula rules

- **Always leave calculations to formulas** — never compute values yourself and hard-code them; the table should update automatically when source data changes.
- If a reference can be expressed with relative/absolute references ($), don't write a constant.
- Stay consistent with existing conventions: if earlier rows use formulas, new rows use formulas too; if earlier tax amounts are negative, newly filled ones should be negative too.

**Wrong** — computed mentally and hard-coded:

```json
{ "op": "set_cell", "sheetId": "s1", "address": "B10", "value": 5000 }
```

**Right** — let the sheet compute it:

```json
{ "op": "set_formula", "sheetId": "s1", "address": "B10", "formula": "=SUM(B2:B9)" }
```

## Share/percentage columns (where the total must be 100%)

- Keep full precision in the cells: use a formula like `=B4/SUM($B$4:$B$8)` — do **not** write rounded constants row by row, which makes the total display 100.1% or 99.9%.
- Rounding belongs to display formatting (format_range's numberFormat, at least `0.000%`), not to the stored value.

## Values and units

- Ratios/growth rates/shares and other relative values are stored as **decimals in (0,1)** (0.1234 means 12.34%); if the source data is a percent number like 26.64, divide by 100 before writing.
- Keep units, decimal places, and currency symbols consistent within a column; note the unit in the header or title (e.g. "Amount (USD millions)").
- Dates: prefer writing ISO form (2026-07-23) + a number format controlling the display.

## Find & replace (find_replace)

`{op:"find_replace", sheetId, range:"A1:D50", find:"old text", replace:"new text", matchCase?, wholeCell?}`

- Replaces **text values only**: formula cells and numbers are untouched; expands into per-cell set_cell, so preview/undo match manual input.
- Case-insensitive by default (matchCase:true enables case sensitivity); with wholeCell:true a cell is replaced only when its whole content matches exactly.
- range is required and ≤2000 cells; in large files it can only act on loaded regions — unloaded regions are rejected.
