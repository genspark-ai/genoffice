# Financial statement formatting guide (financial-formatting)

Applies to balance sheets, income statements, cash flow statements, P&L statements, financial analysis, and other accounting/finance tables. All operations use the `format_range` / `set_cell` / `merge_cells` / `set_hyperlink` DSL; format and content classes can share a batch, structural changes (row/column insertion/deletion) need their own batch.

## Number formats (accounting grade)

- **Lock the currency locale** (prevents `$`→`¥` remapping in Chinese-locale Excel): USD `[$$-409]#,##0`, CNY `[$¥-804]#,##0.00`.
- **Negatives in red parentheses**: `[$$-409]#,##0;[Red]([$$-409]#,##0)`.
- **Zero as a dash** (0 displays as `—`; blank means missing data): `[$$-409]#,##0;[Red]([$$-409]#,##0);"-"`.
- Numbers **right-aligned**, consistent decimal places within a column; note thousand/million units in the header ("Amount (thousands)").

```json
{"op":"format_range","sheetId":"s1","range":"B5:F20",
 "format":{"numberFormat":"[$$-409]#,##0;[Red]([$$-409]#,##0);\"-\"","horizontalAlign":"right"}}
```

## Hierarchy visualization

- Parent line items **bold**; child items express the parent-child relationship via format_range's `indent` field (+1 per level, effective both on screen and in the file).
- Do not simulate indentation with leading spaces — spaces pollute the cell text and break lookups and formula references.

```json
{"op":"set_cell","sheetId":"s1","address":"A6","value":"Accounts receivable"}
{"op":"format_range","sheetId":"s1","range":"A6","format":{"indent":1}}
```

## Total / subtotal rows

Bold + light gray fill `#F2F2F2`, separated with a thin top border + double bottom border:

```json
[
 {"op":"format_range","sheetId":"s1","range":"A20:F20","format":{"bold":true,"fillColor":"#F2F2F2"}},
 {"op":"format_range","sheetId":"s1","range":"A20:F20","format":{"border":{"type":"top","color":"#000000"}}},
 {"op":"format_range","sheetId":"s1","range":"A20:F20","format":{"border":{"type":"bottom","color":"#000000"}}}
]
```

## Headers

- Section titles (e.g. "Balance Sheet"): **merged across columns + centered + bold + visually larger (compensate with bold — this app does not change font size for this)**.
- Period columns ("FY2024 | FY2023 | FY2022"): bold + centered, separated with a medium bottom border.

```json
[
 {"op":"merge_cells","sheetId":"s1","range":"A1:F1"},
 {"op":"set_cell","sheetId":"s1","address":"A1","value":"Balance Sheet"},
 {"op":"format_range","sheetId":"s1","range":"A1:F1","format":{"bold":true,"horizontalAlign":"center"}},
 {"op":"format_range","sheetId":"s1","range":"B2:F2","format":{"bold":true,"horizontalAlign":"center"}}
]
```

## Color palette (conservative, professional)

- Header: navy `#1F3864` fill + white text `#FFFFFF` + bold
- Subtotal rows: light gray `#F2F2F2`
- Data rows: white
- Banding (optional): very light blue `#F8FAFC`
- Thin grid separators: `#D9D9D9` (border inner lines)

```json
{"op":"format_range","sheetId":"s1","range":"A2:F2","format":{"fillColor":"#1F3864","fontColor":"#FFFFFF","bold":true}}
```

## Financial model input coloring (when building models)

- Hard-coded inputs → blue text `#0000FF`
- Formula calculations → black text (default)
- Cross-sheet references → green text `#008000`

## Layout

- Widen the first column (line item names) with `set_col_width` to ~200px; value columns uniformly ~100px.
- Use border inner lines (`#D9D9D9`) for a light grid; avoid heavy outlines.

## Data source attribution

Financial data must cite sources (see the `data-attribution` guide): italic small footer text + hyperlink, e.g. "Source: SEC 10-K FY2024".

```json
[
 {"op":"set_cell","sheetId":"s1","address":"A22","value":"Source: SEC 10-K Filing, FY2024"},
 {"op":"format_range","sheetId":"s1","range":"A22","format":{"italic":true,"fontColor":"#666666"}},
 {"op":"set_hyperlink","sheetId":"s1","address":"A22","target":"https://sec.gov/Archives/edgar/..."}
]
```

## GAAP / IFRS conventions

- Identity: Assets = Liabilities + Equity
- Group line items logically, current before non-current
- Income statement order: Revenue → COGS → Gross profit → Operating expenses → Operating income → Net income
- Cash flow statement: Operating → Investing → Financing → Net change in cash

## Discipline

- **Styling failures never block**: if a format operation is rejected, prioritize delivering correct data; don't retry styling repeatedly.
- Every factual data cell needs a source; derive values with formulas, never compute them mentally and hard-code them.
