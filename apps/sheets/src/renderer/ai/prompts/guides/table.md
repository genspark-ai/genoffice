# Structured table guide (table)

## When to use a structured table

When the data is a regular "header + homogeneous record rows" shape, prefer a Table: automatic banded fill, filter dropdowns, and after saving a native table object in Excel (auto-expansion, structured references).

## Creating a table (add_table)

`{op:"add_table", sheetId, range:"A1:D10", name?, style?, bandedRows?}`

- **The first row of range must be headers**: write the headers and data with set_range/set_cell first, then add_table. Headers must be non-empty and mutually unique — empty headers are auto-filled as Column1/Column2, duplicate headers get numeric suffixes, and both are written back into the cells.
- range needs at least two rows (header + ≥1 data row) and must not overlap merged cells, existing tables, or a worksheet filter (set_filter) region — conflicts are rejected at save time.
- name: table name (starts with a letter/underscore, may contain Chinese; unique within the workbook, shares a namespace with defined names). **Always naming explicitly is recommended**; defaults to Table1, Table2… (if it collides with a table already in the file, saving errors).
- style: built-in styles `TableStyleLight1..21` / `TableStyleMedium1..28` / `TableStyleDark1..11`, default `TableStyleMedium2`.
- bandedRows: banded fill, default true.

```json
[
  {
    "op": "set_range",
    "sheetId": "s1",
    "range": "A1:C1",
    "values": [["Product", "Quantity", "Amount"]]
  },
  {
    "op": "add_table",
    "sheetId": "s1",
    "range": "A1:C20",
    "name": "SalesDetail",
    "style": "TableStyleMedium2"
  }
]
```

Notes:

- After creating the table you can extend its structure directly with `add_table_row` / `add_table_column` etc., no save needed first. But tables that came with the file (existed at open time) cannot be modified with these ops — ask the user to save and reopen.
- Don't merge_cells inside table data; don't stack set_filter on the header row (the table has its own filter).

## Structured references (not supported by the engine yet; fall back to A1 references)

In Excel, in-table formulas can use `[@Column]`/`TableName[Column]`, but the current engine's formula evaluation does not parse structured references — **always write plain A1 references** (e.g. `=SUM(C2:C20)`). Once the saved file is opened in Excel, the user can switch to structured references themselves.

## Inserting/deleting rows and columns

> **Precondition**: these four operations only work on **tables created this session via add_table**; tables that came with the file are unsupported (error: please save and reopen before modifying).

### add_table_row — insert data rows

```json
{ "op": "add_table_row", "sheetId": "s1", "tableName": "SalesDetail", "row": 3, "count": 2 }
```

- `row` (optional, 1-based, relative to the data area): insert position; omitted = append at the end.
- `count` (optional, default 1): number of rows to insert, 1–1000.
- After the operation, fill the new rows with set_range/set_cell.

### delete_table_row — delete data rows

```json
{ "op": "delete_table_row", "sheetId": "s1", "tableName": "SalesDetail", "row": 2, "count": 1 }
```

- `row` (required, 1-based): first row to delete (relative to the data area, excluding the header).
- `count` (optional, default 1): number of rows to delete.
- At least 1 data row must remain; deleting them all is rejected.

### add_table_column — insert columns

```json
{
  "op": "add_table_column",
  "sheetId": "s1",
  "tableName": "SalesDetail",
  "column": 3,
  "columnName": "Notes",
  "count": 1
}
```

- `column` (optional, 1-based): insert position; omitted = append at the far right.
- `columnName` (required): header name of the new column; must not duplicate an existing column in the table.
- `count` (optional, default 1): number of columns to insert. When count>1, subsequent column names get numeric suffixes (Notes, Notes2…).

### delete_table_column — delete columns

```json
{
  "op": "delete_table_column",
  "sheetId": "s1",
  "tableName": "SalesDetail",
  "column": 4,
  "count": 1
}
```

- `column` (required, 1-based): first column to delete (relative to the table's left edge).
- `count` (optional, default 1): number of columns to delete.
- At least 1 column must remain; deleting them all is rejected.

### Typical workflow

1. Create and name the table with `add_table` (**explicit name recommended**).
2. Extend the structure with `add_table_row` / `add_table_column`, then fill data with `set_range` / `set_cell`.
3. When done, remind the user to save (⌘S); after saving, the table ref in the file updates automatically.

## Fallback plan (when a real table doesn't fit)

When the data is irregular (subtotal rows, merged multi-column headers), don't create a Table — simulate the look with a plain range:

```json
[
  {
    "op": "format_range",
    "sheetId": "s1",
    "range": "A1:D1",
    "format": { "bold": true, "fillColor": "#4472C4", "fontColor": "#FFFFFF" }
  },
  {
    "op": "add_conditional_format",
    "sheetId": "s1",
    "range": "A2:D100",
    "rule": {
      "kind": "formula",
      "formula": "=MOD(ROW(),2)=0",
      "format": { "fillColor": "#F2F7FF" }
    }
  },
  { "op": "set_filter", "sheetId": "s1", "range": "A1:D100" }
]
```

- Banded fill via a `MOD(ROW(),2)` conditional format (adapts automatically to row insertion/deletion).
- A filter on the first row (`set_filter`) approximates the Table's filter dropdowns.

## Deleting a table (delete_table)

`{op:"delete_table", sheetId, tableName}` — only tables created this session via add_table can be deleted; the semantics are "convert to a plain range": values and formats are kept, only the Table object is removed (style banding/filter disappear with it). Tables that came with the file cannot be deleted.
