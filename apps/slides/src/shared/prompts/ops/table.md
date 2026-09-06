# Table ops

> Edit an existing table element: cell text, merges, row/column structure, row heights, column widths, cell anchors.

All ops take `target:{slide, el}` where `el` is the table's id (type `table` in
the outline). `row` and `col` are 0-based. Table styling (`setTableStyle`) and
chart edits (`setChart`) live in dedicated tools.

### setTableCell

`{row,col,paragraphs}`

Replaces one cell's text. Runs you do not restyle inherit the cell's current
formatting, so plain `{text}` runs keep size, color and bold.

| Field      | Type                                                                                            | Notes                   |
| ---------- | ----------------------------------------------------------------------------------------------- | ----------------------- |
| row, col   | integer                                                                                         | 0-based                 |
| paragraphs | array of `{runs:[{text, bold?, italic?, fontSize?, color?}], align?, bullet?, lineSpacingPct?}` | Same shape as `setText` |

```json
{
  "op": "setTableCell",
  "target": { "slide": 0, "el": "e_TABLE" },
  "row": 0,
  "col": 0,
  "paragraphs": [{ "runs": [{ "text": "Metric", "bold": true }], "align": "center" }]
}
```

Common mistakes

- Passing the cell text as a string: it must be a paragraph array.
- Row or column outside the grid: the op reports the failing (row, col).

### tableMerge

`{kind:"merge-right"|"merge-down"|"split",row,col}`

Merges the cell at (row, col) with its right or lower neighbor, or splits a
merged cell back into its grid cells.

```json
{
  "op": "tableMerge",
  "target": { "slide": 0, "el": "e_TABLE" },
  "kind": "merge-right",
  "row": 0,
  "col": 0
}
```

Common mistakes

- Merging across an existing merge boundary: the op refuses; split first.

### tableStructure

`{kind:"insert-row"|"delete-row"|"insert-col"|"delete-col",index,before?}`

Inserts or deletes a row or column. Insert goes after `index` unless
`before:true`.

```json
{
  "op": "tableStructure",
  "target": { "slide": 0, "el": "e_TABLE" },
  "kind": "insert-row",
  "index": 1
}
```

```json
{
  "op": "tableStructure",
  "target": { "slide": 0, "el": "e_TABLE" },
  "kind": "delete-col",
  "index": 0
}
```

Common mistakes

- Tables with merged cells refuse row/column surgery: split the merges first (`tableMerge` with `kind:"split"`).

### setTableRowHeight

`{row,hEmu}`

Sets one row's height in EMU (the table frame grows or shrinks accordingly).

```json
{ "op": "setTableRowHeight", "target": { "slide": 0, "el": "e_TABLE" }, "row": 0, "hEmu": 457200 }
```

### setTableCellAnchor

`{row,col,anchor:"top"|"middle"|"bottom"}`

Vertical alignment of the text inside one cell.

```json
{
  "op": "setTableCellAnchor",
  "target": { "slide": 0, "el": "e_TABLE" },
  "row": 0,
  "col": 1,
  "anchor": "middle"
}
```

### setTableColWidth

`{col,wEmu}`

Sets one column's width in EMU.

```json
{ "op": "setTableColWidth", "target": { "slide": 0, "el": "e_TABLE" }, "col": 0, "wEmu": 2743200 }
```

### setTableStyle (not-ai-callable)

`{edit:TableStyleEdit} — use the edit_table_style tool instead`

Header row, banding, shading and border presets. The `edit_table_style` tool
exposes this with a validated schema.

### setChart (not-ai-callable)

`{patch:ChartEdit} — use the edit_chart tool instead`

Changes a chart's type, data, colors or elements. The `edit_chart` tool exposes
this with a validated schema.
