# Sorting & layout guide (layout)

## Sorting

`{op:"sort_range", sheetId, range:"A1:C10", byColumn:"B", order:"asc"|"desc", hasHeader?:true}`

- Stable sort by **current values**: numbers < text < booleans; blank cells always sort last regardless of direction; string comparison is number-aware (item2 < item10).
- byColumn must fall within range; when hasHeader is true the first row does not participate in the sort.
- Ranges of ≤2000 cells expand into per-cell changes in the preview (with concurrency protection), so the user can see each cell's before/after values; the range must be loaded first (read_range loads it).
- Larger ranges (up to 200,000 cells) apply as one range-level operation: the executor loads each region itself and writes the reordered rows in bulk — no per-cell preview.
- Only values move; each cell's formatting stays at its position (unlike Excel, which drags formats along with rows). On a column with mixed number formats the moved values take on the format of the row they land in.
- **A range containing formulas is rejected outright** — moving formula text silently changes what relative references point at. Explain the reason to the user (you may suggest converting the formula column to values first); do not try to work around it.
- Sort the whole table together: range must cover all related columns — sorting a single column tears rows apart.

## Merging

- `{op:"merge_cells", sheetId, range:"A1:C1"}` — rejected if it overlaps an existing merged region; needs at least two cells.
- `{op:"unmerge_cells", sheetId, range}` — clears all merges intersecting the range.
- Merging is mainly for title rows/category headers; avoid merging in data regions (it breaks sorting and formula filling).
- For title-like content prefer "merge + center + fill color" over scattered standalone cells.

## Row height & column width

- `{op:"set_row_height", sheetId, row:2, count?:1, heightPoints:24}` — height in points (2–409).
- `{op:"set_col_width", sheetId, column:"B", count?:1, widthPx:120}` — width in pixels (10–2000).
- Give report title rows extra height (e.g. 24–28pt); size data columns so content doesn't overflow, and keep similar columns the same width.

## Hiding rows/columns

- `{op:"set_rows_hidden", sheetId, row:5, count?:1, hidden:true|false}` — row is 1-based; hidden:false unhides.
- `{op:"set_cols_hidden", sheetId, column:"C", count?:1, hidden:true|false}`
- Hiding doesn't delete data and formula references are unaffected; good for hiding helper-calculation columns. To delete data use delete_rows/delete_cols (mind the batching discipline).

## Freeze panes

`{op:"set_freeze", sheetId, rows:1, columns:0}` — freezes the top rows rows and left columns columns; `rows:0, columns:0` unfreezes. Takes effect on the canvas immediately and is written into the file on save. Typical use: freeze the header row with `rows:1`.

## Page setup (printing)

`{op:"set_page_setup", sheetId, orientation?, paperSize?, scale?, fitToWidth?, fitToHeight?, margins?, printGridlines?, printHeadings?, printArea?}` — at least one property, only works on imported xlsx files, written on save:

- `orientation`: "portrait" | "landscape"
- `paperSize`: OOXML paper code (1=Letter, 8=A3, 9=A4, 11=A5)
- `scale`: print scaling 10–400(%); **mutually exclusive** with fitToWidth/fitToHeight (fit means "fit to N pages wide/tall", 0=automatic)
- `margins`: "normal" | "wide" | "narrow"
- `printGridlines` / `printHeadings`: print gridlines / row-column headings
- `printArea`: "A1:H40" sets the print area, null clears it
- Settings are only written into the file (take effect when printing); the canvas does not show pagination.
