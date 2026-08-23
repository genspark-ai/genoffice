# Charts & shapes guide (charts)

## Creating a chart

`{op:"add_chart", sheetId, chartType:"column"|"bar"|"line"|"area"|"pie"|"doughnut"|"scatter"|"radar"|"combo", dataRange:"A1:C10", title?, anchorCell?}`

- Works for both imported xlsx files and the demo workbook. In xlsx files, saving (⌘S) writes the chart into the file as a brand-new chart part; demo-workbook charts live in memory (undoable with ⌘Z) and cannot be saved to a file yet.
- dataRange is the chart's data source (max 2000 cells): if the first row contains text it is treated as series-name headers, and if the first column contains text it is treated as the category axis. **read_range the data first to confirm its shape before proposing** — unloaded cells read back as empty.
- At least one numeric column is required, otherwise the operation is rejected. In the demo workbook, formula cells have no cached values (they read back empty) — do not point dataRange at a region of pure formula results.
- anchorCell is the anchor cell for the chart frame's top-left corner; by default the chart is placed two columns to the right of the data range.
- title default: the series name for a single series, "Chart Title" for multiple series.
- New column/bar charts come with numeric data labels by default (each bar labeled with its value); if unwanted, follow up with an `edit_chart` setting `dataLabels` to "none".
- Type selection: time series → line/area; category comparison → column/bar; share of total → pie (pie uses only the first numeric column — do not pick pie for multi-series data); two metrics of very different magnitude (e.g. amount + percentage) → combo.
- combo (column + line): the last series is drawn as a line on a secondary right-hand value axis, the remaining series are clustered columns on the left axis; requires ≥2 series — a single series degenerates into a plain column chart. A combo chart cannot be converted to another type once created.
- Charts/shapes/images can also be inserted directly on sheets created in this session.

## Editing an existing chart

`{op:"edit_chart", chartPath:"xl/charts/chart1.xml", title?, chartType?, seriesColors?, legend?, dataLabels?, grouping?, axisTitles?, seriesData?}` — at least one property:

- To see which charts the workbook has, check the "Charts in the workbook" list from get_workbook_context: each entry gives the chartPath, current title, type, and owning sheetId. **chartPath must come from that list** — never invent one.
- Charts that came with the file have chartPaths like `xl/charts/chart1.xml`; charts added this session and demo-workbook charts use the visual id from the list (like `added-chart-…` / `demo-chart-…`) as the chartPath and are equally editable.
- `title`: chart title (≤255 characters)
- `chartType`: "column" | "bar" | "line" | "area" | "pie" | "doughnut" — only charts with a **single plot area** can be converted between these; combo, scatter, and 3D charts do not support conversion (a type-conversion proposal is rejected at the preview stage) — in that case change only the title/colors or explain to the user
- `seriesColors`: {"0":"#4472C4", "1":"#ED7D31"} — keys are series indices (strings), values are 6-digit hex
- `legend`: "none" (hide) | "right" | "bottom" | "top" | "left"
- `dataLabels`: "none" | "value" (numeric) | "percent" (percentage, pie) | "category-percent" (name + percentage, pie)
- `grouping`: "clustered" | "stacked" | "percentStacked" — applies only to column/bar/line/area charts
- `axisTitles`: {category?, value?} — axis titles, null clears; pie/doughnut charts have no axes, not applicable
- `seriesData`: [{index, name?, valuesRange?:"B2:B13", categoriesRange?:"A2:A13", sheetId?}] — rename a series or **repoint its data ranges** (single row or single column; sheetId defaults to the chart's sheet). On apply, the current worksheet values are synced into the chart cache, and reference formulas are written into the file as well.
- Change series colors as a whole set (keep one palette) — recoloring a single series breaks visual consistency; follow a high-contrast, low-saturation palette.

## Inserting a shape/text box

`{op:"add_shape", sheetId, shapeType, anchorCell:"E2", fillColor?, text?}`

- shapeType: any OOXML prstGeom name from the insert gallery ("rect", "roundRect", "ellipse", "heart", "star5", "flowChartProcess", "wedgeRectCallout", …; see the shape-image guide) or "textbox"
- textbox defaults to a white background with the text "Text"; other shapes default to a light-blue fill. text can carry the annotation directly.
- Likewise only works for imported xlsx files; written into the file on save.

## Editing shapes / inserting images

- `{op:"edit_shape", visualId, text?, fillColor?, anchorCell?}` — edits a shape/text box **added this session** (change text, change fill, move; size is preserved). Get visualId from read_sheet_features; shapes that came with the file cannot be modified.
- `{op:"add_image", sheetId, path:"~/logo.png", anchorCell:"B2"}` — inserts an image (PNG/JPEG/GIF). path is either a local file the user specified (≤20MB, absolute path or ~/ prefixed — **must be explicitly given by the user**, never guess file locations) or an https URL returned by image_search / generate_image. Size is computed automatically from the image's aspect ratio; written into the file on save.

## Rules

- add_chart / edit_chart / add_shape are all layout-class: they can share a batch with content/format operations, but not with structural operations (row/column or sheet insertion/deletion).
- On apply, changes render on the canvas immediately and are recorded in the edit journal; ⌘S saves them into the file. Edits to existing charts are surgical rewrites — everything else stays byte-for-byte intact.
- Write data before charting: within a batch you may set_range then add_chart, but add_chart reads **current values** — if the data hasn't been written to the sheet yet, split into two batches (write the data first, apply, then create the chart).

## Chart type selection (by data intent)

| Data intent                        | Recommended type                         | Notes                                                                          |
| ---------------------------------- | ---------------------------------------- | ------------------------------------------------------------------------------ |
| Time trend / continuous change     | `line` / `area`                          | line for multi-period trends; area to emphasize cumulative amount/volume       |
| Category comparison                | `column` (vertical) / `bar` (horizontal) | use bar when category names are long or numerous (easier to read horizontally) |
| Share of total (≤ 6 items)         | `pie`                                    | uses only the first numeric column; switch to bar when there are many items    |
| Relationship between two variables | scatter                                  | first column as X (must be numeric), remaining columns as Y                    |

Selection principle: one chart answers one question; for multiple comparison dimensions, prefer column (detail) + pie (share) as complements.

## Chart conventions

- The title **must be meaningful and include units** (e.g. "Sales by department (USD thousands)") — never "Chart Title"/"Chart 1".
- When there is a clear numeric comparison, **recommend adding a chart** — don't hand over bare data.
- When category-axis labels are long, prefer bar (horizontal) to avoid slanted labels.
- Multiple series share one palette (high contrast, low saturation); change seriesColors as a whole set.

## Deleting a chart/shape/image (delete_visual)

`{op:"delete_visual", visualId}` — for visualId use the path from get_workbook_context's chart list (charts) or the visual id from read_sheet_features (shapes/images). Both file-native objects and objects added this session can be deleted; deleting a file chart also removes its chart part.

## Sparklines (add_sparkline)

`{op:"add_sparkline", sheetId, type:"line"|"column"|"stacked", dataRange:"B2:F9", targetCell?, color?}` — each row of data generates one in-cell sparkline, placed one column to the right of dataRange (or downward starting from targetCell). stacked means win/loss. Only works for imported xlsx files; saved as Excel-native x14 sparklineGroups.
