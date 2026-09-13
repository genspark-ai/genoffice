# Insert ops

> Create new elements on a slide: text boxes, preset shapes and lines (addElement), tables, native charts, SmartArt-style diagrams.

Insert ops take `target:{slide}` (no `el`) and an `offset:{x,y,cx,cy}` frame
in document-space EMU. They report the new element id in `created`; later ops
in the same transaction cannot reference it yet, so insert first and style in
the next call, or pass the style inline where the op supports it.

### addElement

`{kind:"textbox"|<preset geometry>,offset:{x,y,cx,cy},paragraphs?,fill?,stroke?,adjustments?:{<gd name>:val},bodyPr?:{autoFit?:"shrink"|"resize"}}`

Adds a text box, a preset-geometry shape, or a line/connector.

| Field       | Type                                                                                           | Notes                                                                                                                                                                                                                                              |
| ----------- | ---------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| kind        | string                                                                                         | `"textbox"`; a preset geometry (`rect`, `roundRect`, `ellipse`, `triangle`, `diamond`, `rightArrow`, `chevron`, `star5`, `hexagon`, `wedgeRoundRectCallout`, ...); or a line kind `line`, `lineArrow`, `lineArrowDouble`, `lineBent`, `lineCurved` |
| offset      | `{x, y, cx, cy}` EMU                                                                           | Required                                                                                                                                                                                                                                           |
| paragraphs  | array of `{runs:[{text, bold?, italic?, underline?, fontSize?, fontFamily?, color?}], align?}` | Optional text; `fontSize` in pt                                                                                                                                                                                                                    |
| fill        | `"#RRGGBB"` or `"#RRGGBBAA"`                                                                   | Solid fill; text boxes have none by default                                                                                                                                                                                                        |
| stroke      | `{color, widthEmu}`                                                                            | Outline; lines default to a 1 pt black stroke                                                                                                                                                                                                      |
| adjustments | `{<gdName>: number}`                                                                           | Preset adjust values, e.g. `{"adj": 25000}` for a roundRect radius                                                                                                                                                                                 |
| bodyPr      | `{autoFit?, wrap?, anchor?, insetsEmu?}`                                                       | `autoFit` `"shrink"` or `"resize"`; `wrap` `"square"`/`"none"`; `anchor` `"t"`/`"ctr"`/`"b"`                                                                                                                                                       |

```json
{
  "op": "addElement",
  "target": { "slide": 1 },
  "kind": "textbox",
  "offset": { "x": 914400, "y": 685800, "cx": 7315200, "cy": 914400 },
  "paragraphs": [
    { "runs": [{ "text": "Key Takeaways", "bold": true, "fontSize": 32 }], "align": "left" }
  ],
  "bodyPr": { "autoFit": "shrink" }
}
```

```json
{
  "op": "addElement",
  "target": { "slide": 1 },
  "kind": "roundRect",
  "offset": { "x": 914400, "y": 2286000, "cx": 3200400, "cy": 1600200 },
  "fill": "#E8F0FE",
  "stroke": { "color": "#1A73E8", "widthEmu": 12700 },
  "adjustments": { "adj": 16667 },
  "paragraphs": [
    { "runs": [{ "text": "Card title", "bold": true, "color": "#1A73E8" }], "align": "center" }
  ]
}
```

```json
{
  "op": "addElement",
  "target": { "slide": 1 },
  "kind": "lineArrow",
  "offset": { "x": 4572000, "y": 3048000, "cx": 1828800, "cy": 0 }
}
```

Common mistakes

- Building whole pages element by element on an empty deck: use `generate_deck`; `addElement` is for adding to an already designed page.
- Pixel frames: convert with the px-to-EMU factor from `read_slide`.
- `autoFit` values other than `"shrink"`/`"resize"`.

### addPicture (not-ai-callable)

`{bytes:base64|dataURL,ext?,offset} — use insert_web_image instead`

Embeds image bytes as a new picture. The model inserts images through
`insert_web_image` (URL) so the main process downloads and center-crops them.

### replacePicture (not-ai-callable)

`{bytes:base64|dataURL,ext?} — use replace_image instead`

Swaps a picture's bytes in place. The model uses `replace_image` with a URL.

### addTable

`{rows,cols,offset:{x,y,cx,cy},colWidthsEmu?,rowHeightsEmu?,cellProps?:[[{gridSpan?,rowSpan?,hMerge?,vMerge?,anchor?}]]}`

Inserts a native table with PowerPoint's default style and empty cells. Fill
cells afterwards with `setTableCell` (one op per cell, same transaction is
fine once you know the table id from the first result).

| Field         | Type                                                                | Notes                                                                    |
| ------------- | ------------------------------------------------------------------- | ------------------------------------------------------------------------ |
| rows, cols    | integer >= 1                                                        |                                                                          |
| offset        | `{x, y, cx, cy}` EMU                                                | Columns and rows split the frame equally unless widths/heights are given |
| colWidthsEmu  | number[]                                                            | Exactly `cols` positive values                                           |
| rowHeightsEmu | number[]                                                            | Exactly `rows` positive values                                           |
| cellProps     | array of rows of `{gridSpan?, rowSpan?, hMerge?, vMerge?, anchor?}` | Row-major; `anchor` `"t"`/`"ctr"`/`"b"`; use for pre-merged headers      |

```json
{
  "op": "addTable",
  "target": { "slide": 1 },
  "rows": 3,
  "cols": 3,
  "offset": { "x": 914400, "y": 1371600, "cx": 7315200, "cy": 2286000 }
}
```

```json
{
  "op": "addTable",
  "target": { "slide": 1 },
  "rows": 2,
  "cols": 2,
  "offset": { "x": 914400, "y": 1371600, "cx": 5486400, "cy": 1371600 },
  "colWidthsEmu": [1828800, 3657600],
  "cellProps": [
    [{ "gridSpan": 2 }, { "hMerge": true }],
    [{}, {}]
  ]
}
```

Common mistakes

- `colWidthsEmu` with the wrong count: it must list exactly one width per column.
- Passing cell text here: the op creates empty cells; write text with `setTableCell`.

### addChart

`{kind:"bar"|"barStacked"|"line"|"area"|"pie"|"doughnut"|"scatter"|"radar"|"comboBarLine",categories:[…],series:[{name,values:[…]}],offset,title?,colorScheme?:["#RRGGBB"],holeSizePct?}`

Inserts a native, PowerPoint-editable chart.

| Field       | Type                                              | Notes                                                                                                                                                                                                     |
| ----------- | ------------------------------------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| kind        | see signature                                     | `comboBarLine` draws the last series as a line on a secondary axis                                                                                                                                        |
| categories  | string[]                                          | At least one; the x-axis labels                                                                                                                                                                           |
| series      | array of `{name, values:number[]}`                | At least one; each `values` length should equal `categories` length                                                                                                                                       |
| offset      | `{x, y, cx, cy}` EMU                              |                                                                                                                                                                                                           |
| title       | string                                            | Optional chart title                                                                                                                                                                                      |
| colorScheme | `"#RRGGBB"[]`                                     | Optional series colors, cycled; omit to follow the theme                                                                                                                                                  |
| holeSizePct | number                                            | Doughnut hole size (10..90)                                                                                                                                                                               |
| barDir      | `"bar"`                                           | Optional: horizontal bars instead of columns                                                                                                                                                              |
| dataSource  | `"user"` / `"document"` / `"search"` / `"sample"` | Where the numbers came from; checked by apply_ops before the batch runs (not part of the op itself). `"search"` needs a web_search in this conversation; `"sample"` figures must be disclosed to the user |

```json
{
  "op": "addChart",
  "target": { "slide": 1 },
  "kind": "bar",
  "title": "Revenue by quarter",
  "categories": ["Q1", "Q2", "Q3", "Q4"],
  "series": [{ "name": "Revenue", "values": [120, 135, 150, 170] }],
  "offset": { "x": 914400, "y": 1371600, "cx": 7315200, "cy": 3657600 },
  "dataSource": "user"
}
```

```json
{
  "op": "addChart",
  "target": { "slide": 1 },
  "kind": "doughnut",
  "holeSizePct": 55,
  "categories": ["Direct", "Partner", "Online"],
  "series": [{ "name": "Share", "values": [45, 30, 25] }],
  "offset": { "x": 2743200, "y": 1371600, "cx": 3657600, "cy": 3657600 },
  "colorScheme": ["#1A73E8", "#34A853", "#FBBC04"]
}
```

Common mistakes

- Inventing numbers: every value needs a source (user, deck, or a search in this conversation); with none, use illustrative values and say so to the user.
- `values` shorter or longer than `categories`.
- Editing an existing chart with `addChart`: use the `edit_chart` tool, which keeps position and z-order.

### addSmartArt

`{layout,items:[…],offset}`

Inserts a SmartArt-style diagram built from grouped shapes.

| Field  | Type                                                                                     | Notes                     |
| ------ | ---------------------------------------------------------------------------------------- | ------------------------- |
| layout | `"list"` / `"process"` / `"cycle"` / `"hierarchy"` / `"pyramid"` / `"matrix"` / `"venn"` |                           |
| items  | string[]                                                                                 | 1..8 node texts, in order |
| offset | `{x, y, cx, cy}` EMU                                                                     |                           |

```json
{
  "op": "addSmartArt",
  "target": { "slide": 1 },
  "layout": "process",
  "items": ["Discover", "Design", "Deliver"],
  "offset": { "x": 914400, "y": 1828800, "cx": 7315200, "cy": 1828800 }
}
```

### addMedia (not-ai-callable)

`{kind:"video"|"audio",bytes:base64|dataURL,ext,offset}`

Embeds a video or audio file; bytes payload from the UI file picker.

### addModel3d (not-ai-callable)

`{bytes,ext,offset} — bytes payload`

Embeds a 3D model; bytes payload from the UI file picker.

### pasteElements (not-ai-callable)

`{items,dx,dy} — clipboard payload`

Pastes copied elements; the payload comes from the internal clipboard.
