# Text ops

> Replace or restyle the text of a text box, shape or table: whole-body rewrite (setText), run-level font patch (setFont), paragraph format patch (setParagraphFormat).

All three take `target:{slide, el}`. `el` is an element id from the outline or
`read_slide` (`e_*` ids are durable). For a direct child of a group put the
child id in `target.el` and add `group:"<group id>"`.

Font sizes are points. Colors are `"#RRGGBB"`.

### setText

`{paragraphs:[{runs:[{text,bold?,italic?,fontSize?,color?}],align?}]} (group children: add group:"<group id>")`

Replaces the element's entire text with the given paragraphs. Runs you do not
restyle inherit the formatting of the run they replace (position-wise), so a
plain `{text}` run keeps the original size, color and font. A shape that has
never had text gets a text body created (autoshapes default to centered text).

| Field                                                    | Type                                                                                          | Notes                                                                                             |
| -------------------------------------------------------- | --------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| paragraphs                                               | array                                                                                         | Required, one object per paragraph; an empty array clears the text                                |
| paragraphs[].runs                                        | array of `{text, bold?, italic?, underline?, strike?, fontSize?, fontFamily?, color?, link?}` | `fontSize` in pt; `link` is `{kind:"url",url}` or `{kind:"slide",slideIndex}` or `null` to remove |
| paragraphs[].align                                       | `"left"` / `"center"` / `"right"` / `"justify"`                                               | Optional                                                                                          |
| paragraphs[].level                                       | 0..8                                                                                          | Indent level for multi-level lists                                                                |
| paragraphs[].bullet                                      | `"char"` / `"number"` / `"none"`                                                              | Optional; `bulletChar` sets a custom glyph                                                        |
| paragraphs[].lineSpacingPct, spaceBeforePt, spaceAfterPt | number                                                                                        | Optional paragraph spacing                                                                        |
| paragraphs[].rtl                                         | boolean                                                                                       | Optional base direction                                                                           |
| group                                                    | string                                                                                        | Only for a direct child of a group: the group id                                                  |

```json
{
  "op": "setText",
  "target": { "slide": 0, "el": "e_TEXT" },
  "paragraphs": [
    { "runs": [{ "text": "Quarterly Review", "bold": true, "fontSize": 36 }], "align": "left" },
    { "runs": [{ "text": "Q3 highlights and outlook" }] }
  ]
}
```

```json
{
  "op": "setText",
  "target": { "slide": 0, "el": "e_CHILD" },
  "group": "e_GROUP",
  "paragraphs": [{ "runs": [{ "text": "Step 1" }] }]
}
```

Common mistakes

- Passing only the changed paragraph: the array replaces the whole body, so send every paragraph that should remain.
- Targeting a picture, chart or connector: only text boxes, shapes and table cells (`setTableCell`) hold text.
- Writing text into a group child without `group`: the child id alone does not resolve on the slide.

Related: `setFont` (restyle without changing the words), `setTableCell` (table cells), `findReplace` (deck-wide substitutions).

### setFont

`{font:{fontFamily?,fontSizePt?,bold?,italic?,underline?,strike?,color?}} — merges onto every run of the element`

Applies the given properties to every run of the element (or every cell of a
table) and leaves everything else untouched. Use it for "make the title blue
and bold" style requests; use `setText` when the words change.

| Field                                | Type        | Notes                                          |
| ------------------------------------ | ----------- | ---------------------------------------------- |
| font.fontFamily                      | string      | Omit to keep the theme font (recommended)      |
| font.fontSizePt                      | number > 0  | Points                                         |
| font.bold, italic, underline, strike | boolean     | Only the keys you pass change                  |
| font.color                           | `"#RRGGBB"` | Explicit color; clears theme-color inheritance |
| group                                | string      | Only for a direct child of a group             |

```json
{
  "op": "setFont",
  "target": { "slide": 0, "el": "e_TEXT" },
  "font": { "color": "#1A73E8", "bold": true }
}
```

```json
{ "op": "setFont", "target": { "slide": 0, "el": "e_TABLE" }, "font": { "fontSizePt": 12 } }
```

Common mistakes

- Color names such as `"blue"`: only hex `"#RRGGBB"` is accepted.
- `fontSizePt` of 0 or negative: must be a positive number of points.
- Expecting a partial-run change: the patch is element-wide; for one word use `setText` with per-run styling.

### setParagraphFormat

`{format:{align?,bullet?,lineSpacingPct?,spaceBeforePt?,spaceAfterPt?}}`

Paragraph-level formatting for every paragraph of the element: alignment,
bullets, spacing, direction, indent level.

| Field                                     | Type                                            | Notes                                        |
| ----------------------------------------- | ----------------------------------------------- | -------------------------------------------- |
| format.align                              | `"left"` / `"center"` / `"right"` / `"justify"` |                                              |
| format.bullet                             | `"char"` / `"number"` / `"none"`                | `"char"` uses `bulletChar` or a round bullet |
| format.bulletChar                         | string                                          | Custom glyph, with `bullet:"char"`           |
| format.bulletSizePct                      | number >= 0                                     | Percent of text size (100 = same)            |
| format.bulletColor                        | `"#RRGGBB"`                                     |                                              |
| format.bulletHangEmu                      | number >= 0                                     | Hanging indent in EMU                        |
| format.lineSpacingPct                     | number >= 0                                     | 100 = single spacing                         |
| format.spaceBeforePt, format.spaceAfterPt | number >= 0                                     | Points                                       |
| format.rtl                                | boolean                                         | Base direction                               |
| format.indentDelta                        | 1 or -1                                         | Shift the list level (clamped 0..8)          |
| group                                     | string                                          | Only for a direct child of a group           |

```json
{
  "op": "setParagraphFormat",
  "target": { "slide": 0, "el": "e_TEXT" },
  "format": { "bullet": "char", "lineSpacingPct": 120, "spaceAfterPt": 6 }
}
```

Common mistakes

- Negative spacing values: all `...Pct` / `...Pt` / `...Emu` fields must be >= 0.
- Using `setParagraphFormat` to change font size: that is a run property, use `setFont`.
