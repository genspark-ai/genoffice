# Element ops

> Change one existing element: fill, stroke, position and size, crop, opacity, z-order, grouping, shape geometry, text-body properties, effects, hyperlinks, deletion.

Every op here takes `target:{slide, el}` unless its signature says otherwise
(`flipElements` and `groupElements` take a list in `els`). Ops that say
"group children" accept a direct child of a group when you add
`group:"<group id>"` and put the child id in `target.el`.

Geometry is document-space EMU: 1 pt = 12700 EMU, 1 px = 9525 EMU on a
standard 1280 px wide deck (`read_slide` reports the exact factor for the
current deck). Angles are degrees unless the field says otherwise.

### deleteElement

`{} — no fields besides target`

Removes a top-level element. To delete one member of a group, `ungroupElement`
the group first (its children become top-level and get fresh ids).

```json
{ "op": "deleteElement", "target": { "slide": 0, "el": "e_SHAPE" } }
```

Common mistakes

- Deleting a group child directly: the id does not resolve at top level; ungroup first or delete the whole group.

### setFill

`{fill:"#RRGGBB"|"none"|{stops:[{pos,color}],angle?}}`

Solid, translucent or gradient fill of a text box or shape. Pictures, tables,
charts and connectors have no fill. Picture fills are `setImageFill` (UI only).

| Field       | Type                                                                      | Notes                                                      |
| ----------- | ------------------------------------------------------------------------- | ---------------------------------------------------------- |
| fill        | `"none"`, `"#RRGGBB"`, `"#RRGGBBAA"` or `{stops, angle?, radial?, path?}` | Alpha byte `AA`: 00 transparent .. FF opaque               |
| fill.stops  | array of `{pos, color}`                                                   | At least two; `pos` 0..1                                   |
| fill.angle  | number                                                                    | Linear gradient angle in 1/60000 degree (90 deg = 5400000) |
| fill.radial | boolean                                                                   | Radial gradient instead of linear                          |
| group       | string                                                                    | Only for a direct child of a group                         |

```json
{ "op": "setFill", "target": { "slide": 0, "el": "e_SHAPE" }, "fill": "#1A73E8" }
```

```json
{
  "op": "setFill",
  "target": { "slide": 0, "el": "e_SHAPE" },
  "fill": {
    "stops": [
      { "pos": 0, "color": "#1A73E8" },
      { "pos": 1, "color": "#FFFFFF" }
    ],
    "angle": 5400000
  }
}
```

Common mistakes

- `null` to remove a fill: use `"none"`.
- Gradient with one stop: at least two stops are required.
- Targeting a picture: use `setPictureOpacity` / `setImageFill` instead.

Related: `setStroke`, `setBackground` (page background, a slide op).

### setStroke

`{stroke:{color:"#RRGGBB",widthEmu}|null} — null removes the outline`

Outline of a text box, shape or picture (picture border).

| Field           | Type                                                        | Notes                                                          |
| --------------- | ----------------------------------------------------------- | -------------------------------------------------------------- |
| stroke          | object or `null`                                            | `null` removes the outline                                     |
| stroke.color    | `"#RRGGBB"`                                                 | Required when stroke is an object                              |
| stroke.widthEmu | number > 0                                                  | 12700 EMU = 1 pt                                               |
| stroke.dash     | string                                                      | Preset dash name (`"dash"`, `"sysDot"`, ...); `"solid"` clears |
| stroke.cap      | `"flat"` / `"rnd"` / `"sq"`                                 | Optional                                                       |
| stroke.compound | `"sng"` / `"dbl"` / `"thickThin"` / `"thinThick"` / `"tri"` | Optional                                                       |
| stroke.join     | `"round"` / `"bevel"` / `"miter"`                           | Optional                                                       |
| stroke.gradient | `{stops:[{pos,color}], angle}`                              | Gradient line; angle in 1/60000 degree                         |
| group           | string                                                      | Only for a direct child of a group                             |

```json
{
  "op": "setStroke",
  "target": { "slide": 0, "el": "e_SHAPE" },
  "stroke": { "color": "#333333", "widthEmu": 19050 }
}
```

```json
{ "op": "setStroke", "target": { "slide": 0, "el": "e_PICTURE" }, "stroke": null }
```

Common mistakes

- Width in points: `widthEmu` is EMU; multiply points by 12700.
- `remove:true` or `"none"`: the outline is removed by passing `stroke: null`.

### setTransform

`{box:{x,y,cx,cy},rotDeg?} (group children: absBox instead of box, plus group:"<group id>")`

Moves, resizes and rotates an element. `box` is the full new frame in EMU; pass
all four numbers even when only one changes (read the current values from
`read_slide` or an `execute_slide_script` `els` entry).

| Field           | Type                 | Notes                                                                      |
| --------------- | -------------------- | -------------------------------------------------------------------------- |
| box             | `{x, y, cx, cy}` EMU | Required for top-level elements                                            |
| rotDeg          | number               | Degrees clockwise; omitted = 0                                             |
| absBox          | `{x, y, cx, cy}` EMU | Group children only: document-space frame (the op converts to child space) |
| group           | string               | Group children only                                                        |
| resizeTableGrid | boolean              | Tables only: redistribute columns/rows to the new frame                    |

```json
{
  "op": "setTransform",
  "target": { "slide": 0, "el": "e_TEXT" },
  "box": { "x": 914400, "y": 685800, "cx": 7315200, "cy": 1143000 }
}
```

```json
{
  "op": "setTransform",
  "target": { "slide": 0, "el": "e_CHILD" },
  "group": "e_GROUP",
  "absBox": { "x": 1828800, "y": 1828800, "cx": 1828800, "cy": 914400 },
  "rotDeg": 15
}
```

Common mistakes

- Pixel values in `box`: convert with the px-to-EMU factor from `read_slide`.
- Sending only `cx`/`cy`: the rect must carry all of x, y, cx, cy.
- For several elements with relative math (align, distribute) prefer `execute_slide_script`, which computes from live geometry.

### setConnectorEndpoints

`{p1:{x,y},p2:{x,y},start?:{targetId,idx}|null,end?:{targetId,idx}|null}`

Re-routes a connector between two points and optionally attaches its ends to
shapes so it follows later moves. The element must be a connector (line,
arrow, elbow or curved connector).

| Field      | Type                                 | Notes                                                                                                 |
| ---------- | ------------------------------------ | ----------------------------------------------------------------------------------------------------- |
| p1, p2     | `{x, y}` EMU                         | Required; frame and flips are derived from them                                                       |
| start, end | `{targetId, idx}`, `null` or omitted | Attach to connection point `idx` of an element; `null` detaches; omitted keeps the current attachment |

```json
{
  "op": "setConnectorEndpoints",
  "target": { "slide": 0, "el": "e_LINE" },
  "p1": { "x": 914400, "y": 914400 },
  "p2": { "x": 3657600, "y": 2743200 },
  "start": { "targetId": "e_SHAPE", "idx": 3 },
  "end": null
}
```

Common mistakes

- Using it on a plain shape: only connectors are accepted; for shapes use `setTransform`.
- A `targetId` that is not on the slide: the op fails instead of silently detaching; pass `null` to detach.

### flipElements

`{els:[id,…],axis:"h"|"v"} — target.el unused`

Mirrors one or more elements horizontally or vertically around their own
centers. Any element with a frame flips (shapes, text boxes, pictures, groups).

| Field | Type           | Notes                                             |
| ----- | -------------- | ------------------------------------------------- |
| els   | string[]       | Non-empty list of element ids on `target.slide`   |
| axis  | `"h"` or `"v"` |                                                   |
| group | string         | Optional: flip children inside that group instead |

```json
{ "op": "flipElements", "target": { "slide": 0 }, "els": ["e_SHAPE", "e_PICTURE"], "axis": "h" }
```

### setPictureSrcRect

`{srcRect:{l,t,r,b}|null} — crop fractions 0..1, null removes the crop`

Non-destructive crop of a picture: each side is the fraction of the source
image removed from that edge. The frame stays; the remaining image stretches to
fill it. Optionally pass `box` to shrink the frame in the same op.

| Field   | Type                         | Notes                                                                      |
| ------- | ---------------------------- | -------------------------------------------------------------------------- |
| srcRect | `{l?, t?, r?, b?}` or `null` | Each 0..1 (exclusive of 1); `l + r < 1` and `t + b < 1`; missing sides = 0 |
| box     | `{x, y, cx, cy}` EMU         | Optional new frame applied together with the crop                          |

```json
{
  "op": "setPictureSrcRect",
  "target": { "slide": 0, "el": "e_PICTURE" },
  "srcRect": { "l": 0.1, "t": 0, "r": 0.1, "b": 0.2 }
}
```

```json
{ "op": "setPictureSrcRect", "target": { "slide": 0, "el": "e_PICTURE" }, "srcRect": null }
```

Common mistakes

- Percentages such as 10: fractions are 0..1, so 10 percent is 0.1.
- Opposing crops summing to 1 or more: nothing would remain visible.

### setPictureOpacity

`{opacity:0..1}`

Whole-image transparency of a picture. `1` is fully opaque and removes the effect.

```json
{ "op": "setPictureOpacity", "target": { "slide": 0, "el": "e_PICTURE" }, "opacity": 0.6 }
```

Common mistakes

- Percent values: 60 percent is `0.6`.
- Non-picture targets: shapes take translucency through `setFill` with an `"#RRGGBBAA"` color.

### reorderElement

`{dir:"front"|"back"|"forward"|"backward"}`

Changes the z-order (stacking) of an element on its slide.

```json
{ "op": "reorderElement", "target": { "slide": 0, "el": "e_SHAPE" }, "dir": "back" }
```

Common mistakes

- Repeating `front`/`back` on an element already at that extreme fails; check the outline order first.

### groupElements

`{els:[id,id,…]} — at least two ids; target.el unused`

Groups two or more top-level elements. The result reports the new group id in
`created`; the children keep editable through `group:"<group id>"`.

```json
{ "op": "groupElements", "target": { "slide": 0 }, "els": ["e_SHAPE", "e_TEXT"] }
```

### ungroupElement

`{} — target.el is the group`

Dissolves a group; its direct children become top-level elements with new ids
(positions preserved). Read the slide again before editing them.

```json
{ "op": "ungroupElement", "target": { "slide": 0, "el": "e_GROUP" } }
```

### setShapeGeometry

`{prst:"<OOXML preset geometry name>"}`

Swaps the preset geometry of a shape (for example `rect` to `roundRect`,
`ellipse`, `triangle`, `rightArrow`, `chevron`, `star5`, `hexagon`). Text,
fill, stroke and frame stay.

```json
{ "op": "setShapeGeometry", "target": { "slide": 0, "el": "e_SHAPE" }, "prst": "hexagon" }
```

Common mistakes

- Applying it to a picture, table or chart: only shapes with a preset geometry qualify.

### setShapeAdjust

`{adjust:{<gdName>:number,…}} — preset-geometry avLst values (group children: add group)`

Sets adjustment handles of a preset geometry, for example the corner radius of
a rounded rectangle (`adj`, 0..50000) or arrow head proportions (`adj1`,
`adj2`).

```json
{ "op": "setShapeAdjust", "target": { "slide": 0, "el": "e_SHAPE" }, "adjust": { "adj": 16667 } }
```

Common mistakes

- Fractions like 0.25: adjustment values are the raw avLst numbers (25 percent of the range is 25000 for most presets).

### setTextAnchor

`{anchor:"top"|"middle"|"bottom"}`

Vertical alignment of the text inside its box.

```json
{ "op": "setTextAnchor", "target": { "slide": 0, "el": "e_TEXT" }, "anchor": "middle" }
```

### setTextBodyProps

`{props:{vert?:"horz"|"eaVert"|"vert"|"vert270"|"wordArtVert",autofit?:"none"|"shrink"|"resize",insets?:{l?,t?,r?,b?} (EMU),wrap?:boolean}}`

Text-body properties: writing direction, autofit behavior, internal margins and
word wrap. Pass at least one property.

| Field         | Type                                                             | Notes                                                           |
| ------------- | ---------------------------------------------------------------- | --------------------------------------------------------------- |
| props.vert    | `"horz"` / `"eaVert"` / `"vert"` / `"vert270"` / `"wordArtVert"` | `"horz"` clears vertical text                                   |
| props.autofit | `"none"` / `"shrink"` / `"resize"`                               | `shrink` scales text down on overflow, `resize` grows the shape |
| props.insets  | `{l?, t?, r?, b?}` EMU                                           | Only the given sides change                                     |
| props.wrap    | boolean                                                          | `false` lets a line overflow instead of wrapping                |

```json
{
  "op": "setTextBodyProps",
  "target": { "slide": 0, "el": "e_TEXT" },
  "props": { "autofit": "shrink", "insets": { "l": 91440, "r": 91440 } }
}
```

### setEffects

`{effects:{shadow?:{color:"#RRGGBB(AA)",blurRad,dist,dirDeg,inner?,sx?,sy?,kxDeg?}|null,glow?:{color,radius}|null,reflection?:{blurRad,startA(0..1),endPos(0..1),dist}|null,softEdge?:EMU|null}} — null clears; EMU distances (12700/pt)`

Shadow, glow, reflection and soft-edge effects. Each key you pass is replaced;
`null` clears that effect; keys you omit stay as they are. Pass at least one key.

| Field              | Type                                                                         | Notes                                                                         |
| ------------------ | ---------------------------------------------------------------------------- | ----------------------------------------------------------------------------- |
| effects.shadow     | `{color, blurRad, dist, dirDeg, inner?, sx?, sy?, kxDeg?, kyDeg?}` or `null` | `color` may carry alpha `"#RRGGBBAA"`; `blurRad`/`dist` EMU; `dirDeg` degrees |
| effects.glow       | `{color, radius}` or `null`                                                  | `radius` EMU                                                                  |
| effects.reflection | `{blurRad, startA, endPos, dist}` or `null`                                  | `startA`/`endPos` fractions 0..1                                              |
| effects.softEdge   | number EMU or `null`                                                         | Feather radius                                                                |

```json
{
  "op": "setEffects",
  "target": { "slide": 0, "el": "e_SHAPE" },
  "effects": { "shadow": { "color": "#00000066", "blurRad": 50800, "dist": 38100, "dirDeg": 90 } }
}
```

```json
{
  "op": "setEffects",
  "target": { "slide": 0, "el": "e_SHAPE" },
  "effects": { "shadow": null, "glow": null }
}
```

Common mistakes

- An empty `effects` object: at least one of shadow/glow/reflection/softEdge is required.
- Points for `dist`/`blurRad`: these are EMU (12700 per point).

### setLink

`{link:{kind:"url",url}|{kind:"slide",slideIndex}|null}`

Hyperlink on a whole element (click action). `null` removes the link. Links on
individual runs are set through `setText` run `link` fields.

```json
{
  "op": "setLink",
  "target": { "slide": 0, "el": "e_SHAPE" },
  "link": { "kind": "url", "url": "https://example.com" }
}
```

```json
{
  "op": "setLink",
  "target": { "slide": 0, "el": "e_SHAPE" },
  "link": { "kind": "slide", "slideIndex": 1 }
}
```

### setImageFill (not-ai-callable)

`{source:{mediaPath}|{bytes:base64|dataURL,ext},tile?} — use the image tools instead`

Picture or texture fill of a shape. The payload is image bytes or an existing
media part path, which the model cannot produce; the UI's fill picker uses it.
