# Slide ops

> Whole-page operations: delete, duplicate, insert blank or by layout, move, background, hide, transition, auto-advance, animations, speaker notes, comments.

Slide ops take `target:{slide}` only. `slide` is a 0-based index or a durable
`"s_<n>"` id from the outline. Structural ops (delete, duplicate, insert, move)
shift later indices; when you batch several in one transaction, later ops are
still resolved against the live deck, so prefer durable ids or order the batch
from the last page to the first.

### deleteSlide

`{}`

Removes the page. A deck must keep at least one slide.

```json
{ "op": "deleteSlide", "target": { "slide": 1 } }
```

### duplicateSlide

`{}`

Inserts a copy right after the page. `clearText:true` empties the text boxes
so the copy serves as a layout-preserving blank.
The copy lands at the source index + 1; the result echoes only the source
target, so address the new page by that index when you fill it.

```json
{ "op": "duplicateSlide", "target": { "slide": 0 } }
```

```json
{ "op": "duplicateSlide", "target": { "slide": 0 }, "clearText": true }
```

### addBlankSlide

`{} — inserts after target.slide`

Inserts a new blank page (same layout and background as `target.slide`) right
after it.

```json
{ "op": "addBlankSlide", "target": { "slide": "s_1" } }
```

### addSlideWithLayout

`{layout:<name>|<index>} — inserts after target.slide (default: after the last slide)`

Inserts a page bound to one of the deck's layouts, with that layout's
placeholders as empty prompt boxes. `layout` is the layout's gallery name or
its 0-based index as listed by `slides read --layouts`; the error names every
layout when the reference misses. The new slide's id is reported in `created`.

```json
{ "op": "addSlideWithLayout", "target": { "slide": "s_1" }, "layout": "Blank" }
```

Common mistakes

- Guessing layout names from another deck: list them with `slides read --layouts` first.
- Filling the new placeholders in the same batch: they get ids only after the insert; `slides read` the new slide, then `setText`.

### pasteSlide (not-ai-callable)

`{afterIndex,bundle|png,mode?} — clipboard payload`

Pastes a copied slide bundle from the internal clipboard.

### insertSlidePptx (not-ai-callable)

`{source,at?,replace?} — generated-page landing payload (use generate_deck/regenerate_slide)`

Lands a generated one-slide pptx into the deck; the generation tools call it.

### moveSlide

`{to} — 0-based destination index`

Moves the page to a new position.

```json
{ "op": "moveSlide", "target": { "slide": 1 }, "to": 0 }
```

### setSlideLayout

`{layout?:<name>|<index>} — omitted: snap placeholders back to the current layout`

Re-links the page to another layout (existing shapes stay where they are;
missing layout placeholders are added as empty prompt boxes). Without `layout`
the current layout is re-applied, moving placeholders back to their layout
positions.

```json
{ "op": "setSlideLayout", "target": { "slide": 0 }, "layout": 0 }
```

### setBackground

`{kind:"solid"|"gradient"|"reset"|"graphics"|"image",color?,from?,to?,angleDeg?,radial?,hidden?} — image kind needs a bytes source (use set_slide_background)`

Page background. Full-bleed decorative rectangles that act as a backdrop are
recolored along with it. One slide per op; for "all pages" send one op per
slide in the same transaction.

| Kind         | Fields                               | Notes                                                                                 |
| ------------ | ------------------------------------ | ------------------------------------------------------------------------------------- |
| `"solid"`    | `color`                              | `"#RRGGBB"`                                                                           |
| `"gradient"` | `from`, `to`, `angleDeg?`, `radial?` | Two-stop gradient; `angleDeg` in degrees (default 0); `radial:true` ignores the angle |
| `"reset"`    | none                                 | Remove the page override and inherit the layout/master background                     |
| `"graphics"` | `hidden`                             | `true` hides the master's background graphics on this page                            |
| `"image"`    | `source`                             | Bytes or media-part payload; not usable from `apply_ops`                              |

```json
{ "op": "setBackground", "target": { "slide": 0 }, "kind": "solid", "color": "#0B1F3A" }
```

```json
{
  "op": "setBackground",
  "target": { "slide": 1 },
  "kind": "gradient",
  "from": "#0B1F3A",
  "to": "#1A73E8",
  "angleDeg": 90
}
```

```json
{ "op": "setBackground", "target": { "slide": 1 }, "kind": "reset" }
```

Common mistakes

- Dark backgrounds without lightening the text: follow up with `setFont` color changes on the page's text.
- `kind:"image"` from the model: use the picture tools; the op needs bytes.

### setHidden

`{hidden:boolean}`

Hides or shows the page in the slide show (it stays in the deck).

```json
{ "op": "setHidden", "target": { "slide": 1 }, "hidden": true }
```

### setTransition

`{kind:"none"|"fade"|"push"|"wipe"|…}`

Slide transition. Accepted kinds: `none`, `morph`, `fade`, `push`, `wipe`,
`split`, `circle`, `cover`, `pull`, `dissolve`, `zoom`, `random`.

```json
{ "op": "setTransition", "target": { "slide": 0 }, "kind": "fade" }
```

Common mistakes

- Names from other tools such as `"slide"` or `"cut"`: only the kinds listed above are accepted; the error lists them.

### setAdvanceTime

`{ms:number|null} — auto-advance; null clears`

Automatic advance after the given milliseconds; `null` returns to click-to-advance.

```json
{ "op": "setAdvanceTime", "target": { "slide": 0 }, "ms": 5000 }
```

### setAnimations (not-ai-callable)

`{items:[{spid,effect,trigger,durationMs,delayMs,…}]} — spid-addressed (cNvPr id), no id translation yet`

Rewrites the page's animation timeline; addressed by raw shape ids, so the UI
animation pane owns it for now.

### addAnimation

`{effect,kind?,trigger?,duration?,delay?,direction?,paragraph?,motionPath?,presetXml?,after?} — target:{slide, el}; one effect appended to the page timeline`

Adds one animation for the element at the end of the page's timeline (or right
after position `after`). The timeline is read back as `animations[]` in
`slides read`, each item with its `seq` (0-based position), `el`, `effect`,
`kind`, `trigger`, `durationMs`, `delayMs`.

| Field      | Type                                                                                                                                                                                                                                                                                   | Notes                                                                                                                                             |
| ---------- | -------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------- |
| effect     | entrance `appear` `fade` `flyIn` `wipe` `wipeDown` `splitIn` `bounce` `flipIn` `zoom`; emphasis `pulse` `spin` `grow` `teeter`; exit `disappear` `fadeOut` `flyOut` `wipeOut` `shrink` `zoomOut`; `motionPath`; media `mediaPlay` `mediaPause` `mediaStop` (video/audio elements only) | Required unless `presetXml` is given                                                                                                              |
| kind       | `"entrance"` / `"emphasis"` / `"exit"` / `"path"` / `"media"`                                                                                                                                                                                                                          | Optional cross-check; rejected when it does not match the effect                                                                                  |
| trigger    | `"onClick"` (default) / `"withPrev"` / `"afterPrev"`                                                                                                                                                                                                                                   | `withPrev` starts with the previous effect, `afterPrev` after it ends                                                                             |
| duration   | ms >= 0                                                                                                                                                                                                                                                                                | Default per effect (appear 0, fade 500, spin 2000, …); `durationMs` also accepted                                                                 |
| delay      | ms >= 0                                                                                                                                                                                                                                                                                | Default 0; `delayMs` also accepted                                                                                                                |
| direction  | `"top"` / `"bottom"` / `"left"` / `"right"`                                                                                                                                                                                                                                            | Only for `flyIn`, `flyOut`, `wipe`, `wipeOut` (default bottom)                                                                                    |
| paragraph  | 0-based index                                                                                                                                                                                                                                                                          | Animate one paragraph of a text body instead of the whole shape                                                                                   |
| motionPath | `"M 0 0 L 0.25 0"`                                                                                                                                                                                                                                                                     | With `effect:"motionPath"`; coordinates are fractions of the slide size                                                                           |
| presetXml  | one `<p:par>…</p:par>` block                                                                                                                                                                                                                                                           | Verbatim effect exported from PowerPoint for presets outside the list; every `<p:spTgt spid>` must be this element's shape id; kept byte-for-byte |
| after      | seq                                                                                                                                                                                                                                                                                    | Insert right after that timeline position instead of appending                                                                                    |

```json
{
  "op": "addAnimation",
  "target": { "slide": 0, "el": "e_TEXT" },
  "effect": "flyIn",
  "direction": "left",
  "trigger": "afterPrev",
  "duration": 600
}
```

### removeAnimation

`{seq} — target:{slide}; or target:{slide, el} to drop every animation of that element`

Removes one timeline item by `seq`, or all items targeting an element. Later
items renumber; read the page again before the next batch.

```json
{ "op": "removeAnimation", "target": { "slide": 0, "el": "e_TEXT" } }
```

### reorderAnimation

`{seq,to} — target:{slide}; move a timeline item`

Moves the item at `seq` to position `to` (both 0-based, inside the timeline).

```json
{ "op": "reorderAnimation", "target": { "slide": 0 }, "seq": 2, "to": 0 }
```

### setNotes

`{text} — speaker notes`

Replaces the page's speaker notes (paragraphs separated by newlines). An empty
string clears them. Notes never affect the canvas.

```json
{
  "op": "setNotes",
  "target": { "slide": 0 },
  "text": "Open with the headline number.\nPause for questions before the next section."
}
```

### addComment

`{text,author}`

Adds a review comment to the page.

```json
{
  "op": "addComment",
  "target": { "slide": 0 },
  "text": "Consider a stronger verb in the title.",
  "author": "AI Assistant"
}
```

### deleteComment

`{authorId,idx}`

Removes a comment identified by its author id and index within that author's
comments on the page (both come from the comments pane data).

```json
{ "op": "deleteComment", "target": { "slide": 0 }, "authorId": 0, "idx": 0 }
```
