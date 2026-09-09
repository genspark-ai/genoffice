# Slide ops

> Whole-page operations: delete, duplicate, insert blank, move, background, hide, transition, auto-advance, speaker notes, comments.

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

### addSlideWithLayout (not-ai-callable)

`{layoutPath} — layout part path; not discoverable from tool results`

Inserts a page from a specific layout part; the UI's layout gallery uses it.

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

### setSlideLayout (not-ai-callable)

`{layoutPath?} — layout part path; not discoverable from tool results`

Re-links the page to another layout part; the UI's layout gallery uses it.

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
