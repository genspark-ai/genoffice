# Deck ops

> Presentation-wide operations with no target: page size, find and replace, header/footer, theme, sections.

Deck ops carry no `target`. They affect every slide (or the presentation's
metadata), so read the deck outline first and confirm the scope with the user
when the request is ambiguous.

### setSlideSize

`{cx,cy} — EMU page size, whole deck`

Changes the page size of the whole deck. Element positions are not rescaled;
follow up with `execute_slide_script` when content must be re-laid out.
Common sizes: 16:9 is 12192000 x 6858000, 4:3 is 9144000 x 6858000.

```json
{ "op": "setSlideSize", "cx": 12192000, "cy": 6858000 }
```

### findReplace

`{find,replace,matchCase?,slideIndex?} — whole deck unless slideIndex`

Text substitution inside runs across text boxes, shapes, table cells and direct
group children. Matches that span two differently formatted runs are not
replaced.

| Field      | Type             | Notes                                       |
| ---------- | ---------------- | ------------------------------------------- |
| find       | non-empty string |                                             |
| replace    | string           | May be empty to delete the text             |
| matchCase  | boolean          | Default false                               |
| slideIndex | integer          | Restrict to one page                        |
| elementId  | string           | Restrict to one element (with `slideIndex`) |
| firstOnly  | boolean          | Replace only the first match                |

```json
{ "op": "findReplace", "find": "2023", "replace": "2024" }
```

```json
{
  "op": "findReplace",
  "find": "Acme Corp",
  "replace": "Acme Corporation",
  "matchCase": true,
  "slideIndex": 0
}
```

Common mistakes

- Expecting a report of zero changes: the op fails when nothing matched, so a "not found" error is normal feedback, not a fault.

### applyHeaderFooter

`{settings:{footer?,slideNum?,date?,dateAuto?}} — every slide`

Footer text, slide numbers and date on every page (through the master's
placeholders).

| Field             | Type             | Notes                                         |
| ----------------- | ---------------- | --------------------------------------------- |
| settings.footer   | string or `null` | `null` or empty hides the footer              |
| settings.slideNum | boolean          | Show page numbers                             |
| settings.date     | string or `null` | Fixed date text; `null` or empty hides it     |
| settings.dateAuto | boolean          | Dynamic date field that updates in PowerPoint |

```json
{
  "op": "applyHeaderFooter",
  "settings": { "footer": "Acme Corporation - Confidential", "slideNum": true }
}
```

Common mistakes

- Re-sending identical settings: the op reports "nothing changed" as an error; skip it when the values are already in place.

### setSections

`{sections:[{id,name,slideIndices},…]} — full replace`

Replaces the whole section list. Each section lists the 0-based indices of its
slides in order; every slide should belong to exactly one section. Ids are
GUIDs in braces; reuse existing ids to keep a section's identity, mint a new
GUID for a new one.

```json
{
  "op": "setSections",
  "sections": [
    { "id": "SECTION_ID", "name": "Introduction", "slideIndices": [0] },
    { "id": "{6F1D2B0A-4C3E-4F58-9A21-0D7E5B3C9A11}", "name": "Details", "slideIndices": [1] }
  ]
}
```

Related: `addSection`, `renameSection`, `removeSection`, `moveSection` for single changes.

### applyTheme

`{name,colors:{dk1?,lt1?,dk2?,lt2?,accent1?..accent6?,hlink?,folHlink?},majorFont?,minorFont?} — "#RRGGBB" slots, whole deck`

Rewrites the theme color scheme (and optionally the heading/body fonts) of the
whole deck, remapping explicit colors that matched the old scheme. Pages
without their own background get the new `lt1` color.

| Field     | Type                                  | Notes                                                                                                  |
| --------- | ------------------------------------- | ------------------------------------------------------------------------------------------------------ |
| name      | non-empty string                      | Theme name written into the file                                                                       |
| colors    | object of scheme slots to `"#RRGGBB"` | Slots: `dk1`, `lt1`, `dk2`, `lt2`, `accent1`..`accent6`, `hlink`, `folHlink`; six hex digits, no alpha |
| majorFont | string                                | Heading font                                                                                           |
| minorFont | string                                | Body font                                                                                              |

```json
{
  "op": "applyTheme",
  "name": "Ocean",
  "colors": {
    "dk1": "#0B1F3A",
    "lt1": "#FFFFFF",
    "dk2": "#1F3B5C",
    "lt2": "#E8F0FE",
    "accent1": "#1A73E8",
    "accent2": "#34A853",
    "accent3": "#FBBC04",
    "accent4": "#EA4335",
    "accent5": "#9AA0A6",
    "accent6": "#5F6368",
    "hlink": "#1A73E8",
    "folHlink": "#7B1FA2"
  },
  "majorFont": "Georgia",
  "minorFont": "Calibri"
}
```

Common mistakes

- Alpha in scheme colors (`"#RRGGBBAA"`): theme slots take six hex digits only.
- Passing `colors` as an array: it is a slot-to-color object.

### addSection

`{atSlideIndex,name}`

Starts a new section at the given slide (that slide and the following ones up to
the next section move into it).

```json
{ "op": "addSection", "atSlideIndex": 1, "name": "Appendix" }
```

### renameSection

`{id,name}`

```json
{ "op": "renameSection", "id": "SECTION_ID", "name": "Overview" }
```

### removeSection

`{id} — keeps the slides`

Removes the section header; its slides join the previous section.

```json
{ "op": "removeSection", "id": "SECTION_ID" }
```

### moveSection

`{id,dir:"up"|"down"}`

Swaps the section (with all its slides) with its neighbor.

```json
{ "op": "moveSection", "id": "SECTION_ID", "dir": "down" }
```
