# @genoffice/ppt2pptx

Legacy PowerPoint 97-2003 import: parse `.ppt` (OLE2/CFB + MS-PPT records) and rebuild it as `.pptx`. Staged like `@genoffice/doc2docx`: stage 1 laid the editable-text foundation, stage 2 adds the drawing layer. The CLI and the Slides open flow convert through `convertPptToPptx`.

## Stage 2 carries

- the **Escher shape tree** of every slide and master (`escher.ts`): autoshapes at their anchors with the msosp type mapped onto the DrawingML preset, groups flattened through their coordinate spaces, freeforms as custom paths (`pVertices` / `pSegmentInfo`, including the arrays whose declared length leaves out the 6-byte header), rotation and flips, hidden and deleted shapes dropped
- **fills and lines** from the property table: solid, gradient (`fillShadeColors` or the colour pair with focus), picture and texture fills, opacity; OfficeArt colours resolved against the slide's or master's colour scheme, with PowerPoint's scheme-slot defaults (fill → slot 4, line → slot 1)
- **pictures** from the `Pictures` stream (PNG/JPEG/EMF/WMF/DIB/TIFF blips, zlib metafiles, DIB → BMP), indexed through the drawing group's FBSE list with a stream-order fallback, cropped as the shape says
- **text formatting** (`text-style.ts`): the StyleTextPropAtom paragraph and character runs over the master's `TxMasterStyleAtom` per-level defaults — size, bold/italic/underline, colour, font (through the FontCollection), alignment, bullets, spacing, indents — plus the body's insets, anchor, wrap, autofit and vertical flow
- **placeholders** inherit anchor and properties from the master placeholder of their type (boolean property sets merged bit by bit); a title/body text shape without a placeholder atom matches by text type; date / slide-number placeholders holding only a field marker are dropped
- the **slide background** (own or master's, per the SlideAtom flags) and the master's decorations when the slide follows them

## Stage 1 carries

- the **change directory** the way PowerPoint itself resolves it: CurrentUserAtom → UserEditAtom → PersistDirectoryAtom (`persistId:20|cPersist:12` entries, not an implicit id array) → older generations through `offsetLastEdit` — a sequential walk would read stale generations
- slide order from the DocumentContainer's SlideListWithText (instance 0 = slides; each SlidePersistAtom carries the persist id of its SlideContainer), the slide surface from the DocumentAtom (576 master units per inch → EMU)
- every slide's text sequences as **editable text boxes**: the slide's SlideListWithText outline entries (where real PowerPoint stores placeholder title/body text) plus the text boxes nested five levels down under PPDrawing/Escher (OfficeArtClientTextbox), with PowerPoint's runes normalized (`\u000B` → newline, the bullet-in-paragraph marker → space, cp1252 smart quotes named)
- the OLE2/CFB container opener is the **shared** one in `@genoffice/compound-file` (`openOle2Container`: 8-byte signature, 512-byte header, input/stream caps), with the PowerPoint-specific checks (presentation stream, encrypted headerToken 0xF3D1C4DF) layered on top
- an explicit error model (`LegacyPptError`, stable codes, including `encrypted`)
- a regression test on the in-repo real deck `packages/file-parse/tests/fixtures/legacy-sample.ppt` (plus the byte builder for the change-directory cases)

## Not carried yet

notes, animations and transitions, hyperlinks, shadows and 3-D, line dashes and arrowheads, pattern fills (drawn solid), OLE objects beyond their preview picture, tables beyond their cell shapes, headers/footers text fields.
