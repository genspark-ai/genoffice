# @genoffice/ppt2pptx

Legacy PowerPoint 97-2003 import: parse `.ppt` (OLE2/CFB + MS-PPT records) and rebuild it as `.pptx`. Staged like `@genoffice/doc2docx` — stage 1 is the internal editable-text foundation (not user-visible yet); the CLI and the Slides open flow wire it up in the following stages.

## Stage 1 carries

- the **change directory** the way PowerPoint itself resolves it: CurrentUserAtom → UserEditAtom → PersistDirectoryAtom (`persistId:20|cPersist:12` entries, not an implicit id array) → older generations through `offsetLastEdit` — a sequential walk would read stale generations
- slide order from the DocumentContainer's SlideListWithText (instance 0 = slides; each SlidePersistAtom carries the persist id of its SlideContainer), the slide surface from the DocumentAtom (576 master units per inch → EMU)
- every slide's text sequences as **editable text boxes**: the slide's SlideListWithText outline entries (where real PowerPoint stores placeholder title/body text) plus the text boxes nested five levels down under PPDrawing/Escher (OfficeArtClientTextbox), with PowerPoint's runes normalized (`\u000B` → newline, the bullet-in-paragraph marker → space, cp1252 smart quotes named)
- the OLE2/CFB container opener is the **shared** one in `@genoffice/compound-file` (`openOle2Container`: 8-byte signature, 512-byte header, input/stream caps), with the PowerPoint-specific checks (presentation stream, encrypted headerToken 0xF3D1C4DF) layered on top
- an explicit error model (`LegacyPptError`, stable codes, including `encrypted`)
- a regression test on the in-repo real deck `packages/file-parse/tests/fixtures/legacy-sample.ppt` (plus the byte builder for the change-directory cases)

## Stage 3 carries

- **Escher geometry**: `OfficeArtChildAnchor` / `OfficeArtClientAnchor` rects
  (master units, 576/inch) become real `<a:xfrm>` offsets and extents, FSP
  shape types become DrawingML presets, FSP flags carry horizontal/vertical
  flips, and the shape's OPT table supplies fill/line colors and line width.
  A shape with no anchor still takes the stage-1 stacked fallback, so no text
  is lost.
- **StyleTextProp formatting**: `StyleTextPropAtom` (0x0FA1) paragraph and
  character run tables are decoded per the MS-PPT masks and mapped to
  `<a:pPr algn>` and `<a:rPr>` (bold, italic, underline, point size, RGB
  color). Outline-referenced text boxes receive their SlideListWithText text.
- **Pictures**: embedded PNG/JPEG blips from the global BStore's FBSE entries
  (or the `Pictures` stream when `foDelay` points there) are written as
  `ppt/media/imageN.*` with `<p:pic>`, slide relationships and content types.
  Unsupported kinds are omitted with a warning.
- **Masters**: the `MainMasterContainer`'s anchored shapes are reproduced in
  `slideMaster1.xml`, so every slide's layout inherits them.

The parse model (`PptPresentation`, `parsePptPresentation`) now exposes
`slides[].shapes`, `images` and `master` for callers; the third argument of
`parsePptPresentation` (the Pictures stream) is optional, so stage-1 callers
and the hostile-input fuzz are unaffected.

## Not carried yet

Shape rotation, metafile/EMF/WMF/TIFF/DIB/PICT blips, font-face resolution
from the FontCollection, bullets and numbering, tab stops, grouped shapes,
placeholder inheritance semantics and notes. Stage 3 keeps the geometry and
formatting it can prove from the bytes; the rest remains stacked text.
