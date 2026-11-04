# @genoffice/ppt2pptx

Legacy PowerPoint 97-2003 import: parse `.ppt` (OLE2/CFB + MS-PPT records) and rebuild it as `.pptx`. Staged like `@genoffice/doc2docx` — stage 1 is the internal editable-text foundation (not user-visible yet); the CLI and the Slides open flow wire it up in the following stages.

## Stage 1 carries

- the **change directory** the way PowerPoint itself resolves it: CurrentUserAtom → UserEditAtom → PersistDirectoryAtom (`persistId:20|cPersist:12` entries, not an implicit id array) → older generations through `offsetLastEdit` — a sequential walk would read stale generations
- slide order from the DocumentContainer's SlideListWithText (instance 0 = slides; each SlidePersistAtom carries the persist id of its SlideContainer), the slide surface from the DocumentAtom (576 master units per inch → EMU)
- every slide's text sequences as **editable text boxes**: the slide's SlideListWithText outline entries (where real PowerPoint stores placeholder title/body text) plus the text boxes nested five levels down under PPDrawing/Escher (OfficeArtClientTextbox), with PowerPoint's runes normalized (`\u000B` → newline, the bullet-in-paragraph marker → space, cp1252 smart quotes named)
- the OLE2/CFB container opener is the **shared** one in `@genoffice/doc2docx` (`openOle2Container`: 8-byte signature, 512-byte header, input/stream caps), with the PowerPoint-specific checks (presentation stream, encrypted headerToken 0xF3D1C4DF) layered on top
- an explicit error model (`LegacyPptError`, stable codes, including `encrypted`)
- a regression test on the in-repo real deck `packages/file-parse/tests/fixtures/legacy-sample.ppt` (plus the byte builder for the change-directory cases)

## Not carried yet

shape geometry (the Escher anchors), text run formatting (StyleTextProp), images, masters, notes. Stage 1 stacks each slide's text boxes down the surface — editable and complete in words, not in layout.
