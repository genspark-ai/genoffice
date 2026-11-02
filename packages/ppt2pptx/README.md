# @genoffice/ppt2pptx

Legacy PowerPoint 97-2003 import: parse `.ppt` (OLE2/CFB + MS-PPT records) and rebuild it as `.pptx`. Staged like `@genoffice/doc2docx` — stage 1 is the internal editable-text foundation (not user-visible yet); the CLI and the Slides open flow wire it up in the following stages.

## Stage 1 carries

- the **change directory** the way PowerPoint itself resolves it: CurrentUserAtom → UserEditAtom → PersistPtr blocks with implicit ids (older blocks first, newer blocks override) — a sequential walk would read stale generations
- slide order from the DocumentContainer's SlideListWithText, the slide surface from the DocumentAtom (master units → EMU)
- every slide's text sequences (TextHeaderAtom kinds + TextCharsAtom UTF-16 / TextBytesAtom codepage) as **editable text boxes**, with PowerPoint's runes normalized (`\u000B` → newline, the bullet-in-paragraph marker → space, cp1252 smart quotes named)
- an explicit error model (`LegacyPptError`, stable codes)

## Not carried yet

shape geometry (the Escher anchors), text run formatting (StyleTextProp), images, masters, notes. Stage 1 stacks each slide's text boxes down the surface — editable and complete in words, not in layout.
