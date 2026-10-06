# Docs: word processing

Docs is the Word-like processor: reads and writes genuine .docx with true WYSIWYG pagination.

## The ribbon

Tabs: **Home / Insert / Layout / Design / References / Review / View**, plus contextual tabs for the selected object (table design, pictures).

- **Home**: clipboard; font (incl. CJK sizes and emphasis marks); paragraph (align/indent/spacing/lists); styles (Heading 1-6/Normal/Quote, modifiable).
- **Insert**: page/section breaks, tables (incl. quick tables), pictures, shapes, hyperlinks, header/footer, page numbers, date, text boxes.
- **Layout**: margins, orientation and paper size, columns, paragraph indents and spacing.
- **Design**: themes, color sets, watermark, page borders.
- **References**: table of contents (updatable), footnotes/endnotes, captions, cross-references.

  ![The References tab](img/docs-references.png)

- **Review**: spell check, comments, track changes (All/Simple markup views), word count.
- **View**: ruler, gridlines, navigation pane, zoom, and the searchable **keyboard-shortcuts dialog**.

## The navigation pane

**View ▸ Navigation Pane** opens a side pane that holds the document's heading
outline, a find box over the whole document, and a thumbnail per page. Whether it
is open is remembered between launches, so a document you navigate by outline
stays navigable.

**The outline** is the heading tree. Right-click a heading in it to fold and restructure rather than just navigate:

- **Collapse / Expand** on a heading folds that heading's whole subtree — the
  chapter disappears, its text stays in the document.
- **Collapse All / Expand All** folds or unfolds the lot at once. On a long report
  this is the difference between a readable outline and a wall of text.
- **Show Heading Levels** filters the tree to the depths you care about, so
  *Show Heading 1* leaves you a table of contents you can actually scan.
- **Promote / Demote** change the heading's level, and with it the level every
  heading below it inherits — the way a chapter becomes a section.
- **New Heading Before / After** insert one at the caret, without leaving the pane.
- **Delete** removes the heading *and everything under it*, which is the one to be
  careful with: it is a subtree delete, not a line delete.
- **Select Heading and Content** selects from the heading to the end of its
  subtree, ready for a whole-section edit.

## Right-click menu

Right-click anywhere in the body — the menu follows what you clicked. Main groups:

- **Clipboard**: cut / copy / paste / **paste as plain text**.
- **Font, paragraph**: change family and size, bold/italic/underline, alignment/indent/spacing without visiting the ribbon.
- **Synonyms**: lists synonyms for the selected word; click one to replace.
- **Translate** (AI): translate the selection into the target language (English, Simplified Chinese, Japanese, Korean, French, German, Spanish, ...) via the AI panel.
- **New comment**: attach a comment to the selection.
- **Spelling** (on a misspelled word): suggested replacements, ignore all, add to dictionary, set proofing language.
- **Hyperlink**: open / edit / copy link / remove hyperlink.
- **Picture**: view image, save image as, **wrap text** (inline / square left & right / top and bottom / behind text / in front of text), arrangement order.
- **Fields** (TOC, page numbers): update field / toggle field codes / edit field.
- **List numbering** (inside a list): restart numbering / continue numbering / change list level / set numbering value.
- **Table** (cursor inside a table): insert rows/columns, merge / split cells, split table, autofit, cell alignment, distribute rows/columns, table properties, the delete menu, select.

## Editing

- Find & replace (ctrl+F / ctrl+H): case-sensitive, whole-word, regex.
- Format painter; deep undo/redo; paste options.
- Tables: merge/split cells, row/column ops, borders and shading, sort, formulas.
- Pictures: text wrapping, crop, compress; drawing canvas.

## CJK typography

- Punctuation compression and kinsoku line-breaking match Word; full/half-width conversion.
- Font candidates cover the common Windows and macOS CJK family names.

## AI

- Ribbon AI button and the side panel: rewrite, expand, translate, summarize, insert-table presets plus free-form instructions.
- Every AI turn snapshots first; roll back from the version list, and the rollback itself is undoable.

## Saving and export

- Saves .docx rewriting only changed paragraphs — untouched content stays byte-identical.
- Exports PDF (as paginated) and per-page images.

## Print

ctrl+P through the system dialog, WYSIWYG pages.
