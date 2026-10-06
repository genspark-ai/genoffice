# Docs: word processing

Docs is the Word-like processor: reads and writes genuine .docx with true WYSIWYG pagination.

## The ribbon

Tabs: **Home / Insert / Draw / Layout / Design / References / Review / View**, plus contextual tabs for the selected object — table design and layout for a table, picture format for a picture, shape format for a shape, header & footer while you are editing one.

- **Home**: clipboard; font (incl. CJK sizes and emphasis marks) with **Clear All Formatting** and a **Show/hide formatting marks** toggle; paragraph (align/indent/spacing/lists) plus **Define New Bullet / Number Format / Multilevel List**, which save your own list styles into the document; styles (Heading 1-6/Normal/Quote, modifiable) with a **Styles Pane** for the full list.
- **Insert**: page and section breaks; tables (a rows × cols grid, or **Insert Table…** for an exact size); pictures, shapes, text boxes; **Cover Page** and **Blank Page** from a preset gallery; **Chart**; **Drop Cap**; **WordArt**; fields (date, time, page, page count, file name); hyperlinks, **Bookmark** and cross-references; comments; header/footer and page numbers; symbols and equations.
  - **Chart** inserts a real chart object with its own data — bar, line or pie — not a picture. Word's _Edit Data_ opens the numbers behind it.
- **Draw**: ink over the page, in a **Drawing Tools** group — **Select** to go back to text editing, then **Pen**, **Highlighter** and **Eraser** (a click or a swipe removes the whole stroke). Beside them, **Pen Style** / **Highlighter Style** is one control holding colour swatches and a row of widths; its label follows whichever tool is active. Ink is stored in the document as an annotation floating above the text, so it survives save and reopen, and **Clear All** in the next group removes it all.
- **Layout**: margins, orientation and paper size, columns, paragraph indents and spacing.
- **Design**: themes, color sets, watermark, page borders.
- **References**: table of contents (updatable), footnotes/endnotes, captions, cross-references.

  ![The References tab](img/docs-references.png)

- **Review**: **Editor** proofreads the whole document for spelling, grammar and punctuation; **Translate**; spell check; comments (**AI Resolve Comments** works through the open ones); track changes with All/Simple markup views, accept/reject, and **AI Revision Summary**; word count; **Compare** against another file; **Protect Document**.
- **View**: five ways to look at the file — **Print Layout**, **Web Layout**, **Outline**, **Read Mode** and **Page Preview**; zoom out/in/100%/page width/one page; **AI Panel**; **Dark Mode**; ruler, gridlines and the navigation pane; **New Tab**, **Split** and **Switch Tabs**; the searchable **keyboard-shortcuts dialog**.
  - **Dark Mode** darkens the page and the canvas around it, never the ribbon — Word's split between a dark editing surface and a dark window. The choice is remembered, and it wins over the app theme either way.
  - **Split** opens a second pane below that scrolls independently and mirrors the first; close it with the × on its edge.

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
  _Show Heading 1_ leaves you a table of contents you can actually scan.
- **Promote / Demote** change the heading's level, and with it the level every
  heading below it inherits — the way a chapter becomes a section.
- **New Heading Before / After** insert one at the caret, without leaving the pane.
- **Delete** removes the heading _and everything under it_, which is the one to be
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
