# The Markdown editor

The Markdown editor opens .md / .markdown with a source + rendered-preview experience.

- **Open**: from Home or File ▸ Open; the CLI works too.
- **Edit**: plain-text editing; GFM extensions (tables, task lists, strikethrough, autolinks) render in the preview.
- **Preview**: live; relative assets like images resolve next to the document.
- **Saving**: byte-faithful — BOM, CRLF and the presence of a trailing newline are preserved; an unmodified save does not rewrite the file.
- **Find & replace**: ctrl+F searches the source; Replace All writes back.
- **AI**: preset buttons let the assistant rewrite, extend or translate the document.

## The toolbar

One row of buttons above the editor (hover for tooltips):

![The Markdown toolbar](img/md-toolbar.png)

- **File & history**: Save, Save As, Undo, Redo, Find; the **AutoSave** toggle on the right writes changes to disk on a timer.
- **AI button**: opens the AI panel; next to it are the rewrite / extend / translate presets.
- **Paragraph style** (dropdown): switch between body text and heading levels.
- **Inline formatting**: **bold**, _italic_, ~~strikethrough~~, `inline code`, link.
- **Lists**: bullet, numbered, task list.
- **Insert**: table, image, horizontal rule.
- **Properties**: insert or jump to the YAML front matter block at the top of the file.
- **Outline**: jump by heading hierarchy.
- **Spellcheck**: toggle spellcheck for this document.

Three quick examples:

- **Heading**: put the cursor on the line ▸ paragraph-style dropdown ▸ "Heading 1".
- **Table**: click **Insert table** ▸ drag the row/column count ▸ type into the cells; the preview renders it immediately.
- **Task list**: select a few lines ▸ click **Task list** ▸ each line becomes `- [ ]`, rendered as checkboxes in the preview.

## Exporting

File menu, all local and all asking where to put the result:

- **Export as Word…** and **Export as PDF…** write a real .docx or .pdf.
- **Export as Images…** writes a PNG per content element into a directory you pick.
- **Convert and Open in Docs** converts to .docx and opens it in the built-in Docs tab here in the app — it is not a hand-off to anything in the cloud, and the converted copy lives in a cache folder that is cleaned up after about a week.

## Source view

The ribbon carries a **Source** toggle (localized with the app). Turn it on and the editor is replaced by the raw Markdown: exactly the text a save writes, nothing prettified, nothing normalised underneath you.

- **Editing is byte-faithful.** A save from source view produces the same bytes a save from the editor does — BOM, CRLF and the presence of a trailing newline all survive.
- **It is the same document.** Toggle back and forth freely; the source is the editor's own text, not a copy that has to be merged.
- **The formatting toolbar is unavailable** while it is open, because most of those buttons insert editor constructs that only mean something in the rendered side. It comes back when you close the view.
- **JSON and other source-mode files** open here directly: there is nothing to render, so the source _is_ the document.
