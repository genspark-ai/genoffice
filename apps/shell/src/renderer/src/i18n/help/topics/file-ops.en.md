# File operations: rename, delete, export

This chapter covers the file operations shared by every editor; each editor's own export options are in its chapter.

The row's **⋯ menu** (hover a file row) gathers these actions:

![A file row's ⋯ menu](img/file-ops.png)

## Rename

Two entry points, one set of checks:

- Home's row **⋯ ▸ Rename**.
- **Double-click a file tab** to rename inline (see [Tabs and window management](help://tabs-and-windows)).

Rules: the extension is preserved automatically; illegal characters, trailing dots and reserved names (CON/NUL and friends) are rejected with a message; a same-folder name conflict is blocked too. The real file on disk is renamed, and recents and stars follow.

## Delete

- Home **⋯ ▸ Delete**: asks which files, then moves them to the **system trash**, where the OS can restore them. GenOffice keeps no undo of its own for this — recovery is the trash's job, not a toast's.

## Duplicate

**⋯ ▸ Duplicate** creates <name> copy in the same folder; collisions append a counter automatically. The copy is added to **Recent** rather than opened for you, so it is one click away instead of in front of you.

## Save and Save As

- **⌘S / ctrl+S** saves the current file; an untitled file asks for location and name first.
- **Save As** writes a new file and leaves the original untouched; subsequent edits target the new file.
- Every save is atomic (temp file + rename); a mid-write quit cannot corrupt the file.
- Autosave only kicks in after the first manual save (see [Quick start](help://getting-started)).

## Export to PDF

- **Docs**: File ▸ Export as PDF (or the ribbon button), paginated as laid out.
- **Slides**: export rasterizes page by page, with progress for large decks.
- **Sheets**: export follows print pagination.
- Exports render in a hidden window and land wherever you choose.

## Export to Word / images

- **PDF ▸ Convert to Word**: turns the PDF into a .docx (converted locally; complex layouts are best-effort).
- **Docs** can export pages as images (PNG per page).

## Print

File ▸ Print in each editor (⌘P/ctrl+P) opens the system print dialog; PDFs print with the current page order and rotations.

## Where untitled files live

The location you pick at the first save is its home; before that the document exists only in memory. Autosave takes over only after that first save.
