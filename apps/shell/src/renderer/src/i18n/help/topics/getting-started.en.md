# Quick start: the interface and the basics

GenOffice is an office suite that runs entirely on your machine: one window, one row of tabs, holding six editors — Docs (word processing), Sheets (spreadsheets), Slides (presentations), PDF, Markdown and HTML. Files are genuine .docx / .xlsx / .pptx / .pdf, fully interchangeable with Word, Excel and PowerPoint. No network required.

## Interface overview

![The Home screen](img/home-screen.en.png)

The window has three parts:

- **Tab strip (top)**: every open file is a tab. The leftmost Home tab is always present and cannot be closed; the other tabs are your documents. Double-click a tab to rename its file inline.
- **Content area**: the editor (or Home) belonging to the active tab.
- **Menu bar**: in the system menu bar on macOS, at the window top on Windows/Linux. File/Edit/View menus switch to match the active editor.

## Creating a document

Any of:

- Click a quick-create card in the [Quick start](help://getting-started) section of **Home** (AI Docs, AI Sheets, AI Slides, ...).
- Menu **File ▸ New**: Docs (⌘N/ctrl+N), Sheets, Slides, Markdown, HTML or PDF.
- Drag a file onto the window, or double-click it in your file manager (if GenOffice is the default app).

A new document opens untitled; the file on disk is only created at the first save.

## Opening files

- Menu **File ▸ Open** (⌘O/ctrl+O) opens the system picker: .docx / .doc / .xlsx / .xls / .csv / .tsv / .pptx / .ppt / .pdf / .md / .html.
- Click anything in Home's **Recent** list.
- `genoffice <file>` from a terminal also opens files.

## The save model

- **Manual save**: ⌘S/ctrl+S, or File ▸ Save / Save As. The first save asks for location and name.
- **Autosave** turns on only after you have saved the file manually at least once — a PDF you only read is never silently rewritten. Autosave fires shortly after content changes.
- Closing a tab with unsaved changes asks Save / Discard / Cancel first.
- Every write lands atomically (temp file + rename), so a power cut cannot leave half a file.

## Common shortcuts

| Action           | macOS | Windows / Linux |
| ---------------- | ----- | --------------- |
| New document     | ⌘N    | ctrl+N          |
| Open             | ⌘O    | ctrl+O          |
| Save             | ⌘S    | ctrl+S          |
| Close tab        | ⌘W    | ctrl+W          |
| Open this manual | F1    | F1              |

Shortcuts inside each editor (format painter, find & replace, table ops, ...) are in their chapters; Docs additionally ships a searchable keyboard-shortcuts dialog (see its chapter).

## Where to go next

- Where files live: [The Home screen](help://home-screen).
- Managing many open files: [Tabs and window management](help://tabs-and-windows).
- Having the AI do the work: [The AI assistant panel](help://ai-panel).
- Language, theme, default apps: [Settings, language, theme and MCP integrations](help://settings-integrations).
