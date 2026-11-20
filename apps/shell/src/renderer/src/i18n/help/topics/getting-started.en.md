# Quick start: the interface and the basics

GenOffice is a desktop office suite that runs entirely on your own machine, bringing six full editors—Docs for word processing, Sheets for spreadsheets, Slides for presentations, PDF, Markdown, and HTML—together inside one window and one shared row of tabs. Every editor reads and writes genuine .docx, .xlsx, .pptx, and .pdf files that round-trip cleanly with Microsoft Word, Excel, and PowerPoint, with no internet connection required for your everyday document work.

## Interface overview

![The Home screen](img/home-screen.png)

When you launch GenOffice, the workspace is organized around three main areas plus a window-wide interface scale control:

- **Tab strip (top)**: Every file you open gets its own tab along the top of the window so you can switch, reorder, or tear off documents just like browser tabs. The leftmost **Home** tab is permanently pinned and cannot be closed, giving you an instant way back to your file lists, while any saved document tab can be double-clicked (or right-clicked via **Rename**) to rename its underlying file on disk inline.
- **Content area**: The main workspace below the tab strip automatically switches to show the full editor (or the Home screen) belonging to whichever tab is currently active.
- **Menu bar**: Located in the system menu bar at the top of the screen on macOS and along the top edge of the window on Windows and Linux, the File, Edit, and View menus dynamically swap their commands to match the active editor tab.
- **Interface scale**: When working on a high-DPI monitor, a large external display, or a compact laptop screen, you can scale the entire application interface—tabs, ribbons, sidebars, and dialogs together—between **80% and 150%** using **View ▸ Zoom In / Zoom Out / Reset** (`⌘⌃=` / `⌘⌃-` / `⌘⌃0` on macOS; `Ctrl+Alt+=` / `Ctrl+Alt+-` / `Ctrl+Alt+0` on Windows and Linux), while native window controls (macOS traffic lights and Windows/Linux caption buttons) automatically keep their reserved clearance so tabs never slide underneath them.

## Creating a document

Whenever you want to start a fresh document, you can pick whichever of these three entry points is closest at hand:

- Click one of the quick-create cards in the [Quick start](help://getting-started) strip at the top of **Home** (**AI Docs**, **AI Sheets**, **AI Slides**, **AI Markdown**, **AI HTML**, or **AI PDF**) to open a new file of that type straight in its editor.
- Open **File ▸ New** from the menu bar at any time and choose **Docs** (`⌘N` / `Ctrl+N`), **Sheets**, **Slides**, **Markdown**, **HTML**, or **PDF**.
- Drag an existing file from Finder or File Explorer directly into the GenOffice window, or simply double-click the file on your desktop once GenOffice is registered as your default handler.

Aside from blank PDFs (which are created on disk immediately so the PDF engine can operate on a real file), newly created documents open as untitled in-memory drafts first and only write a file to disk the first time you save them.

## Opening files

- Choose **File ▸ Open** (`⌘O` / `Ctrl+O`) from the menu bar to open your operating system's file picker, with native support for `.docx`, `.doc`, `.xlsx`, `.xls`, `.csv`, `.tsv`, `.pptx`, `.ppt`, `.pdf`, `.md`, `.html`, `.txt`, and `.json` files (multi-select opens each chosen file in its own tab).
- Click any row in the **Recent**, **Starred**, or **Folders** views on the Home screen to open that file immediately, or switch to its existing tab if it is already open.
- Run `genoffice <file>` from a terminal or shell script to hand a file directly to the running GenOffice window from the command line.

## How saving and AutoSave work

- **Manual Save and Save As**: Press `⌘S` / `Ctrl+S` or choose **File ▸ Save / Save As** at any time to write your changes to disk; on a brand-new untitled document, the first save prompts you to pick a destination folder and file name.
- **Safe AutoSave activation**: AutoSave turns on for a document only after you have saved that file manually at least once—that way, a PDF or Word document you merely open to read is never silently rewritten or given a new modification timestamp behind your back. Once active, AutoSave writes your latest edits to disk shortly after you pause typing.
- **Unsaved-change protection**: If you try to close a tab or quit the window while a document still has unsaved edits, GenOffice always stops and asks whether to **Save**, **Discard**, or **Cancel** before closing.
- **Atomic disk writes**: Every save writes to a temporary file alongside your document first and swaps it into place with an atomic rename, ensuring that an unexpected power loss or crash mid-save can never leave you with a corrupted, half-written file.

## Common shortcuts

| Action           | macOS | Windows / Linux |
| ---------------- | ----- | --------------- |
| New document     | ⌘N    | ctrl+N          |
| Open             | ⌘O    | ctrl+O          |
| Save             | ⌘S    | ctrl+S          |
| Close tab        | ⌘W    | ctrl+W          |
| Open this manual | F1    | F1              |
| Fold the ribbon  | ⌥⌘R   | Ctrl+F1         |

**Folding and compacting the ribbon** works across every ribbon-based editor in the suite. Pressing `⌥⌘R` (`Ctrl+F1` on Windows and Linux) or double-clicking the active ribbon tab title hides the command band below while keeping the top row of tab labels visible, freeing up vertical room for your document; clicking any ribbon tab brings the command band right back, and each editor remembers how you left it between sessions. For times when you want the buttons visible without taking up full height, Docs also offers a middle **compact** density (`⌥⌘K` on macOS, `Ctrl+Alt+K` on Windows and Linux) that strips the text labels and leaves a sleek single row of icons.

Editor-specific shortcuts (format painter, find and replace, table operations, slide show controls, and more) are covered in each editor's own chapter and in the full Keyboard Shortcuts reference; inside Docs, you can also press **⌘/** (`Ctrl+/`) at any time to open a live searchable keyboard-shortcuts dialog.

## The Option+Command chords

Desktop Word reserves the `Option+Command` modifier layer (`Ctrl+Alt` on Windows and Linux) for structural jumps and document styles, and GenOffice maps that layer the exact same way so your existing muscle memory carries over. Docs uses most of these chords, Sheets claims two of them for Excel border parity, and one chord works globally across every editor.

**Docs**

| Chord (macOS)   | What it does       | Windows / Linux    |
| --------------- | ------------------ | ------------------ |
| ⌥⌘1 / ⌥⌘2 / ⌥⌘3 | Heading 1 / 2 / 3  | Ctrl+Alt+1 / 2 / 3 |
| ⌥⌘0             | Normal (body text) | Ctrl+Alt+0         |
| ⌥⌘M             | Paragraph dialog   | Ctrl+Alt+M         |
| ⌥⌘A             | New comment        | Ctrl+Alt+A         |
| ⌥⌘F             | Insert footnote    | Ctrl+Alt+F         |
| ⌥⌘E             | Insert endnote     | Ctrl+Alt+D         |
| ⌥⌘G             | Go To              | Ctrl+G             |

Two of these chords differ slightly between macOS and Windows for the exact same reason Microsoft Word splits them. Because **macOS reserves `⌥⌘D` system-wide** to show and hide the Dock, inserting an endnote uses `⌥⌘E` on a Mac while keeping the familiar `Ctrl+Alt+D` on Windows and Linux; similarly, **Go To** uses `⌥⌘G` on macOS (leaving `⌘G` and `⇧⌘G` for Find Next and Word Count) and simplifies to `Ctrl+G` on Windows.

**Sheets** (while the spreadsheet grid has focus)

| Chord (macOS) | What it does    | Windows / Linux |
| ------------- | --------------- | --------------- |
| ⌥⌘0           | Outside borders | Ctrl+Shift+7    |
| ⌥⌘−           | No border       | Ctrl+Shift+−    |

Here too, the platform difference reflects genuine Excel conventions rather than an arbitrary remap: Excel for Mac supports **both** `⌘⇧7` and `⌥⌘0` for applying outside borders, whereas Excel on Windows has always kept outside borders and border removal on `Ctrl+Shift+7` and `Ctrl+Shift+-`.

Notice that **`⌥⌘0` applies the Normal body style in Docs and applies Outside Borders in Sheets**. Because the two editors live in separate tabs and never share focus at the same moment, the two bindings never collide in practice—though it does mean `⌥⌘0` is spoken for and cannot be assigned as a suite-wide global shortcut.

**Every editor**: As noted earlier, **`⌥⌘R` / `Ctrl+F1`** folds or unfolds the ribbon toolbar from anywhere in the suite.

Keeping that layout consistent across the suite leaves the rest of the modifier layer clean and predictable on both macOS and Windows.

## Where to go next

- Organizing your recent files, starred items, local folders, and knowledge-base Q&A: see [The Home screen](help://home-screen).
- Working with many open documents, inline tab renaming, and tear-off windows: see [Tabs and window management](help://tabs-and-windows).
- Letting the built-in AI assistant draft, rewrite, analyze, and format your documents: see [The AI assistant panel](help://ai-panel).
- Changing the UI language, light/dark theme, default file associations, and coding-agent connections: see [Settings, language, theme and MCP integrations](help://settings-integrations).
