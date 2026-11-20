# Docs: word processing

GenOffice Docs is a Word-compatible word processor that reads and writes genuine `.docx` files with true WYSIWYG page layout—every line break, multi-page table split, header, footer, and footnote you see on screen matches what lands on paper or in an exported PDF.

## The ribbon

The ribbon across the top of Docs is organized into eight main tabs—**Home / Insert / Draw / Layout / Design / References / Review / View**—plus contextual tabs that appear automatically whenever you select a specific object (Table Design and Layout when the cursor is inside a table, Picture Format for an image, Shape Format for a drawing shape, and Header & Footer while editing a page header or footer). In addition to folding the ribbon completely, Docs offers a middle **compact** density (`⌥⌘K` on macOS, `Ctrl+Alt+K` on Windows and Linux) that keeps every icon on screen while hiding the text labels to save vertical space, remembering your preference across sessions.

- **Home**: Clipboard and format painter; font controls (including traditional CJK point sizes, emphasis marks, superscript/subscript, and text shading) paired with **Clear All Formatting** and a **Show/hide formatting marks** toggle; paragraph formatting (alignment, indents, line and paragraph spacing, bulleted and numbered lists) plus **Define New Bullet / Number Format / Multilevel List** for saving custom list numbering styles straight into the document; and the styles gallery (Heading 1–6, Normal, Quote, all customizable) alongside a **Styles Pane** listing every style in the file. Any style can specify separate **East Asian and Latin font families** (for instance, pairing a Chinese or Japanese typeface with Times New Roman for English words and digits), and that dual-font binding survives the `.docx` round-trip cleanly with Microsoft Word.
- **Insert**: Page breaks and section breaks; tables (drag out a rows × columns grid or open **Insert Table…** for exact dimensions and autofit rules); pictures, geometric shapes, and text boxes; **Cover Page** and **Blank Page** templates from a built-in gallery; native **Chart** objects; **Drop Cap** and **WordArt**; dynamic fields (date, time, page number, total page count, file name); hyperlinks, **Bookmark** anchors, and cross-references; margin comments; per-section headers, footers, and page numbers; plus special symbols and mathematical equations.
  - **Chart** inserts a genuine editable chart object backed by its own embedded data table—bar, line, or pie—rather than a flat bitmap image, so clicking _Edit Data_ in either GenOffice or Microsoft Word reopens the underlying numbers at any time.
- **Draw**: Freehand inking directly over the page inside a dedicated **Drawing Tools** group—click **Select** to return to regular text editing, or arm **Pen**, **Highlighter**, or **Eraser** (clicking or swiping across any stroke erases the entire stroke cleanly). Right beside them, **Pen Style / Highlighter Style** combines color swatches and stroke-width presets into a single control whose label automatically tracks whichever ink tool is currently active. Every ink stroke is stored inside the `.docx` file as a floating annotation above the text so it survives saving and reopening, and clicking **Clear All** in the next group wipes every stroke from the page at once.
- **Layout**: Configure page margins, portrait or landscape orientation, standard paper sizes, single- or multi-column text flow, and exact paragraph indentation and before/after spacing.
- **Design**: Switch document-wide themes, coordinated color palettes, and font pairs, or add diagonal/picture **watermarks** and full-page **page borders**.
- **References**: Build an automatic, clickable **Table of Contents** that refreshes its page numbers on demand, insert numbered **footnotes and endnotes**, attach numbered **captions** to figures and tables, and insert live **cross-references** pointing to headings, bookmarks, footnotes, or captions.

  ![The References tab](img/docs-references.png)

- **Review**: Run **Editor** to proofread the entire document for spelling, grammar, and punctuation; translate passages with AI or run standalone spell check; collaborate through margin comments (where **AI Resolve Comments** walks through open review comments and applies the requested edits for you); use **Track Changes** with **All Markup** or **Simple Markup** views, individual or batch accept/reject controls, and **AI Revision Summary** for a plain-language digest of what changed; check live word and character counts; **Compare** two versions of a document side by side; or lock editing with **Protect Document**.
- **View**: Switch between five distinct ways of looking at your document—**Print Layout**, **Web Layout**, **Outline**, **Read Mode**, and **Page Preview**; adjust zoom (zoom out/in, 100%, page width, or whole page); toggle the side **AI Panel**, **Dark Mode**, ruler, gridlines, and the left navigation pane; open a **New Tab**, **Split** the window horizontally, or **Switch Tabs**; and open the live searchable **keyboard-shortcuts dialog** (`⌘/` or `Ctrl+/`).
  - **Dark Mode** darkens the editing page surface and the surrounding workspace canvas while leaving the top ribbon in its familiar chrome—matching Word's separation between a dark page canvas and the window theme, and never altering the underlying authored colors printed on paper or exported to PDF. Your choice is remembered across launches and takes precedence over the suite theme.
  - **Split** opens a second independently scrolling pane directly below the first so you can view or edit two distant parts of the same document simultaneously; click the `×` on the divider edge whenever you are ready to close the lower pane.
- **Inferred outline for unstyled documents**: Many legal contracts, memos, and imported files mark section breaks using short all-bold lines, all-caps lines, or numbered lines (such as `I. BACKGROUND` or `1. Scope of Work`) written on ordinary Normal paragraphs without formal Heading styles. When a document contains no heading styles of its own, Docs automatically detects those structural heading lines—without modifying a single byte of the file—and lists them with a subtle inferred indicator in the **Outline view** and the **Story view**, so you can still jump section by section; applying a real Heading style to any inferred line immediately promotes it into a formal document heading.

## The navigation pane

Choosing **View ▸ Navigation Pane** opens a persistent left sidebar with three tabs: the document's hierarchical **heading outline**, a whole-document **Find** box with live result excerpts, and a visual **page thumbnail** strip. Docs remembers whether you left the navigation pane open between launches, so if you rely on an outline tree to move around long reports, it is already waiting for you next time.

The **Outline** tab is far more than a read-only jump list—right-clicking any heading in the tree lets you fold chapters or reorganize the structure of the document on the spot:

- **Collapse / Expand**: Folds or unfolds that heading's entire subtree in the outline pane so you can tuck finished chapters out of sight while keeping their text safely in the document.
- **Collapse All / Expand All**: Folds or expands the entire heading tree in one click—on a fifty-page report, this turns a dense wall of subsections back into a clean top-level table of contents you can scan at a glance.
- **Show Heading Levels**: Filters the outline tree down to the exact heading depth you want to see (for example, choosing _Show Heading 1_ or _Show Heading 2_ to focus purely on major chapters and sections).
- **Promote / Demote**: Moves the clicked heading up or down one heading level and automatically shifts every child heading underneath it in lockstep—turning a section into a top-level chapter (or vice versa) in a single click.
- **New Heading Before / After**: Inserts a new sibling heading at the right spot in the document and places the caret there without making you scroll the canvas first.
- **Delete**: Removes the clicked heading **along with its entire body text and all child subsections underneath it**—a true subtree delete designed for cutting a whole section cleanly (and fully undoable with `⌘Z` / `Ctrl+Z` if you change your mind).
- **Select Heading and Content**: Highlights the entire span of the document from that heading down to the end of its last child subsection, ready for you to copy, move, or hand to the AI assistant for a section-wide rewrite.

## Right-click menu

Right-clicking anywhere in the document body opens a context-sensitive menu tailored to whatever is under your pointer:

- **Clipboard**: Cut, copy, paste with formatting, or **Paste as Plain Text** to strip away external web fonts and background colors.
- **Font and paragraph formatting**: Change font family, point size, bold/italic/underline, text alignment, indentation, and line spacing right at your cursor without reaching up to the ribbon.
- **Synonyms**: Look up context-appropriate synonyms for the selected word and click any candidate in the submenu to replace it in place.
- **Translate** (AI): Translate the selected passage into your target language (English, Simplified Chinese, Japanese, Korean, French, German, Spanish, and more) directly through the AI assistant panel.
- **New comment**: Attach a new margin review comment anchored to the current text selection.
- **Spelling** (when right-clicking a word flagged with a red spelling underline): Pick from suggested corrections at the top of the menu, **Ignore All** occurrences in the document, **Add to Dictionary**, or change the active **Proofing Language**.
- **Hyperlink** (when clicking a link): **Open Hyperlink** in your browser, **Edit Hyperlink** text or URL, **Copy Link**, or **Remove Hyperlink**.
- **Picture** (when right-clicking an image): Open a full-size preview with **View Image**, export the embedded picture via **Save Image As…**, switch **Wrap Text** modes (In Line with Text, Square Left/Right, Top and Bottom, Behind Text, or In Front of Text), and adjust z-order layering (Bring to Front, Send to Back, Move Forward/Backward).
- **Fields** (inside a table of contents, page number, or date field): **Update Field** to refresh its latest value, **Toggle Field Codes** to inspect the raw field syntax, or open **Edit Field…**.
- **List numbering** (when the caret is inside a numbered list): **Restart Numbering** back at 1, **Continue Numbering** from a previous list, **Change List Level**, or open **Set Numbering Value…** to start from any custom number.
- **Table** (when the caret is inside a table): Insert rows above/below or columns left/right, merge or split cells, **Split Table** in two at the active row, adjust **AutoFit** behavior, set 9-way cell alignment, distribute rows and columns evenly, open **Table Properties…**, or use the **Delete** and **Select** submenus.

## Editing

- **Find and replace (`⌘F` / `Ctrl+F` and `⌃H` / `Ctrl+H`)**: Search and batch-replace across the document with toggles for case sensitivity, whole-word matching, and full regular expressions.
- **Format painter, deep undo/redo, and paste options**: Copy styling with single-click or double-click format painting, step through an unlimited undo/redo history, and control how pasted rich text merges with surrounding styles.
- **Tables that paginate like Word**: Merge and split cells, add or remove rows and columns, customize borders and cell shading, sort table rows, and run in-table formulas. Long table rows split cleanly across page boundaries when needed, and pressing `Ctrl+Enter` (`⌘⏎` on macOS) inside a table cell forces an immediate page break right at that line—honored identically on screen, in print, and in exported PDFs.
- **Pictures and drawing canvas**: Crop image borders non-destructively, compress pictures to reduce file size, wrap text around floating graphics, and group complex diagrams inside a drawing canvas.

## CJK typography

- **Word-aligned punctuation compression and kinsoku line breaking**: Closing punctuation marks never wrap onto the start of a new line alone, opening quotes and brackets never dangle at the end of a line, consecutive Chinese/Japanese punctuation marks compress proportionally, and built-in tools handle full-width/half-width and Simplified/Traditional Chinese conversion.
- **Cross-platform CJK font resolution**: Built-in font metrics and fallback candidates recognize localized and English family names for standard Windows and macOS Chinese, Japanese, and Korean fonts alike, keeping line heights and page breaks consistent across operating systems.

## AI assistant in Docs

- **One-click presets and free-form instructions**: Click the **AI** button on the ribbon or open the side AI panel to use presets like **Rewrite**, **Expand**, **Translate**, **Summarize**, and **Insert Table**, or type any custom drafting and formatting instruction in plain language.
- **Automatic per-turn snapshots and undoable rollback**: Before the AI assistant applies any edit to your document, Docs automatically captures a snapshot in the panel's **version list**; you can roll back to any earlier turn in one click, the snapshots persist even after you close and reopen the file, and a rollback itself can be undone with `⌘Z` / `Ctrl+Z`.

## Saving and export

- **Byte-preserving incremental `.docx` saves**: When saving an edited `.docx` file, Docs rewrites **only the specific paragraphs you actually modified**, leaving every untouched paragraph, complex field, and XML structure byte-for-byte identical to the original file so external formatting is never disturbed.
- **WYSIWYG PDF and per-page image export**: Export your document straight to a paginated `.pdf` file or render each page out as a high-resolution `.png` image locally.

## Print

Press `⌘P` / `Ctrl+P` to open the built-in WYSIWYG paginated print preview, verify your paper size and page range, and send the job to your system printer.

- **Print preview opens on the caret's page**: When you are editing deep inside a long document and press `⌘P` / `Ctrl+P` (or choose **File ▸ Print**), the print preview dialog opens directly on the exact page where your text cursor is currently sitting rather than forcing you to scroll down from page 1 every time.
