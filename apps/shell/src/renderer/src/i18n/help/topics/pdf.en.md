# PDF: reading, annotating and redacting

GenOffice's PDF editor is a full read-and-write workspace organized around five ribbon tabs: **Home / Annotate / Edit / Page / View**. Beyond fast reading and search, you can highlight and comment on text, edit existing paragraphs directly on the page, run OCR over scanned pages and edit the recognized text in place, permanently redact sensitive passages, sign with handwritten ink or cryptographic certificates, and fill interactive forms.

## Reading and navigation

- **Left navigation sidebar**: Toggle the left pane between **page thumbnails** (showing a live visual preview of every page with the currently visible viewport range highlighted, so you can click any thumbnail to jump straight there) and the document **outline** (its hierarchical bookmark tree).
- **Editable bookmarks**: The outline pane is no longer read-only navigation—click the edit toggle (`✎`) in the outline header to **add a bookmark** for the current page, **rename** any existing bookmark inline, or **delete** entries you no longer need. Clicking `✓` exits edit mode while keeping your updated tree queued as pending changes, and your next save writes the rebuilt `/Outlines` hierarchy back into the PDF (if you delete or reorder pages before saving, bookmarks pointing to removed pages are cleaned up and surviving page targets stay accurately aligned).
- **Zoom controls**: Step the zoom ratio up or down using the zoom control at the bottom right of the window, or hold `Ctrl` (`⌘` on macOS) and roll the mouse wheel to zoom smoothly around the page.
- **Page rotation**: Rotate either the current page or every page in the document from the **Page** tab; rotations mark the document as modified and are written permanently into the PDF when you save.
- **Full-text search**: Press `⌘F` / `Ctrl+F` to open the search bar and find text across the entire PDF, with every match highlighted on the canvas and toggles for case-sensitive and whole-word matching.
- **Password-protected PDFs**: Opening an encrypted PDF pops up a dedicated password dialog; the password you enter is held in memory for the current viewing session only and is never written to disk or logs.

## Text selection and markup (Annotate)

Try the text-markup flow on any paragraph in a searchable PDF:

1. **Click and drag your pointer across a sentence on the page**—as soon as you release the mouse button, a compact floating annotation bar appears right above your selection:

![The annotation bar after selecting text](img/pdf-highlight.png)

2. Click **Highlight** (click the colored square beside it to pick a different swatch from the palette), **Underline**, or **Strikethrough** to mark the passage immediately; or click **Ask AI** on the right end of the bar to send the selected passage together with your question straight to the AI assistant panel.
3. To remove a markup you just applied, simply drag-select the same passage again and click the active markup button on the floating bar (it works as a Word-style on/off toggle), or click the existing annotation on the page and press `Delete`.

A few practical details to keep in mind while marking up text:

- Whenever you drag over text in selection mode, the floating popup bar gives you instant access to **Highlight / Underline / Strikethrough / Copy / Ask AI** right at your cursor.
- Choosing a color from the swatch palette applies to new markups right away, and **applying the same markup type to an already-marked passage toggles it back off**, matching Microsoft Word's behavior.
- Markups that were already saved into the PDF—including annotations created earlier in Adobe Acrobat or Preview—can be clicked directly on the page and deleted via the `⋯` menu or the `Delete` key.
- **Automatic return to select mode**: While a drawing tool (such as rectangle or arrow) is armed, the underlying text layer is temporarily locked so you can draw cleanly without selecting letters by mistake; as soon as you finish placing one shape, the drawing tool automatically disarms itself and returns you to text-selection mode.

## Drawing tools (Annotate)

The **Annotate** tab provides six freehand and geometric markup tools: **Ink, Rectangle, Ellipse, Arrow, Note**, plus the dedicated **Redact area** tool for permanent data removal.

- **Single-use toggle behavior**: Clicking any drawing tool arms it for placement, and **the tool automatically disarms itself as soon as your shape lands on the page** so your next click can select or move items immediately (if you want to draw another shape of the same kind, click the tool button again; clicking an armed tool before drawing also disarms it).
- **Stroke width and palette colors**: Freehand ink follows your chosen stroke thickness, while rectangles, ellipses, and arrows are dragged out to size on the canvas using whichever color you pick from the drawing palette.
- **Selecting, moving, and resizing**: Any shape you place on a page remains a live editable object—click it to drag it to a new position, pull the corner handles on a rectangle or ellipse to resize it, or press `Delete` to remove it.
- **Redaction, the complete workflow** (for example, permanently stripping a confidential line of text or an account number before sharing a document):

  1. Open the **Annotate** tab on the ribbon and click **Redact area** to arm the redaction tool.
  2. **Drag a box over the text or image you want to erase**—the region is covered with a hatched preview overlay so you can still verify what is marked, and the toolbar reveals **Clear marks** and **Apply redaction** buttons:

  ![The page after marking a redaction](img/pdf-redact.png)

  3. Click **Apply redaction** and confirm in the prompt—GenOffice asks where to save a new redacted working copy and opens that copy in place. Inside the redacted file, the underlying characters and image pixels under your redaction boxes are **physically stripped from the PDF structure** (never merely hidden behind a black rectangle that someone could copy through or delete), while your original unredacted file on disk stays completely untouched.

  If you draw a redaction box in the wrong spot before applying, simply click **Clear marks** on the toolbar to wipe the pending hatched overlays and draw again.

## Sticky notes and comment threads

- **Dropping a sticky note**: Select the **Note** tool and click anywhere on the page to place a pin and open a margin comment card (with a configurable author name); confirming writes a standard PDF Text annotation that opens cleanly in Acrobat and other PDF readers.
- **Replying, editing, and deleting threads**: Click any pin on the page to expand its discussion card, where you can **reply** in a clean flat thread (matching the WPS and Acrobat comment style), **edit** your own remarks, or **delete** either an individual reply or the entire thread at once.
- **Preserving reply chains on save**: In-progress comment edits stay live in your session until you save, at which point GenOffice writes the updated text back into the exact same PDF annotation object so the reply hierarchy stays intact.

## Editing PDF content (Edit)

- **Edit existing text blocks**: On the **Edit** tab, click **Edit text** and click any paragraph on the page to edit its wording block by block in place (powered by the pdfium engine, which matches the surrounding font metrics and styling as closely as possible).
- **Insert new searchable text**: Choose **Insert text** and click anywhere on the page to place a fresh text block using the font family, point size, and color you select.
- **Insert pictures and stamps**: Place an image file from disk onto any page and drag its handles to position and scale it, or drop an approval stamp onto the document.
- **OCR text layer and in-place scan editing**: On a scanned PDF whose pages are flat bitmap images, click **Recognize text** to run local OCR across every scanned page in the document (with no artificial page cap) and embed a searchable text layer; once recognized, you can edit that OCR text directly in place while the original scanned page image stays underneath for reference, and use the **eye toggle** on the edit bar to switch between rendering your edited text visibly on the page surface or keeping it in the invisible search layer.
- **Interactive AcroForm fields**: Standard PDF form fields—text inputs, checkboxes, and choice fields—can be clicked and filled directly on the canvas, and every entered value is written back into the file when you save.

## Signatures

- **Hand-drawn ink signature**: Draw your signature smoothly with a mouse, trackpad, or stylus, then place it anywhere on the page or bind it directly to a PDF signature form field.
- **Image signature**: Import a transparent image of your handwritten signature or company chop and position it on the page.
- **Reusable signature library**: Once created, your ink and image signatures are saved locally so you can drop them onto future documents in a single click without redrawing.
- **Cryptographic certificate signing (Annotate tab)**: Click **Certificate sign** on the Annotate tab to sign the PDF using a personal certificate from your operating system's certificate store or a local `.p12` / `.pfx` file; you can optionally render a visible signature stamp on the page and optionally certify the document to restrict subsequent edits.
- **Incremental signing that preserves earlier signers**: Applying a certificate signature always writes out a **new signed file** via a standard PDF incremental update and opens the signed copy in place—that way, earlier digital signatures on a multi-party contract remain cryptographically valid. Remember to save any pending text or annotation edits first, since a digital signature seals the exact bytes saved on disk.
- **Signature verification banner**: Whenever you open a digitally signed PDF, a banner across the top of the document summarizes the signatures present; click **View signatures** to inspect each signature's cryptographic validity, byte-range coverage, and certificate chain locally (network CRL/OCSP revocation lookups are not performed).

## Page operations (Page)

- **Rotate, delete, and drag to reorder**: Drag page thumbnails up or down in the sidebar or on the **Page** tab to rearrange the document's page sequence, or delete unwanted pages with a confirmation prompt.
- **Import pages and insert blank pages**: Use **Import pages** to pull selected pages from another PDF straight into the current document, or click **Insert blank page** to add a clean sheet for notes and diagrams.
- **Replace and crop page ranges**: **Replace pages** swaps a contiguous range of pages in the current PDF with updated pages from another file, while **Crop pages** trims away excess margins on the active page or across the entire document at once.
- **Normalize page sizes and reverse page order**: **Page size** rescales every page in a mixed-size PDF to a single uniform paper dimension, and **Reverse order** flips the entire document from back to front (ideal for fixing a stack of pages scanned in reverse).
- **Extract pages**: Export your selected pages out into a brand-new standalone PDF file while leaving the current document intact.
- **Split a PDF**: Choose between two ways to split—either **split by page ranges** into multiple separate PDF files, or **slice each individual page into a grid** of smaller pages (handy for splitting two-up scanned book spreads or poster sheets).
- **Merge PDFs**: Combine documents in two ways—either **append whole PDFs** together in sequence, or **tile multiple pages onto a single sheet** (N-up). Before reading any bytes into memory, GenOffice totals the file sizes of every selected input and **refuses merges exceeding 1 GiB** with a clear explanation so an oversized batch can never exhaust system memory.
- **Saving page-level changes**: Page rotations, crops, deletions, and reordering are written into the PDF on your next save, while **Save As** lets you write the reorganized pages to a new file without altering the original.

## Export and print

- **Local export to Word, PowerPoint, and Excel**: Choose **File ▸ Export as Word… / PowerPoint… / Excel…** from the menu bar—or click **PDF Converter** on the ribbon for the same three actions—to convert your PDF locally on your machine without uploading anything to the cloud. Exporting to `.pptx` turns each PDF page into a slide, exporting to `.xlsx` places each page on its own worksheet, and every export prompts you for where to save the output file.
- **Print**: Press `⌘P` / `Ctrl+P` to print through the system print dialog using your current page order and rotations, with full support for custom page ranges.

## Saving

- **Standard Save and AutoSave**: Everyday highlights, sticky notes, text edits, OCR layers, form values, and page rearrangements write straight back into the file atomically when you save.
- **Redaction always writes a separate copy**: Because redaction permanently destroys the underlying text and image data, **Apply redaction** always writes out a separate working copy and leaves your original source file untouched on disk.
