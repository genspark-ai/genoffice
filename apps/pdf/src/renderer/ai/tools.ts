import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { AgentToolCall, AgentToolDef, ToolExecution } from '@genoffice/agent-core'
import type { OutlineNode } from '../OutlinePanel'
import type { PageEntry, SearchIndex } from '../search'
import { searchInIndex } from '../search'
import { geomDispSize, pdfRectToCss, pdfToView, quadToRect, viewToPdf } from '../annotations'
import type { PageGeom } from '../annotations'
import { EDIT_FONTS } from '../../shared/ipc'
import type {
  CreateDocumentRequest,
  CreateDocumentResult,
  CreateDocumentType,
  FormValueInput,
  ImageLayer,
  ImageSearchResponse,
  MarkupType,
  MetadataInput,
  PageImageRef,
  TextEditInput,
  TextInsertInput,
} from '../../shared/ipc'
import { groupPageBlocks } from '../text-block'
import type { TextBlock } from '../text-block'
import { joinBlockLines, measurePt, wrapText } from '../text-wrap'
import { isScannedText } from '../ocr-layer'
import { t } from '../i18n/locale'
import { buildFormCatalog } from '../form-catalog'
import { flattenThread, type NoteThreadItem } from '../note-threads'
import type { StampConfig } from '../edit-state'
import type { LocalTextInsert } from '../text-edit-preview'
import { DEFAULT_HEADER_FOOTER, DEFAULT_WATERMARK } from '../stamps'
import { STATIC_FORM_MARK_SIZE } from '../static-form-fill'
import { PAPER_SIZES, parsePageRanges } from '../view-config'
import { DEFAULT_CUTOUT_TOLERANCE, cropRect } from '../image-bake'
import type { CropFractions, ImageBakeOp } from '../image-bake'

/** Text cap per read_pages fed back to the model (the payload is resent in full each turn, so volume must be limited) */
const READ_CHUNK_CHARS = 24_000

/** Capability surface App provides to AI tools; all getters, since the loop outlives render closures */
export type RotateDelta = 90 | -90 | 180
export type TextAlign = 'left' | 'center' | 'right'

/** Fields of a pending insert an edit may replace; offsets ride along when they were remeasured */
export type TextInsertEdit = Partial<
  Pick<TextInsertInput, 'text' | 'fontSize' | 'color' | 'align' | 'lineLeading' | 'lineXOffsets'>
>

export interface PdfAiDeps {
  doc(): PDFDocumentProxy | null
  fileName(): string
  pageCount(): number
  /** Original page number of the currently visible page (1-based) */
  currentPage(): number
  readOnly(): boolean
  /** OCR-recovered text of a scanned page; null when the page was not OCR'd */
  ocrText(origIdx: number): string | null
  /** Text selection cached at mouseup (native DOM selections collapse when focus moves
      into the panel); page..lastPage is the original-page span it covers */
  selection(): { page: number; lastPage: number; text: string } | null
  /** One-line summary of this session's unsaved pending edits; '' when none */
  pendingSummary(): string
  outline(): OutlineNode[] | null
  searchIndex(): Promise<SearchIndex> | null
  isDeleted(origIdx: number): boolean
  /** Original page number → scroll to that page; returns false if the page was deleted */
  gotoPage(origPage: number): boolean
  /** color (rgb 0-1) omitted → the user's current ribbon color for the type */
  addMarkup(
    type: MarkupType,
    origIdx: number,
    rects: [number, number, number, number][],
    color?: [number, number, number],
  ): void
  /** Note threads + text markups on a page: saved (minus pending deletions) with pending overlays */
  annotationsOn(origIdx: number): Promise<{
    threads: NoteThreadItem[]
    markups: { key: string; type: MarkupType; quads: number[][]; saved: boolean }[]
  }>
  /** Counts of saved+pending annotations for the per-run context; '' when none */
  annotationSummary(): string
  /** Queue a pending sticky note authored by the AI; returns its thread key.
      color (rgb 0-1) omitted → the user's current draw color */
  addNote(
    origIdx: number,
    at: [number, number],
    contents: string,
    color?: [number, number, number],
  ): string
  /** Thread root by key ('S<objNum>' saved / 'P<id>' pending); null when not found */
  findNoteRoot(origIdx: number, rootKey: string): Promise<NoteThreadItem | null>
  /** Queue a pending AI-authored reply to a thread root */
  replyToThread(origIdx: number, root: NoteThreadItem, contents: string): void
  /** Rewrite one comment's text (root or reply): pending notes in place, saved notes as a pending content edit */
  editNote(origIdx: number, item: NoteThreadItem, contents: string): void
  /** Remove markups by the keys annotationsOn reports (saved → pending deletion, session → dropped) */
  deleteMarkups(origIdx: number, keys: string[]): Promise<void>
  /** Delete a note thread root and every reply under it */
  deleteNoteThread(origIdx: number, root: NoteThreadItem): void
  /** Queue a pending text edit (dry-run validated against the file when possible); null = accepted, string = rejection reason */
  editText(input: TextEditInput): Promise<string | null>
  /** Queue a pending move of a whole paragraph by a PDF-user-space delta. Stacks onto a
      pending edit that already owns the block, so moveBy is the block's total pending
      displacement (not just this delta) */
  moveTextBlock(
    origIdx: number,
    block: TextBlock,
    delta: [number, number],
  ): Promise<{ reason: string } | { moveBy: [number, number] }>
  /** Queue a pending insert of a new text object (same pipeline as the UI "add text" tool) */
  /** Queue a check/cross mark on a static form at rect (PDF user space); same pipeline as the Fill Form ribbon marks */
  addFormMark(
    origIdx: number,
    kind: 'check' | 'cross',
    rect: [number, number, number, number],
  ): void
  /** Queue a new text block; returns its pending id */
  insertText(input: TextInsertInput): string
  /** Pending (unsaved) inserted text blocks */
  textInserts(): LocalTextInsert[]
  /** Merge edit over the pending block's input (same record the re-edit dialog commits) */
  updateTextInsert(id: string, edit: TextInsertEdit): void
  /** New first-line baseline origin in PDF user space */
  moveTextInsert(id: string, origin: [number, number]): void
  deleteTextInsert(id: string): void
  /** Edit-font ids available on this machine (EDIT_FONTS subset) */
  editFonts(): string[]
  formEdits(): ReadonlyMap<string, FormValueInput>
  applyFormEdit(v: FormValueInput): void
  /** Rotate the given pages (original indices) in one undo step */
  rotatePages(origIdxs: number[], dir: RotateDelta): void
  deletePage(origIdx: number): boolean
  /** Effective document properties: pending unsaved edits over the file's values */
  metadata(): MetadataInput
  /** Replace the pending document properties (applied on save) */
  setMetadata(meta: MetadataInput): void
  /** Current visible page order as original page indices (unsaved reorder applied, deleted pages excluded) */
  pageOrder(): number[]
  /** Move a page (original index) to visible position to (0-based) */
  movePage(origIdx: number, to: number): void
  reversePages(): void
  /** Page geometry (unrotated size + total display rotation); null while the document is loading */
  pageGeom(origIdx: number): PageGeom | null
  /** Content-stream images currently in the saved file (pending unsaved inserts not included) */
  listImages(): Promise<PageImageRef[]>
  /** True when a pending unsaved edit already targets this image */
  isImageClaimed(ref: PageImageRef): boolean
  /** Queue a pending insert of a PNG (base64, no data: prefix) at rect (PDF user space) */
  insertImage(
    origIdx: number,
    png: string,
    rect: [number, number, number, number],
    layer: ImageLayer,
  ): void
  /** Queue a pending move/resize (and optional z-band change / quarter-turn rotation) of an existing image */
  transformImage(
    ref: PageImageRef,
    rect: [number, number, number, number],
    layer?: ImageLayer,
    quarterTurns?: number,
  ): void
  /** Queue a pending in-place pixel swap of an existing image (footprint/z-order kept) */
  replaceImage(ref: PageImageRef, png: string): void
  /** Re-encode an existing image's pixels through the floating-bar bakes; false when the
      pixels could not be read, the signal aborted, or another pending edit claimed the image */
  bakeImage(ref: PageImageRef, op: ImageBakeOp, signal?: AbortSignal): Promise<boolean>
  /** Queue a pending delete of an existing image */
  deleteImage(ref: PageImageRef): void
  searchImages(query: string, maxResults: number): Promise<ImageSearchResponse>
  /** live predicate: gsk login && the Genspark-cloud-tools toggle; false hides generate_image */
  gskTools?(): boolean
  generateImage(op: { prompt: string; aspectRatio?: string }): Promise<{
    url?: string
    error?: string
  }>
  /** Download a URL (main-process, SSRF-guarded) and re-encode as PNG; null on failure */
  fetchImage(url: string): Promise<{ png: string; width: number; height: number } | null>
  /** Session watermark / header-footer configuration; null when none is queued */
  stamps(): StampConfig | null
  /** Replace the session stamp configuration (null clears it); rendered on every page at save */
  setStamps(cfg: StampConfig | null): void
  /** AI create_document: write a new standalone file (pdf/docx/md) into the default folder and open it in a new tab */
  createDocument(request: CreateDocumentRequest): Promise<CreateDocumentResult>
  /** Inline confirmation card for a file-level operation; false when the user declines, the run is stopped, or the panel goes away */
  confirmFileOp(req: FileOpConfirm, signal?: AbortSignal): Promise<boolean>
  /** File-level page operations (below): flush unsaved edits, rewrite the file on disk or write a new
      one, reload. Irreversible, so every call must pass confirmFileOp first. Page indices are positions
      in the flushed file, i.e. visible positions. */
  insertBlankPage(afterVisIdx: number): Promise<FileOpResult>
  setPageSize(width: number, height: number): Promise<FileOpResult>
  cropPages(visIdxs: number[], rect: CropRect): Promise<FileOpResult>
  /** main pops a native picker for the replacement PDF */
  replacePages(visIdxs: number[]): Promise<FileOpResult>
  extractPages(visIdxs: number[]): Promise<NewFileResult>
  /** main pops a native folder picker */
  splitPdf(pagesPerFile: number): Promise<SplitPdfOutcome>
  splitPages(perPage: 2 | 4 | 9): Promise<NewFileResult>
  mergePages(
    perSheet: number,
    direction: 'horizontal' | 'vertical',
    separator: boolean,
  ): Promise<NewFileResult>
}

export interface FileOpConfirm {
  summary: string
  detail?: string
}

export type CropRect = { l: number; t: number; r: number; b: number }

/** A native picker was dismissed; flushed = the pending edits had already been saved to disk first */
export type FileOpCanceled = { ok: true; canceled: true; flushed: boolean }

/** In-place rewrite: pageCount is read from the reloaded document */
export type FileOpResult =
  { ok: true; pageCount: number } | FileOpCanceled | { ok: false; error: string }

export type NewFileResult =
  { ok: true; savedPath: string } | FileOpCanceled | { ok: false; error: string }

export type SplitPdfOutcome =
  { ok: true; savedDir: string; count: number } | FileOpCanceled | { ok: false; error: string }

/** App-side capability surface: everything but the confirmation card, which the panel renders */
export type PdfAppDeps = Omit<PdfAiDeps, 'confirmFileOp'>

export const AGENT_TOOLS: AgentToolDef[] = [
  {
    name: 'read_pages',
    description:
      'Read the text content of a page range (with [Page N] markers). Read the relevant pages before answering questions; at most 10 pages per call, over-long output is truncated.',
    inputSchema: {
      type: 'object',
      properties: {
        start: { type: 'integer', description: 'Start page number (1-based)' },
        end: {
          type: 'integer',
          description: 'End page number (inclusive); if omitted, only the start page is read',
        },
      },
      required: ['start'],
    },
  },
  {
    name: 'search_text',
    description:
      'Search the full text for a string; returns the page number and a context excerpt for each hit. Prefer this when locating which page something is on.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Text to search for (case-insensitive)' },
      },
      required: ['query'],
    },
  },
  {
    name: 'goto_page',
    description: 'Scroll the reading view to the given page so the user can see it.',
    inputSchema: {
      type: 'object',
      properties: { page: { type: 'integer', description: 'Page number (1-based)' } },
      required: ['page'],
    },
  },
  {
    name: 'markup_text',
    description:
      'Add a markup (highlight/underline/strikeout) to a text passage on the given page. text must be a verbatim fragment that actually exists on that page (confirm with read_pages or search_text first); by default only the first occurrence is marked, all=true marks every occurrence on that page.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        text: { type: 'string', description: 'Verbatim text fragment from the page' },
        type: {
          type: 'string',
          enum: ['highlight', 'underline', 'strikeout'],
          description: 'Markup type',
        },
        all: {
          type: 'boolean',
          description: 'Whether to mark every occurrence on the page; defaults to false',
        },
        color: {
          type: 'string',
          description:
            "Markup color as #RRGGBB hex; omit to use the user's current color for the type",
        },
      },
      required: ['page', 'text', 'type'],
    },
  },
  {
    name: 'read_annotations',
    description:
      'List the comment notes (sticky notes, with their reply threads) and text markups (highlight/underline/strikeout) in the document — both saved in the file and queued unsaved this session. Optional page range; omit to read the whole document. Reply to a note with reply_note or change its text with edit_note using the returned note ids (replies have their own ids).',
    inputSchema: {
      type: 'object',
      properties: {
        start: { type: 'integer', description: 'First page (1-based); omit to start at page 1' },
        end: { type: 'integer', description: 'Last page (inclusive); omit to read to the end' },
      },
    },
  },
  {
    name: 'add_note',
    description:
      'Add a sticky-note comment to a page (takes effect on save; authored as "AI Assistant"). Anchor it with anchor_text (a verbatim fragment on the page — the note pin lands at its end) or explicit x/y in points from the page top-left as displayed.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        text: { type: 'string', description: 'Note contents' },
        anchor_text: {
          type: 'string',
          description: 'Verbatim text fragment on the page the note refers to',
        },
        x: { type: 'number', description: 'Pin x in points from the page left edge' },
        y: { type: 'number', description: 'Pin y in points from the page TOP edge' },
        color: {
          type: 'string',
          description: "Note pin color as #RRGGBB hex; omit to use the user's current draw color",
        },
      },
      required: ['page', 'text'],
    },
  },
  {
    name: 'reply_note',
    description:
      'Append a reply to an existing note thread (takes effect on save; authored as "AI Assistant"). note_id is a thread id from read_annotations.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based) the thread is on' },
        note_id: { type: 'string', description: 'Thread id from read_annotations (e.g. "S12")' },
        text: { type: 'string', description: 'Reply contents' },
      },
      required: ['page', 'note_id', 'text'],
    },
  },
  {
    name: 'edit_note',
    description:
      'Replace the text of an existing sticky-note comment — a thread root or a reply — keeping its author, position, and thread (takes effect on save). note_id is a note or reply id from read_annotations. Only edit notes the user explicitly asked to change.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based) the note is on' },
        note_id: {
          type: 'string',
          description: 'Note or reply id from read_annotations (e.g. "S12", "Pd3k")',
        },
        text: { type: 'string', description: 'New note contents (replaces the old text)' },
      },
      required: ['page', 'note_id', 'text'],
    },
  },
  {
    name: 'delete_markup',
    description:
      'Remove text markups (highlight / underline / strikeout) from a page, whether saved in the file or added this session (takes effect on save). Call read_annotations first: it lists every markup with its id. Pass markup_ids to remove specific ones, or omit it to remove all markups on the page; type narrows either form to one kind.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        markup_ids: {
          type: 'array',
          items: { type: 'string' },
          description: 'Markup ids from read_annotations; omit to target every markup on the page',
        },
        type: {
          type: 'string',
          enum: ['highlight', 'underline', 'strikeout'],
          description: 'Only remove markups of this kind',
        },
      },
      required: ['page'],
    },
  },
  {
    name: 'delete_note',
    description:
      'Delete a sticky-note thread (the root comment and all of its replies) from a page (takes effect on save). note_id is the thread id shown by read_annotations. Only delete notes the user explicitly asked to remove.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        note_id: { type: 'string', description: 'Thread id from read_annotations (e.g. "S12")' },
      },
      required: ['page', 'note_id'],
    },
  },
  {
    name: 'edit_text',
    description:
      'Replace a short text run on a page (rewrites the PDF content; takes effect on save). old_text must be a verbatim fragment that actually exists on that page (confirm with read_pages or search_text first); only the first occurrence on the page is edited unless occurrence is given. The replacement is drawn from the original position without reflowing the page, so keep it close to the original length. An empty new_text deletes the run outright (the space it occupied stays blank; nothing moves up) — never substitute a placeholder such as "." for deletion. Alignment cannot be changed here (a run keeps its position); use edit_block with align for that.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        old_text: { type: 'string', description: 'Verbatim text fragment currently on the page' },
        new_text: {
          type: 'string',
          description: 'Replacement text; "\\n" splits it into stacked lines; "" deletes the run',
        },
        occurrence: {
          type: 'integer',
          description: 'Which occurrence of old_text on the page to edit (1-based); defaults to 1',
        },
        font_size: {
          type: 'number',
          description: 'New font size in PDF points; omit to keep the original size',
        },
        color: {
          type: 'string',
          description: 'New text color as #RRGGBB hex; omit to keep the original color',
        },
        font: {
          type: 'string',
          enum: ['arial', 'times', 'courier'],
          description:
            'Font for the replacement; ignored when the face lacks glyphs for it (e.g. CJK). Omit for automatic',
        },
        bold: { type: 'boolean', description: 'Set the replacement in the bold variant' },
        italic: { type: 'boolean', description: 'Set the replacement in the italic variant' },
      },
      required: ['page', 'old_text', 'new_text'],
    },
  },
  {
    name: 'edit_block',
    description:
      "Rewrite a whole paragraph on a page (rewrites the PDF content; takes effect on save). The paragraph is located by paragraph_text — a distinctive verbatim fragment of it. The ENTIRE paragraph is replaced by new_text, which is re-wrapped automatically within the paragraph's original width (the block grows downward when the text is longer). An empty new_text deletes the whole paragraph (its area stays blank; content below does not move up). Use edit_text instead to change a few words without reflowing.",
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        paragraph_text: {
          type: 'string',
          description:
            'Verbatim fragment identifying the paragraph; must match exactly one paragraph on the page',
        },
        new_text: {
          type: 'string',
          description:
            'Full replacement for the paragraph; "\\n" separates paragraphs within the block; "" deletes the paragraph',
        },
        font_size: {
          type: 'number',
          description: 'New font size in PDF points; omit to keep the original size',
        },
        color: {
          type: 'string',
          description: 'New text color as #RRGGBB hex; omit to keep the original color',
        },
        font: {
          type: 'string',
          enum: ['arial', 'times', 'courier'],
          description:
            'Font for the replacement; ignored when the face lacks glyphs for it (e.g. CJK). Omit for automatic',
        },
        bold: { type: 'boolean', description: 'Set the paragraph in the bold variant' },
        italic: { type: 'boolean', description: 'Set the paragraph in the italic variant' },
        align: {
          type: 'string',
          enum: ['left', 'center', 'right'],
          description:
            "Alignment of the reflowed lines within the paragraph's original width; omit to keep the paragraph's current alignment",
        },
      },
      required: ['page', 'paragraph_text', 'new_text'],
    },
  },
  {
    name: 'move_text_block',
    description:
      'Move a whole paragraph to another position on its page without changing its text (takes effect on save; fonts and spacing are kept as-is). The paragraph is located like in edit_block: paragraph_text must match exactly one paragraph. dx/dy are the offset in PDF points as displayed: positive dx moves right, positive dy moves DOWN on screen (negative dy moves up). Nothing else on the page moves, so check the target area is free (read the page layout first). Typical use: after deleting a paragraph, move the following paragraph up by the deleted height to close the gap.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        paragraph_text: {
          type: 'string',
          description:
            'Verbatim fragment identifying the paragraph; must match exactly one paragraph on the page',
        },
        dx: { type: 'number', description: 'Horizontal offset in points; positive moves right' },
        dy: {
          type: 'number',
          description: 'Vertical offset in points as displayed; positive moves down, negative up',
        },
      },
      required: ['page', 'paragraph_text', 'dx', 'dy'],
    },
  },
  {
    name: 'insert_text',
    description:
      'Add NEW text to a page (drawn as new content on top of the page; takes effect on save). Use this for blank pages, empty areas, and filling in non-interactive form blanks — to change text that already exists, use edit_text or edit_block instead. Position with anchor_text (a verbatim fragment on the page, e.g. a form label) plus placement, or with x/y as the TOP-LEFT corner of the text block in points measured from the page top-left as displayed; with neither, the block is centered horizontally near the top. "\\n" starts a new line; pass max_width to auto-wrap long paragraphs within that width.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        text: { type: 'string', description: 'Text to insert; "\\n" separates lines/paragraphs' },
        anchor_text: {
          type: 'string',
          description:
            'Verbatim text fragment on the page (e.g. a form label) to position the new text against',
        },
        placement: {
          type: 'string',
          enum: ['right', 'below', 'above'],
          description: 'Which side of anchor_text to place the text on; defaults to right',
        },
        x: {
          type: 'number',
          description: 'Left edge of the text block in points from the page left edge',
        },
        y: {
          type: 'number',
          description: 'Top edge of the text block in points from the page TOP edge',
        },
        font_size: { type: 'number', description: 'Font size in PDF points; defaults to 14' },
        color: { type: 'string', description: 'Text color as #RRGGBB hex; defaults to black' },
        max_width: {
          type: 'number',
          description:
            'Wrap width in points: paragraphs are re-wrapped to fit; omit to keep the given line breaks as-is',
        },
        align: {
          type: 'string',
          enum: ['left', 'center', 'right'],
          description: 'Line alignment within max_width (or the widest line); defaults to left',
        },
        font: {
          type: 'string',
          enum: ['arial', 'times', 'courier'],
          description:
            'Font for the text; ignored when the face lacks glyphs for it (e.g. CJK). Omit for automatic',
        },
        bold: { type: 'boolean', description: 'Set the text in the bold variant' },
        italic: { type: 'boolean', description: 'Set the text in the italic variant' },
      },
      required: ['page', 'text'],
    },
  },
  {
    name: 'add_form_mark',
    description:
      'Place a check mark or cross on a page (drawn as new content; takes effect on save). Use it to tick check boxes printed on non-interactive forms — interactive check boxes are set with fill_form_field instead. Position exactly like insert_text: anchor_text (verbatim fragment, e.g. the label next to the box) plus placement, with the mark centered against that side of the anchor, or x/y as the TOP-LEFT corner of the mark in points from the page top-left as displayed.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        kind: { type: 'string', enum: ['check', 'cross'], description: 'Mark to draw' },
        anchor_text: {
          type: 'string',
          description:
            'Verbatim text fragment on the page (e.g. the box label) to position the mark against',
        },
        placement: {
          type: 'string',
          enum: ['right', 'left', 'below', 'above'],
          description: 'Which side of anchor_text the mark goes on; defaults to right',
        },
        x: {
          type: 'number',
          description: 'Left edge of the mark in points from the page left edge',
        },
        y: { type: 'number', description: 'Top edge of the mark in points from the page TOP edge' },
        size: {
          type: 'number',
          description: `Side of the square mark in points; defaults to ${STATIC_FORM_MARK_SIZE}`,
        },
      },
      required: ['page', 'kind'],
    },
  },
  {
    name: 'list_inserted_text',
    description:
      'List the text blocks added with insert_text this session that are still unsaved: id, page, top-left position in points as displayed, font size, and a text preview. Their ids feed edit_inserted_text / move_inserted_text / delete_inserted_text. Once the user saves, inserted text becomes page content and is no longer listed — change it with edit_text then.',
    inputSchema: {
      type: 'object',
      properties: {
        page: {
          type: 'integer',
          description: 'Only list blocks on this page (1-based); omit for all',
        },
      },
    },
  },
  {
    name: 'edit_inserted_text',
    description:
      'Change an unsaved inserted text block (same as the user double-clicking it): new text, font size, color, or alignment; omitted fields keep their value. The block stays anchored where it is. Get ids from list_inserted_text or the insert_text output. For text that is already saved use edit_text instead.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'Inserted text id (e.g. "Td3k9a1")' },
        text: { type: 'string', description: 'Replacement text; "\\n" separates lines' },
        font_size: { type: 'number', description: 'New font size in PDF points' },
        color: { type: 'string', description: 'New text color as #RRGGBB hex' },
        align: {
          type: 'string',
          enum: ['left', 'center', 'right'],
          description: 'New line alignment',
        },
      },
      required: ['id'],
    },
  },
  {
    name: 'move_inserted_text',
    description:
      'Move an unsaved inserted text block (same as the user dragging it). Give x/y as the new TOP-LEFT corner in points from the page top-left as displayed, or dx/dy as an offset in points (positive dy moves down). Only pending inserts can be moved; saved text is page content.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'Inserted text id from list_inserted_text' },
        x: { type: 'number', description: 'New left edge in points from the page left edge' },
        y: { type: 'number', description: 'New top edge in points from the page top edge' },
        dx: { type: 'number', description: 'Horizontal offset in points (positive = right)' },
        dy: { type: 'number', description: 'Vertical offset in points (positive = down)' },
      },
      required: ['id'],
    },
  },
  {
    name: 'delete_inserted_text',
    description:
      'Remove an unsaved inserted text block (same as the user pressing Delete on it). Only pending inserts can be removed this way; to delete saved text use edit_text with an empty new_text.',
    inputSchema: {
      type: 'object',
      properties: {
        id: { type: 'string', description: 'Inserted text id from list_inserted_text' },
      },
      required: ['id'],
    },
  },
  {
    name: 'image_search',
    description:
      'Search the web for images. Returns a numbered list with direct imageUrl links; pick one and pass its URL to insert_image. Use for real photos/logos; use generate_image for custom illustrations.',
    inputSchema: {
      type: 'object',
      properties: {
        query: { type: 'string', description: 'Image search keywords (English works better)' },
        max_results: { type: 'integer', description: 'Max results, default 8' },
      },
      required: ['query'],
    },
  },
  {
    name: 'generate_image',
    description:
      'Generate an image with AI from a text prompt; returns an image URL to pass to insert_image. Use for custom illustrations/icons/diagrams; for real photos prefer image_search.',
    inputSchema: {
      type: 'object',
      properties: {
        prompt: {
          type: 'string',
          description:
            'Image description, English works better (keep any text to render in the image verbatim)',
        },
        aspect_ratio: {
          type: 'string',
          description: 'Aspect ratio: 1:1|4:3|16:9|9:16|3:4|2:3|3:2|auto',
        },
      },
      required: ['prompt'],
    },
  },
  {
    name: 'list_page_images',
    description:
      'List the images embedded in the page content: per-page image numbers with position/size in points as displayed (x from the left edge, y from the page TOP, page rotation applied). Call this before transform_image or delete_image. Images inserted in this session but not yet saved are not listed.',
    inputSchema: {
      type: 'object',
      properties: {
        page: {
          type: 'integer',
          description: 'Page number (1-based); omit to list images on every page',
        },
      },
    },
  },
  {
    name: 'insert_image',
    description:
      'Download an image URL (from image_search or generate_image) and place it on a page (takes effect on save). Position it either with anchor_text (a verbatim text fragment on the page) plus placement, or with explicit x/y in points measured from the page top-left as displayed; with neither, the image is centered on the page.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        url: {
          type: 'string',
          description: 'Direct image link (from image_search or generate_image)',
        },
        anchor_text: {
          type: 'string',
          description: 'Verbatim text fragment on the page to position the image relative to',
        },
        placement: {
          type: 'string',
          enum: ['below', 'above', 'right', 'left'],
          description: 'Which side of anchor_text to place the image on; defaults to below',
        },
        x: {
          type: 'number',
          description: 'Left edge in points from the page left edge as displayed',
        },
        y: {
          type: 'number',
          description: 'Top edge in points from the page TOP edge as displayed',
        },
        width: {
          type: 'number',
          description:
            'Display width in PDF points; height follows the aspect ratio. Default: natural size, capped to half the page width',
        },
        layer: {
          type: 'string',
          enum: ['above_text', 'below_text'],
          description: 'Z-order relative to the page text; below_text (default) never covers text',
        },
      },
      required: ['page', 'url'],
    },
  },
  {
    name: 'transform_image',
    description:
      'Move/resize an existing page image, or change its z-order relative to the text (takes effect on save). Call list_page_images first; image_number refers to that listing. Omitted parameters keep their current value; giving only width (or only height) scales the other side by the aspect ratio.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        image_number: {
          type: 'integer',
          description: 'Image number from list_page_images (1-based, per page)',
        },
        x: {
          type: 'number',
          description: 'New left edge in points from the page left edge as displayed',
        },
        y: {
          type: 'number',
          description: 'New top edge in points from the page TOP edge as displayed',
        },
        width: { type: 'number', description: 'New width in PDF points' },
        height: { type: 'number', description: 'New height in PDF points' },
        layer: {
          type: 'string',
          enum: ['above_text', 'below_text'],
          description: 'Z-order relative to the page text',
        },
      },
      required: ['page', 'image_number'],
    },
  },
  {
    name: 'rotate_image',
    description:
      'Rotate an existing page image about its center (takes effect on save). Call list_page_images first; image_number refers to that listing. Quarter turns only; the footprint width/height swap for 90°/270°.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        image_number: {
          type: 'integer',
          description: 'Image number from list_page_images (1-based, per page)',
        },
        direction: {
          type: 'string',
          enum: ['cw', 'ccw', '180'],
          description: 'cw = 90° clockwise (default), ccw = 90° counter-clockwise, 180 = half turn',
        },
      },
      required: ['page', 'image_number'],
    },
  },
  {
    name: 'replace_image',
    description:
      'Swap an existing page image\'s pixels for a downloaded URL (from image_search or generate_image) in place — footprint and z-order survive; the new image stretches to the old footprint (takes effect on save). Call list_page_images first; image_number refers to that listing. This is the tool for "change/AI-edit this image": generate_image with the desired edit, then replace_image with the returned URL.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        image_number: {
          type: 'integer',
          description: 'Image number from list_page_images (1-based, per page)',
        },
        url: { type: 'string', description: 'Direct image link (PNG/JPEG)' },
      },
      required: ['page', 'image_number', 'url'],
    },
  },
  {
    name: 'flip_image',
    description:
      'Mirror an existing page image horizontally (left-right as displayed) or vertically (top-bottom as displayed) in place; footprint and z-order kept (takes effect on save). Call list_page_images first; image_number refers to that listing.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        image_number: {
          type: 'integer',
          description: 'Image number from list_page_images (1-based, per page)',
        },
        axis: {
          type: 'string',
          enum: ['horizontal', 'vertical'],
          description:
            "'horizontal' mirrors left-right, 'vertical' mirrors top-bottom, both as the page is displayed",
        },
      },
      required: ['page', 'image_number', 'axis'],
    },
  },
  {
    name: 'set_image_opacity',
    description:
      'Make an existing page image semi-transparent (fade/watermark look) by baking the opacity into its pixels; footprint and z-order kept (takes effect on save). Absolute, not cumulative. Call list_page_images first; image_number refers to that listing.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        image_number: {
          type: 'integer',
          description: 'Image number from list_page_images (1-based, per page)',
        },
        opacity: {
          type: 'number',
          description:
            'Resulting opacity from 0 (invisible) to 1 (fully opaque); e.g. 0.5 = half transparent',
        },
      },
      required: ['page', 'image_number', 'opacity'],
    },
  },
  {
    name: 'crop_image',
    description:
      "Crop an existing page image: trims the given margins off the picture and shrinks its footprint on the page to the kept region (the kept pixels stay at their current size and place; z-order kept; takes effect on save). Each inset is a fraction 0-1 of the image's current displayed width (left/right) or height (top/bottom) as listed by list_page_images, measured inward from that edge; omitted insets are 0. Example: left 0.25 removes the left quarter. Call list_page_images first; image_number refers to that listing.",
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        image_number: {
          type: 'integer',
          description: 'Image number from list_page_images (1-based, per page)',
        },
        left: {
          type: 'number',
          description: 'Fraction of the width to trim from the left edge (0-1)',
        },
        top: {
          type: 'number',
          description: 'Fraction of the height to trim from the top edge (0-1)',
        },
        right: {
          type: 'number',
          description: 'Fraction of the width to trim from the right edge (0-1)',
        },
        bottom: {
          type: 'number',
          description: 'Fraction of the height to trim from the bottom edge (0-1)',
        },
      },
      required: ['page', 'image_number'],
    },
  },
  {
    name: 'remove_image_background',
    description:
      'Remove the background of an existing page image (makes the edge-connected background color transparent, like "Remove Background" in Office); footprint and z-order kept (takes effect on save). Works best on photos/logos on a plain background. Call list_page_images first; image_number refers to that listing.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        image_number: {
          type: 'integer',
          description: 'Image number from list_page_images (1-based, per page)',
        },
        tolerance: {
          type: 'number',
          description: `Color tolerance 0-100 (default ${DEFAULT_CUTOUT_TOLERANCE}); raise it when background remains, lower it when parts of the subject disappear`,
        },
      },
      required: ['page', 'image_number'],
    },
  },
  {
    name: 'delete_image',
    description:
      'Delete an existing page image (takes effect on save; the user can undo before saving). Call list_page_images first; image_number refers to that listing.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        image_number: {
          type: 'integer',
          description: 'Image number from list_page_images (1-based, per page)',
        },
      },
      required: ['page', 'image_number'],
    },
  },
  {
    name: 'list_form_fields',
    description:
      'List all form fields in the document (name/type/current value/options/page). Must be called before filling forms to learn the fields.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'fill_form_field',
    description:
      'Fill in one form field. For text/choice/radio fields pass value (radio: the exportValue; choice: an option exportValue); for checkboxes pass checked.',
    inputSchema: {
      type: 'object',
      properties: {
        name: { type: 'string', description: 'Field name (from list_form_fields)' },
        value: { type: 'string', description: 'Value for text/choice/radio fields' },
        checked: { type: 'boolean', description: 'Checked state for checkboxes' },
      },
      required: ['name'],
    },
  },
  {
    name: 'rotate_page',
    description:
      'Rotate one page, or every page with all=true, by a quarter or half turn (takes effect on save).',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based); ignored when all is true' },
        direction: {
          type: 'string',
          enum: ['left', 'right', '180'],
          description: 'left = 90° counterclockwise, right = 90° clockwise, 180 = half turn',
        },
        all: { type: 'boolean', description: 'Rotate all pages instead of one (default false)' },
      },
      required: ['direction'],
    },
  },
  {
    name: 'set_metadata',
    description:
      'Set document properties (takes effect on save). Omitted fields keep their current value; pass "" to clear one. Call with no fields to read the current values.',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string' },
        author: { type: 'string' },
        subject: { type: 'string' },
        keywords: { type: 'string', description: 'Comma-separated keywords' },
      },
    },
  },
  {
    name: 'delete_page',
    description: 'Delete the given page (takes effect on save; the user can undo before saving).',
    inputSchema: {
      type: 'object',
      properties: { page: { type: 'integer', description: 'Page number (1-based)' } },
      required: ['page'],
    },
  },
  {
    name: 'move_page',
    description:
      "Move a page to another position in the page order (takes effect on save). page is the document's original page number (as in the [Page N] markers, unchanged by earlier moves); to is the 1-based target position in the current visible order. Example: move_page(page 5, to 1) makes page 5 the first page.",
    inputSchema: {
      type: 'object',
      properties: {
        page: {
          type: 'integer',
          description: 'Original page number of the page to move (1-based)',
        },
        to: { type: 'integer', description: 'Target position in the current order (1-based)' },
      },
      required: ['page', 'to'],
    },
  },
  {
    name: 'reverse_pages',
    description: 'Reverse the order of all pages (takes effect on save).',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'insert_blank_page',
    description:
      'Insert a blank page (same size as its neighbor) after the given page; after_page 0 inserts it as the first page. IRREVERSIBLE FILE OPERATION: the user is asked to confirm in a card first; on confirm every unsaved edit is saved and the file on disk is rewritten, which cannot be undone. Afterwards the document reloads and page numbers change; re-read (search_text/read_pages) before any further edit. Prefer delete_page/move_page/rotate_page when they suffice.',
    inputSchema: {
      type: 'object',
      properties: {
        after_page: {
          type: 'integer',
          description:
            'Original page number to insert after (1-based); 0 = insert as the first page',
        },
      },
      required: ['after_page'],
    },
  },
  {
    name: 'set_page_size',
    description:
      'Resize every page of the document to a paper size (content scaled to fit, portrait). Pass either preset or both width and height in points. IRREVERSIBLE FILE OPERATION: the user is asked to confirm in a card first; on confirm every unsaved edit is saved and the file on disk is rewritten, which cannot be undone. Afterwards the document reloads and page numbers change; re-read (search_text/read_pages) before any further edit. Prefer delete_page/move_page/rotate_page when they suffice.',
    inputSchema: {
      type: 'object',
      properties: {
        preset: {
          type: 'string',
          enum: PAPER_SIZES.map((p) => p.label),
          description: 'Paper size preset',
        },
        width: { type: 'number', description: 'Page width in points (when no preset)' },
        height: { type: 'number', description: 'Page height in points (when no preset)' },
      },
    },
  },
  {
    name: 'crop_pages',
    description:
      'Crop pages to a sub-rectangle of the displayed page. left/top/right/bottom are fractions 0-1 of the page width/height describing the rectangle that is KEPT (left=0, top=0, right=1, bottom=1 is the whole page and is rejected); e.g. left 0.1, right 0.9 trims 10% off each side. IRREVERSIBLE FILE OPERATION: the user is asked to confirm in a card first; on confirm every unsaved edit is saved and the file on disk is rewritten, which cannot be undone. Afterwards the document reloads and page numbers change; re-read (search_text/read_pages) before any further edit. Prefer delete_page/move_page/rotate_page when they suffice.',
    inputSchema: {
      type: 'object',
      properties: {
        pages: {
          type: 'string',
          description: '"all" or original page numbers as a range string, e.g. "1-3,5"',
        },
        left: { type: 'number', description: 'Left edge of the kept area (0-1)' },
        top: { type: 'number', description: 'Top edge of the kept area (0-1)' },
        right: { type: 'number', description: 'Right edge of the kept area (0-1)' },
        bottom: { type: 'number', description: 'Bottom edge of the kept area (0-1)' },
      },
      required: ['pages', 'left', 'top', 'right', 'bottom'],
    },
  },
  {
    name: 'extract_pages',
    description:
      'Copy the given pages into a new PDF. IRREVERSIBLE FILE OPERATION: the user is asked to confirm in a card first; on confirm every unsaved edit is saved to disk (cannot be undone) and the result is written as a NEW file in the default save folder and opened in a new tab; the current document keeps its pages.',
    inputSchema: {
      type: 'object',
      properties: {
        pages: {
          type: 'string',
          description: 'Original page numbers as a range string, e.g. "1-3,5"',
        },
      },
      required: ['pages'],
    },
  },
  {
    name: 'split_pdf',
    description:
      'Split the document into several PDF files of pages_per_file pages each (at least two files). A native folder picker opens for the user to choose where the files go; the current document keeps its pages. IRREVERSIBLE FILE OPERATION: the user is asked to confirm in a card first; on confirm every unsaved edit is saved to disk (cannot be undone) and the result is written as a NEW file in the default save folder and opened in a new tab; the current document keeps its pages.',
    inputSchema: {
      type: 'object',
      properties: {
        pages_per_file: {
          type: 'integer',
          description: 'Pages per output file (1 to page count - 1)',
        },
      },
      required: ['pages_per_file'],
    },
  },
  {
    name: 'split_pages',
    description:
      'Cut every page into a grid of 2, 4, or 9 smaller pages (inverse of merge_pages). IRREVERSIBLE FILE OPERATION: the user is asked to confirm in a card first; on confirm every unsaved edit is saved to disk (cannot be undone) and the result is written as a NEW file in the default save folder and opened in a new tab; the current document keeps its pages.',
    inputSchema: {
      type: 'object',
      properties: {
        per_page: { type: 'integer', enum: [2, 4, 9], description: 'Pieces per page' },
      },
      required: ['per_page'],
    },
  },
  {
    name: 'merge_pages',
    description:
      'N-up imposition: place per_sheet consecutive pages onto one sheet (e.g. 2 = two pages side by side). IRREVERSIBLE FILE OPERATION: the user is asked to confirm in a card first; on confirm every unsaved edit is saved to disk (cannot be undone) and the result is written as a NEW file in the default save folder and opened in a new tab; the current document keeps its pages.',
    inputSchema: {
      type: 'object',
      properties: {
        per_sheet: { type: 'integer', description: 'Pages per sheet, 2-16' },
        direction: {
          type: 'string',
          enum: ['horizontal', 'vertical'],
          description:
            'Fill order: horizontal = left to right then down, vertical = top to bottom then right (default)',
        },
        separator: {
          type: 'boolean',
          description: 'Draw hairlines between the placed pages (default false)',
        },
      },
      required: ['per_sheet'],
    },
  },
  {
    name: 'replace_pages',
    description:
      'Replace the given pages with all pages of another PDF. A native file picker opens for the user to choose that PDF. IRREVERSIBLE FILE OPERATION: the user is asked to confirm in a card first; on confirm every unsaved edit is saved and the file on disk is rewritten, which cannot be undone. Afterwards the document reloads and page numbers change; re-read (search_text/read_pages) before any further edit. Prefer delete_page/move_page/rotate_page when they suffice.',
    inputSchema: {
      type: 'object',
      properties: {
        pages: {
          type: 'string',
          description: 'Original page numbers to replace as a range string, e.g. "1-3,5"',
        },
      },
      required: ['pages'],
    },
  },
  {
    name: 'get_outline',
    description:
      'Read the document outline (bookmarks) tree, including entry titles. Returns empty if the document has no outline.',
    inputSchema: { type: 'object', properties: {} },
  },
  {
    name: 'set_watermark',
    description:
      'Add or replace a text watermark drawn diagonally across every page (takes effect on save); pass text "" to remove the watermark added this session. It is stamped above the page content and cannot remove a watermark that is part of the original document.',
    inputSchema: {
      type: 'object',
      properties: {
        text: { type: 'string', description: 'Watermark text; "" removes it' },
        angle: {
          type: 'number',
          description: `Counterclockwise angle in degrees (default ${DEFAULT_WATERMARK.angle})`,
        },
        opacity: {
          type: 'number',
          description: `Opacity 0–1 (default ${DEFAULT_WATERMARK.opacity})`,
        },
        color: {
          type: 'string',
          description: `Text color as #RRGGBB (default ${DEFAULT_WATERMARK.color})`,
        },
        size: {
          type: 'number',
          description: `Font size as a fraction of the page width, 0.02–0.5 (default ${DEFAULT_WATERMARK.sizeRatio})`,
        },
      },
      required: ['text'],
    },
  },
  {
    name: 'set_header_footer',
    description:
      'Set the header and footer stamped on every page (takes effect on save): six text slots (header/footer × left/center/right) plus an automatic page number in the footer center. Text may contain {page} and {total} placeholders. The call replaces the header/footer set earlier this session — omitted slots become empty — and a call with every slot empty and page_number false removes it.',
    inputSchema: {
      type: 'object',
      properties: {
        header_left: { type: 'string' },
        header_center: { type: 'string' },
        header_right: { type: 'string' },
        footer_left: { type: 'string' },
        footer_center: { type: 'string' },
        footer_right: { type: 'string' },
        page_number: {
          type: 'boolean',
          description:
            'Print the page number in the footer center (overrides footer_center); default false',
        },
        start_at: {
          type: 'integer',
          description: `Number of the first page when page_number is on (default ${DEFAULT_HEADER_FOOTER.startAt})`,
        },
        font_size: {
          type: 'number',
          description: `Font size in points (default ${DEFAULT_HEADER_FOOTER.fontSize})`,
        },
        color: {
          type: 'string',
          description: `Text color as #RRGGBB (default ${DEFAULT_HEADER_FOOTER.color})`,
        },
      },
    },
  },
  {
    name: 'create_document',
    description:
      'Create a NEW standalone file in the default save folder and open it in a new tab; the current PDF is not modified. Use when the user asks to put content (a summary, an extraction, an analysis result) into a new/separate document. ' +
      "type 'pdf' (default) and 'docx' take simple HTML in content (<h1>-<h6>, <p>, <ul>/<ol>/<li>, <table>, <pre>, <blockquote>; inline <strong>/<em>/<u>/<s>); type 'md' takes Markdown source. Images are not supported in the new file's content.",
    inputSchema: {
      type: 'object',
      properties: {
        type: {
          type: 'string',
          enum: ['pdf', 'docx', 'md'],
          description: "target file type (default 'pdf')",
        },
        title: { type: 'string', description: 'document title, used as the file name' },
        content: {
          type: 'string',
          description: 'full document content: simple HTML for pdf/docx, Markdown for md',
        },
      },
      required: ['title', 'content'],
    },
  },
]

const READONLY_OUTPUT =
  'The document is encrypted and read-only; it cannot be modified. Inform the user.'

/** mirror of the docs tool-echo guard: reject tool-protocol output pasted as document content */
function contentEchoError(content: string): string | null {
  if (/<\/?tool_response>/i.test(content)) {
    return 'content contains a literal <tool_response> tag — that is tool-protocol output, not document content; retry with the actual document HTML'
  }
  if (/"index"\s*:\s*\d+\s*,\s*"type"\s*:\s*"/.test(content)) {
    return 'content contains a raw JSON dump, not an HTML fragment; retry with simple HTML (e.g. <p>…</p>)'
  }
  return null
}

function err(output: string, summary: string): ToolExecution {
  return { output, isError: true, summary }
}

/** Validate a 1-based page number; returns the original page index or an error */
function resolvePage(deps: PdfAiDeps, raw: unknown): { origIdx: number } | { bad: string } {
  const page = Number(raw)
  if (!Number.isInteger(page) || page < 1 || page > deps.pageCount()) {
    return {
      bad: `Page number ${String(raw)} is out of range (document has ${deps.pageCount()} pages)`,
    }
  }
  if (deps.isDeleted(page - 1)) return { bad: `Page ${page} has been deleted (unsaved)` }
  return { origIdx: page - 1 }
}

async function readPages(deps: PdfAiDeps, input: Record<string, unknown>): Promise<ToolExecution> {
  const doc = deps.doc()
  if (!doc) return err('Document not ready', t('aiToolReadPages', { start: '?', end: '?' }))
  const start = Number(input.start)
  const end = Math.min(Number(input.end ?? start), start + 9)
  const summary = t('aiToolReadPages', { start, end })
  if (!Number.isInteger(start) || start < 1 || end < start || start > doc.numPages) {
    return err(`Invalid page range (document has ${doc.numPages} pages)`, summary)
  }
  let out = ''
  for (let n = start; n <= Math.min(end, doc.numPages); n++) {
    const page = await doc.getPage(n)
    const content = await page.getTextContent()
    let text = ''
    for (const item of content.items) {
      if ('str' in item) {
        text += item.str
        if (item.hasEOL) text += '\n'
      }
    }
    page.cleanup()
    if (isScannedText(text)) {
      const ocr = deps.ocrText(n - 1)
      if (ocr && ocr.trim()) {
        text = `(recovered by OCR; may contain recognition errors)\n${ocr}`
      }
    }
    out += `[Page ${n}]\n${text.trim()}\n\n`
    if (out.length > READ_CHUNK_CHARS) {
      out = `${out.slice(0, READ_CHUNK_CHARS)}\n… (truncated; read the rest in further calls)`
      break
    }
  }
  return {
    output: out.trim() || '(No extractable text in this range; the pages may be scanned images)',
    summary,
  }
}

async function searchText(deps: PdfAiDeps, input: Record<string, unknown>): Promise<ToolExecution> {
  const query = String(input.query ?? '').trim()
  if (!query) return err('query must not be empty', t('aiToolSearch', { query: '', count: 0 }))
  const indexPromise = deps.searchIndex()
  if (!indexPromise) return err('Document not ready', t('aiToolSearch', { query, count: 0 }))
  const index = await indexPromise
  const matches = searchInIndex(index, query)
  const lines: string[] = []
  for (const m of matches.slice(0, 40)) {
    const entry = index[m.pageIndex]!
    const pos = entry.lower.indexOf(query.toLowerCase())
    const from = Math.max(0, pos - 40)
    const snippet = entry.text.slice(from, pos + query.length + 40).replace(/\s+/g, ' ')
    lines.push(`Page ${m.pageIndex + 1}: …${snippet}…`)
  }
  if (matches.length > 40) lines.push(`(${matches.length} matches total; only the first 40 listed)`)
  return {
    output: lines.join('\n') || 'No matches found',
    summary: t('aiToolSearch', { query, count: matches.length }),
  }
}

async function markupText(deps: PdfAiDeps, input: Record<string, unknown>): Promise<ToolExecution> {
  const type = String(input.type) as MarkupType
  const summary = t('aiToolMarkup', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  if (!['highlight', 'underline', 'strikeout'].includes(type))
    return err(`Invalid type: ${type}`, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const text = String(input.text ?? '').trim()
  if (!text) return err('text must not be empty', summary)
  const color = hexToRgb01(input.color)
  if (color === null) return err(`Invalid color "${String(input.color)}"; use #RRGGBB`, summary)
  const indexPromise = deps.searchIndex()
  if (!indexPromise) return err('Document not ready', summary)
  const index = await indexPromise
  const onPage = searchInIndex(index, text).filter((m) => m.pageIndex === r.origIdx)
  if (onPage.length === 0) {
    return err(
      `"${text}" not found on page ${r.origIdx + 1}; use read_pages to verify the exact text`,
      summary,
    )
  }
  const targets = input.all === true ? onPage : onPage.slice(0, 1)
  for (const m of targets) deps.addMarkup(type, r.origIdx, m.rects, color)
  deps.gotoPage(r.origIdx + 1)
  return {
    output: `Marked ${targets.length} occurrence(s) on page ${r.origIdx + 1} (unsaved; the user saves with ⌘S)`,
    mutated: true,
    summary,
  }
}

/** #RRGGBB → rgb 0-1; undefined when omitted, null when malformed */
const hexToRgb01 = (raw: unknown): [number, number, number] | undefined | null => {
  if (raw === undefined) return undefined
  const hex = HEX_COLOR.exec(String(raw))
  if (!hex) return null
  const v = parseInt(hex[1]!, 16)
  return [((v >> 16) & 255) / 255, ((v >> 8) & 255) / 255, (v & 255) / 255]
}

const fmtNoteDate = (ms: number | null): string =>
  ms === null ? '' : ` (${new Date(ms).toISOString().slice(0, 10)})`

/** Page text covered by a markup rect, best-effort via the search index */
function textUnderRect(entry: PageEntry, rect: readonly [number, number, number, number]): string {
  const [x1, y1, x2, y2] = rect
  let out = ''
  for (const it of entry.items) {
    const yOverlap = Math.min(y2, it.y + it.h) - Math.max(y1, it.y)
    if (yOverlap < it.h * 0.5) continue
    const lo = Math.max(0, Math.min(1, (x1 - it.x) / (it.w || 1)))
    const hi = Math.max(0, Math.min(1, (x2 - it.x) / (it.w || 1)))
    if (hi <= lo) continue
    const len = it.end - it.start
    out += entry.text.slice(it.start + Math.round(len * lo), it.start + Math.round(len * hi))
  }
  return out.replace(/\s+/g, ' ').trim()
}

async function readAnnotations(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const summary = t('aiToolReadAnnots')
  const pageCount = deps.pageCount()
  const start = input.start === undefined ? 1 : Math.trunc(Number(input.start))
  let end = input.end === undefined ? pageCount : Math.trunc(Number(input.end))
  if (!(start >= 1) || !(end >= start) || start > pageCount)
    return err(`Invalid page range (the document has ${pageCount} pages)`, summary)
  end = Math.min(end, pageCount)
  const indexPromise = deps.searchIndex()
  const index = indexPromise ? await indexPromise : null
  const lines: string[] = []
  let notes = 0
  let marks = 0
  for (let p = start; p <= end; p++) {
    const origIdx = p - 1
    if (deps.isDeleted(origIdx)) continue
    const { threads, markups } = await deps.annotationsOn(origIdx)
    if (threads.length === 0 && markups.length === 0) continue
    lines.push(`[Page ${p}]`)
    for (const root of threads) {
      notes++
      for (const { item, depth } of flattenThread(root)) {
        const head = depth === 0 ? `- Note ${root.key}` : `${'  '.repeat(depth)}- reply ${item.key}`
        const unsaved = item.saved ? '' : ' (unsaved)'
        lines.push(
          `${head}${unsaved} by ${item.author || 'unknown'}${fmtNoteDate(item.timeMs)}: ${JSON.stringify(item.contents)}`,
        )
      }
    }
    const entry = index?.[origIdx]
    for (const m of markups) {
      marks++
      const covered = entry
        ? m.quads
            .map((q) => textUnderRect(entry, quadToRect(q)))
            .filter(Boolean)
            .join(' ')
        : ''
      const quote = covered
        ? ` on ${JSON.stringify(covered.length > 120 ? `${covered.slice(0, 120)}…` : covered)}`
        : ''
      lines.push(`- Markup ${m.key}: ${m.type}${m.saved ? '' : ' (unsaved)'}${quote}`)
    }
    if (lines.join('\n').length > READ_CHUNK_CHARS) {
      lines.push('…output truncated; call read_annotations with a narrower page range for the rest')
      break
    }
  }
  if (notes === 0 && marks === 0)
    return { output: `No notes or markups on pages ${start}-${end}.`, summary }
  return { output: `${notes} note thread(s), ${marks} markup(s):\n${lines.join('\n')}`, summary }
}

async function addNoteTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const summary = t('aiToolAddNote', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const text = String(input.text ?? '').trim()
  if (!text) return err('text must not be empty', summary)
  const color = hexToRgb01(input.color)
  if (color === null) return err(`Invalid color "${String(input.color)}"; use #RRGGBB`, summary)
  const geom = deps.pageGeom(r.origIdx)
  if (!geom) return err('Document not ready', summary)
  let at: [number, number]
  const anchor = String(input.anchor_text ?? '').trim()
  if (anchor) {
    const indexPromise = deps.searchIndex()
    if (!indexPromise) return err('Document not ready', summary)
    const entry = (await indexPromise)[r.origIdx]
    const located = entry ? locateOccurrence(entry, anchor, 1) : null
    if (!located) {
      return err(
        `"${anchor}" not found on page ${r.origIdx + 1}; use read_pages to verify the exact text`,
        summary,
      )
    }
    at = [located.rect[2], located.rect[3]]
  } else if (input.x !== undefined || input.y !== undefined) {
    const tx = Number(input.x ?? 0)
    const ty = Number(input.y ?? 0)
    if (!Number.isFinite(tx) || !Number.isFinite(ty))
      return err('x and y must be numbers (points from the page top-left as displayed)', summary)
    at = viewToPdf(geom, tx, ty)
  } else {
    return err('Position the note with anchor_text or x/y', summary)
  }
  if (signal?.aborted) return err('stopped by the user; nothing was changed', summary)
  const key = deps.addNote(r.origIdx, at, text, color)
  deps.gotoPage(r.origIdx + 1)
  return {
    output: `Added note ${key} on page ${r.origIdx + 1} (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

async function replyNoteTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const summary = t('aiToolReplyNote', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const text = String(input.text ?? '').trim()
  if (!text) return err('text must not be empty', summary)
  const noteId = String(input.note_id ?? '').trim()
  if (!noteId) return err('note_id must not be empty', summary)
  const root = await deps.findNoteRoot(r.origIdx, noteId)
  if (!root) {
    return err(
      `Note "${noteId}" not found on page ${r.origIdx + 1}; call read_annotations to list thread ids`,
      summary,
    )
  }
  if (signal?.aborted) return err('stopped by the user; nothing was changed', summary)
  deps.replyToThread(r.origIdx, root, text)
  deps.gotoPage(r.origIdx + 1)
  return {
    output: `Replied to note ${noteId} on page ${r.origIdx + 1} (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

async function editNoteTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const summary = t('aiToolEditNote', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const page = r.origIdx + 1
  const text = String(input.text ?? '').trim()
  if (!text) return err('text must not be empty; use delete_note to remove a note', summary)
  const noteId = String(input.note_id ?? '').trim()
  if (!noteId) return err('note_id must not be empty', summary)
  const { threads } = await deps.annotationsOn(r.origIdx)
  const item = threads
    .flatMap((root) => flattenThread(root))
    .map(({ item: it }) => it)
    .find((it) => it.key === noteId)
  if (!item) {
    return err(
      `Note ${noteId} not found on page ${page}; call read_annotations for the current note and reply ids`,
      summary,
    )
  }
  if (text === item.contents) return { output: `Note ${noteId} already says that.`, summary }
  if (signal?.aborted) return err('stopped by the user; nothing was changed', summary)
  deps.editNote(r.origIdx, item, text)
  deps.gotoPage(page)
  return {
    output: `Edited note ${noteId} on page ${page} (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

/** nth (1-based, non-overlapping) case-insensitive occurrence of query on a page →
    verbatim text, union rect (PDF space), and line height (≈ font size) */
function locateOccurrence(
  entry: PageEntry,
  query: string,
  occurrence: number,
): { oldText: string; rect: [number, number, number, number]; fontSize: number } | null {
  const q = query.toLowerCase()
  let pos = -1
  let from = 0
  for (let i = 0; i < occurrence; i++) {
    pos = entry.lower.indexOf(q, from)
    if (pos < 0) return null
    from = pos + q.length
  }
  const end = pos + q.length
  let x1 = Infinity
  let y1 = Infinity
  let x2 = -Infinity
  let y2 = -Infinity
  let fontSize = 0
  for (const it of entry.items) {
    if (it.end <= pos || it.start >= end) continue
    const len = it.end - it.start
    const lo = (Math.max(pos, it.start) - it.start) / len
    const hi = (Math.min(end, it.end) - it.start) / len
    x1 = Math.min(x1, it.x + it.w * lo)
    x2 = Math.max(x2, it.x + it.w * hi)
    y1 = Math.min(y1, it.y)
    y2 = Math.max(y2, it.y + it.h)
    fontSize = Math.max(fontSize, it.h)
  }
  if (fontSize <= 0 || x2 - x1 < 0.01) return null
  return { oldText: entry.text.slice(pos, end), rect: [x1, y1, x2, y2], fontSize }
}

const countOccurrences = (entry: PageEntry, query: string): number => {
  const q = query.toLowerCase()
  let n = 0
  for (let from = 0; ; n++) {
    const pos = entry.lower.indexOf(q, from)
    if (pos < 0) return n
    from = pos + q.length
  }
}

const HEX_COLOR = /^#?([0-9a-f]{6})$/i

const MARKUP_TYPES: MarkupType[] = ['highlight', 'underline', 'strikeout']

async function deleteMarkupTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const summary = t('aiToolDeleteMarkup', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const type = input.type === undefined ? undefined : (String(input.type) as MarkupType)
  if (type !== undefined && !MARKUP_TYPES.includes(type)) {
    return err(`Invalid type "${type}"; use one of ${MARKUP_TYPES.join(', ')}`, summary)
  }
  const raw = input.markup_ids
  const ids =
    raw === undefined
      ? undefined
      : Array.isArray(raw)
        ? raw.map(String)
        : typeof raw === 'string'
          ? [raw]
          : null
  if (ids === null) return err('markup_ids must be an array of markup ids', summary)
  const page = r.origIdx + 1
  const { markups } = await deps.annotationsOn(r.origIdx)
  let targets = markups
  if (ids) {
    const missing = ids.filter((id) => !markups.some((m) => m.key === id))
    if (missing.length > 0) {
      return err(
        `No markup ${missing.join(', ')} on page ${page}; call read_annotations for the current ids`,
        summary,
      )
    }
    const want = new Set(ids)
    targets = markups.filter((m) => want.has(m.key))
  }
  if (type) targets = targets.filter((m) => m.type === type)
  if (targets.length === 0)
    return err(`No ${type ? `${type} ` : ''}markups on page ${page}`, summary)
  await deps.deleteMarkups(
    r.origIdx,
    targets.map((m) => m.key),
  )
  deps.gotoPage(page)
  const counts = MARKUP_TYPES.map((k) => [k, targets.filter((m) => m.type === k).length] as const)
    .filter(([, n]) => n > 0)
    .map(([k, n]) => `${n} ${k}`)
    .join(', ')
  return {
    output: `Removed ${counts} markup(s) from page ${page} (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

/** Optional numeric parameter within [min, max]; undefined when absent, string = rejection */
function numberParam(
  input: Record<string, unknown>,
  name: string,
  min: number,
  max: number,
): number | undefined | string {
  if (input[name] === undefined) return undefined
  const v = Number(input[name])
  if (!Number.isFinite(v) || v < min || v > max)
    return `${name} must be a number between ${min} and ${max}`
  return v
}

const hexColorParam = (input: Record<string, unknown>): string | undefined | null => {
  if (input.color === undefined) return undefined
  const hex = HEX_COLOR.exec(String(input.color))
  return hex ? `#${hex[1]!.toLowerCase()}` : null
}

const ROTATE_DELTAS: Record<string, RotateDelta> = { left: -90, right: 90, '180': 180 }

function rotatePageTool(deps: PdfAiDeps, input: Record<string, unknown>): ToolExecution {
  const all = input.all === true
  const summary = all ? t('aiToolRotateAll') : t('aiToolRotate', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const dir = ROTATE_DELTAS[String(input.direction)]
  if (dir === undefined) return err('direction must be left, right, or 180', summary)
  if (all) {
    const pages = deps.pageOrder()
    deps.rotatePages(pages, dir)
    return {
      output: `Rotated all ${pages.length} pages by ${dir}° (unsaved)`,
      mutated: true,
      summary,
    }
  }
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  deps.rotatePages([r.origIdx], dir)
  deps.gotoPage(r.origIdx + 1)
  return { output: `Rotated page ${r.origIdx + 1} by ${dir}° (unsaved)`, mutated: true, summary }
}

const METADATA_FIELDS = ['title', 'author', 'subject', 'keywords'] as const

function describeMetadata(meta: MetadataInput): string {
  return METADATA_FIELDS.map((k) => `${k}: ${meta[k] ? `"${meta[k]}"` : '(empty)'}`).join(', ')
}

function setMetadataTool(deps: PdfAiDeps, input: Record<string, unknown>): ToolExecution {
  const summary = t('aiToolSetMetadata')
  const current = deps.metadata()
  const next: MetadataInput = { ...current }
  let changed = false
  for (const k of METADATA_FIELDS) {
    if (input[k] === undefined || input[k] === null) continue
    const v = String(input[k]).trim()
    if (v === (current[k] ?? '')) continue
    next[k] = v
    changed = true
  }
  if (!changed)
    return { output: `Document properties unchanged: ${describeMetadata(current)}`, summary }
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  deps.setMetadata(next)
  return {
    output: `Document properties updated (unsaved; written on save): ${describeMetadata(next)}`,
    mutated: true,
    summary,
  }
}

function setWatermarkTool(deps: PdfAiDeps, input: Record<string, unknown>): ToolExecution {
  const summary = t('aiToolWatermark')
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const prev = deps.stamps()
  const text = String(input.text ?? '').trim()
  if (!text) {
    if (!prev?.wm) return err('No watermark was added this session', summary)
    deps.setStamps(prev.hf ? { wm: null, hf: prev.hf } : null)
    return { output: 'Removed the watermark (unsaved).', mutated: true, summary }
  }
  const angle = numberParam(input, 'angle', -180, 180)
  const opacity = numberParam(input, 'opacity', 0, 1)
  const size = numberParam(input, 'size', 0.02, 0.5)
  for (const v of [angle, opacity, size]) if (typeof v === 'string') return err(v, summary)
  const color = hexColorParam(input)
  if (color === null) return err(`Invalid color "${String(input.color)}"; use #RRGGBB`, summary)
  const base = prev?.wm ?? DEFAULT_WATERMARK
  deps.setStamps({
    wm: {
      text,
      angle: (angle as number | undefined) ?? base.angle,
      opacity: (opacity as number | undefined) ?? base.opacity,
      color: color ?? base.color,
      sizeRatio: (size as number | undefined) ?? base.sizeRatio,
    },
    hf: prev?.hf ?? null,
  })
  return {
    output: `Watermark "${text}" applied to every page (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

async function deleteNoteTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const summary = t('aiToolDeleteNote', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const noteId = String(input.note_id ?? '').trim()
  if (!noteId) return err('note_id must not be empty', summary)
  const page = r.origIdx + 1
  const root = await deps.findNoteRoot(r.origIdx, noteId)
  if (!root) {
    return err(
      `Note ${noteId} not found on page ${page}; call read_annotations for the current ids`,
      summary,
    )
  }
  deps.deleteNoteThread(r.origIdx, root)
  deps.gotoPage(page)
  const replies = flattenThread(root).length - 1
  const tail = replies > 0 ? ` and its ${replies} repl${replies === 1 ? 'y' : 'ies'}` : ''
  return {
    output: `Deleted note ${noteId}${tail} on page ${page} (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

/** Visible order as original page numbers, elided past 40 entries */
function describeOrder(order: number[]): string {
  const nums = order.map((i) => i + 1)
  const shown = nums.length > 40 ? `${nums.slice(0, 40).join(', ')}, …` : nums.join(', ')
  return `Current order (original page numbers): ${shown}`
}

function movePageTool(deps: PdfAiDeps, input: Record<string, unknown>): ToolExecution {
  const summary = t('aiToolMovePage', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const order = deps.pageOrder()
  const to = Number(input.to)
  if (!Number.isInteger(to) || to < 1 || to > order.length) {
    return err(`to must be a position between 1 and ${order.length}`, summary)
  }
  const from = order.indexOf(r.origIdx)
  if (from < 0) return err(`Page ${r.origIdx + 1} is not in the page order`, summary)
  if (from === to - 1) {
    return { output: `Page ${r.origIdx + 1} is already at position ${to}.`, summary }
  }
  deps.movePage(r.origIdx, to - 1)
  return {
    output: `Moved page ${r.origIdx + 1} to position ${to} (unsaved). ${describeOrder(deps.pageOrder())}`,
    mutated: true,
    summary,
  }
}

const DECLINED_OUTPUT = 'The user declined the operation; nothing was changed.'
const STOPPED_OUTPUT = 'The run was stopped before the user confirmed; nothing was changed.'

/** Honest report for a dismissed picker: the flush before it may already have rewritten the file */
const canceledOutput = (r: FileOpCanceled, what: string): string =>
  r.flushed
    ? `The user canceled the ${what}, so no pages were changed by this operation — but the pending edits had already been saved to the file first: page numbers now reflect the saved file (pending deletions/reorders baked in) and the undo history is cleared. Re-read with search_text/read_pages before any further edit.`
    : `The user canceled the ${what}; nothing was changed.`

const LAYOUT_CHANGED_OUTPUT =
  'The pages changed while the confirmation card was open (a page was deleted, moved, or added), so the operation was not run; re-read the document with search_text/read_pages and ask again.'

const sameList = (a: number[], b: number[]): boolean =>
  a.length === b.length && a.every((v, i) => v === b[i])

/** Model-facing tail of every in-place file operation */
const reloadedNote = (pageCount: number): string =>
  ` The file was saved, rewritten on disk (cannot be undone) and reloaded; it now has ${pageCount} pages. Original page numbers have changed (pending deletions/reorders were baked in), so re-read with search_text/read_pages before any further edit.`

/** Gate shared by the file-level tools: null = confirmed, otherwise the output to return */
async function fileOpGate(
  deps: PdfAiDeps,
  req: FileOpConfirm,
  signal?: AbortSignal,
): Promise<string | null> {
  if (await deps.confirmFileOp(req, signal)) return null
  return signal?.aborted ? STOPPED_OUTPUT : DECLINED_OUTPUT
}

/** Range string of original page numbers → positions in the flushed file (= visible positions) */
function resolveVisibleRange(
  deps: PdfAiDeps,
  raw: unknown,
): { vis: number[]; map: number[]; label: string } | { bad: string } {
  const text = String(raw ?? '').trim()
  const order = deps.pageOrder()
  if (text.toLowerCase() === 'all') {
    const vis = order.map((_, i) => i)
    return { vis, map: [...order], label: order.length > 1 ? `1-${order.length}` : '1' }
  }
  const pages = parsePageRanges(text, deps.pageCount())
  if (!pages) {
    return {
      bad: `pages must be a range string like "1-3,5" within 1-${deps.pageCount()}${text ? '' : ' (got nothing)'}`,
    }
  }
  const vis: number[] = []
  for (const n of pages) {
    const at = order.indexOf(n - 1)
    if (at < 0) return { bad: `Page ${n} has been deleted (unsaved)` }
    vis.push(at)
  }
  return { vis: [...vis].sort((a, b) => a - b), map: vis, label: text.replace(/\s+/g, '') }
}

/** The card describes a page→position mapping, but the document stays editable while it is
    open; the mapping is taken again after Confirm and must match page for page */
const rangeStillValid = (deps: PdfAiDeps, raw: unknown, map: number[]): boolean => {
  const again = resolveVisibleRange(deps, raw)
  return 'map' in again && sameList(again.map, map)
}

const fmtPath = (r: { savedPath: string }): string =>
  `saved at ${r.savedPath} and opened in a new tab`

async function insertBlankPageTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const after = Number(input.after_page)
  const resolveAfterVis = (): number | { bad: string } => {
    if (after === 0) return -1
    const r = resolvePage(deps, input.after_page)
    if ('bad' in r) return r
    const vis = deps.pageOrder().indexOf(r.origIdx)
    return vis < 0 ? { bad: `Page ${after} is not in the page order` } : vis
  }
  const afterVis = resolveAfterVis()
  if (typeof afterVis !== 'number')
    return err(afterVis.bad, t('aiToolInsertBlankPage', { pos: '?' }))
  const pos = afterVis + 2
  const summary = t('aiToolInsertBlankPage', { pos })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const denied = await fileOpGate(deps, { summary }, signal)
  if (denied) return { output: denied, summary }
  if (resolveAfterVis() !== afterVis) return { output: LAYOUT_CHANGED_OUTPUT, summary }
  const r = await deps.insertBlankPage(afterVis)
  if (!r.ok) return err(r.error, summary)
  if ('canceled' in r) return { output: canceledOutput(r, 'operation'), summary }
  return {
    output: `Inserted a blank page as page ${pos}; every page from there on moved down by one.${reloadedNote(r.pageCount)}`,
    mutated: true,
    summary,
  }
}

async function setPageSizeTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  let size: { w: number; h: number }
  let label: string
  if (input.preset !== undefined) {
    const wanted = String(input.preset).toLowerCase()
    const preset = PAPER_SIZES.find((p) => p.label.toLowerCase() === wanted)
    if (!preset) {
      return err(
        `preset must be one of ${PAPER_SIZES.map((p) => p.label).join('/')}`,
        t('aiToolSetPageSize', { size: String(input.preset) }),
      )
    }
    size = preset
    label = `${preset.label} (${preset.w} × ${preset.h} pt)`
  } else {
    const unknown = t('aiToolSetPageSize', { size: '?' })
    const w = numberParam(input, 'width', 72, 14_400)
    if (typeof w === 'string') return err(w, unknown)
    const h = numberParam(input, 'height', 72, 14_400)
    if (typeof h === 'string') return err(h, unknown)
    if (w === undefined || h === undefined) {
      return err('pass a preset or both width and height in points', unknown)
    }
    size = { w, h }
    label = `${Math.round(w)} × ${Math.round(h)} pt`
  }
  const summary = t('aiToolSetPageSize', { size: label })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const denied = await fileOpGate(deps, { summary }, signal)
  if (denied) return { output: denied, summary }
  const r = await deps.setPageSize(size.w, size.h)
  if (!r.ok) return err(r.error, summary)
  if ('canceled' in r) return { output: canceledOutput(r, 'operation'), summary }
  return {
    output: `Resized all pages to ${label}.${reloadedNote(r.pageCount)}`,
    mutated: true,
    summary,
  }
}

async function cropPagesTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const fallback = t('aiToolCropPages', { pages: '?', w: '?', h: '?' })
  const edges = ['left', 'top', 'right', 'bottom'].map((k) => {
    const v = input[k] === undefined ? `${k} is required` : numberParam(input, k, 0, 1)
    return v
  })
  const bad = edges.find((v): v is string => typeof v === 'string')
  if (bad) return err(bad, fallback)
  const [l, tp, rt, b] = edges as [number, number, number, number]
  if (l >= rt || tp >= b)
    return err('left must be less than right and top less than bottom', fallback)
  if (l <= 0 && tp <= 0 && rt >= 1 && b >= 1) {
    return err('left 0, top 0, right 1, bottom 1 keeps the whole page; nothing to crop', fallback)
  }
  const range = resolveVisibleRange(deps, input.pages)
  if ('bad' in range) return err(range.bad, fallback)
  const pct = (v: number) => Math.round(v * 100)
  const summary = t('aiToolCropPages', { pages: range.label, w: pct(rt - l), h: pct(b - tp) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const denied = await fileOpGate(deps, { summary }, signal)
  if (denied) return { output: denied, summary }
  if (!rangeStillValid(deps, input.pages, range.map))
    return { output: LAYOUT_CHANGED_OUTPUT, summary }
  const r = await deps.cropPages(range.vis, { l, t: tp, r: rt, b })
  if (!r.ok) return err(r.error, summary)
  if ('canceled' in r) return { output: canceledOutput(r, 'operation'), summary }
  return {
    output: `Cropped pages ${range.label} to the area left ${pct(l)}%, top ${pct(tp)}%, right ${pct(rt)}%, bottom ${pct(b)}% of the page.${reloadedNote(r.pageCount)}`,
    mutated: true,
    summary,
  }
}

async function extractPagesTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const range = resolveVisibleRange(deps, input.pages)
  if ('bad' in range) return err(range.bad, t('aiToolExtractPages', { pages: '?' }))
  const summary = t('aiToolExtractPages', { pages: range.label })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const denied = await fileOpGate(deps, { summary, detail: t('aiFileOpNewFile') }, signal)
  if (denied) return { output: denied, summary }
  if (!rangeStillValid(deps, input.pages, range.map))
    return { output: LAYOUT_CHANGED_OUTPUT, summary }
  const r = await deps.extractPages(range.vis)
  if (!r.ok) return err(r.error, summary)
  if ('canceled' in r) return { output: canceledOutput(r, 'operation'), summary }
  return {
    output: `Extracted pages ${range.label} into a new PDF ${fmtPath(r)}. The current document was saved and keeps all its pages.`,
    mutated: true,
    summary,
  }
}

async function splitPdfTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const n = Number(input.pages_per_file)
  const total = deps.pageOrder().length
  const summary = t('aiToolSplitPdf', { n: Number.isFinite(n) ? n : '?' })
  if (!Number.isInteger(n) || n < 1 || n >= total) {
    return err(
      `pages_per_file must be an integer between 1 and ${total - 1} (the document has ${total} pages)`,
      summary,
    )
  }
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const denied = await fileOpGate(deps, { summary, detail: t('aiFileOpPickFolder') }, signal)
  if (denied) return { output: denied, summary }
  if (deps.pageOrder().length !== total) return { output: LAYOUT_CHANGED_OUTPUT, summary }
  const r = await deps.splitPdf(n)
  if (!r.ok) return err(r.error, summary)
  if ('canceled' in r) return { output: canceledOutput(r, 'folder picker'), summary }
  return {
    output: `Wrote ${r.count} PDF files of up to ${n} pages each into ${r.savedDir}. The current document was saved and keeps all its pages.`,
    mutated: true,
    summary,
  }
}

async function splitPagesTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const n = Number(input.per_page)
  const summary = t('aiToolSplitPages', { n: Number.isFinite(n) ? n : '?' })
  if (n !== 2 && n !== 4 && n !== 9) return err('per_page must be 2, 4, or 9', summary)
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const denied = await fileOpGate(deps, { summary, detail: t('aiFileOpNewFile') }, signal)
  if (denied) return { output: denied, summary }
  const r = await deps.splitPages(n)
  if (!r.ok) return err(r.error, summary)
  if ('canceled' in r) return { output: canceledOutput(r, 'operation'), summary }
  return {
    output: `Split every page into ${n} pages; the new PDF is ${fmtPath(r)}. The current document was saved and is unchanged.`,
    mutated: true,
    summary,
  }
}

async function mergePagesTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const n = Number(input.per_sheet)
  const summary = t('aiToolMergePages', { n: Number.isFinite(n) ? n : '?' })
  if (!Number.isInteger(n) || n < 2 || n > 16)
    return err('per_sheet must be an integer between 2 and 16', summary)
  const direction = input.direction === undefined ? 'vertical' : String(input.direction)
  if (direction !== 'horizontal' && direction !== 'vertical') {
    return err('direction must be horizontal or vertical', summary)
  }
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const denied = await fileOpGate(deps, { summary, detail: t('aiFileOpNewFile') }, signal)
  if (denied) return { output: denied, summary }
  const r = await deps.mergePages(n, direction, input.separator === true)
  if (!r.ok) return err(r.error, summary)
  if ('canceled' in r) return { output: canceledOutput(r, 'operation'), summary }
  return {
    output: `Placed ${n} pages per sheet (${direction} order); the new PDF is ${fmtPath(r)}. The current document was saved and is unchanged.`,
    mutated: true,
    summary,
  }
}

async function replacePagesTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const range = resolveVisibleRange(deps, input.pages)
  if ('bad' in range) return err(range.bad, t('aiToolReplacePages', { pages: '?' }))
  const summary = t('aiToolReplacePages', { pages: range.label })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const denied = await fileOpGate(deps, { summary, detail: t('aiFileOpPickFile') }, signal)
  if (denied) return { output: denied, summary }
  if (!rangeStillValid(deps, input.pages, range.map))
    return { output: LAYOUT_CHANGED_OUTPUT, summary }
  const r = await deps.replacePages(range.vis)
  if (!r.ok) return err(r.error, summary)
  if ('canceled' in r) return { output: canceledOutput(r, 'file picker'), summary }
  return {
    output: `Replaced pages ${range.label} with the pages of the chosen PDF.${reloadedNote(r.pageCount)}`,
    mutated: true,
    summary,
  }
}

function setHeaderFooterTool(deps: PdfAiDeps, input: Record<string, unknown>): ToolExecution {
  const summary = t('aiToolHeaderFooter')
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const prev = deps.stamps()
  const slot = (name: string) => String(input[name] ?? '').trim()
  const slots = {
    headerLeft: slot('header_left'),
    headerCenter: slot('header_center'),
    headerRight: slot('header_right'),
    footerLeft: slot('footer_left'),
    footerCenter: slot('footer_center'),
    footerRight: slot('footer_right'),
  }
  const pageNumber = input.page_number === true
  if (!pageNumber && Object.values(slots).every((v) => !v)) {
    if (!prev?.hf) return err('No header/footer was added this session', summary)
    deps.setStamps(prev.wm ? { wm: prev.wm, hf: null } : null)
    return { output: 'Removed the header/footer (unsaved).', mutated: true, summary }
  }
  const startAt = numberParam(input, 'start_at', 0, 100000)
  const fontSize = numberParam(input, 'font_size', 4, 72)
  for (const v of [startAt, fontSize]) if (typeof v === 'string') return err(v, summary)
  const color = hexColorParam(input)
  if (color === null) return err(`Invalid color "${String(input.color)}"; use #RRGGBB`, summary)
  deps.setStamps({
    wm: prev?.wm ?? null,
    hf: {
      ...slots,
      pageNumber,
      startAt: Math.trunc((startAt as number | undefined) ?? DEFAULT_HEADER_FOOTER.startAt),
      fontSize: (fontSize as number | undefined) ?? DEFAULT_HEADER_FOOTER.fontSize,
      color: color ?? DEFAULT_HEADER_FOOTER.color,
    },
  })
  const filled = Object.entries(slots)
    .filter(([, v]) => v)
    .map(([k, v]) => `${k}=${JSON.stringify(v)}`)
  if (pageNumber) filled.push('page numbers in the footer center')
  return {
    output: `Header/footer applied to every page: ${filled.join(', ')} (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

async function editText(deps: PdfAiDeps, input: Record<string, unknown>): Promise<ToolExecution> {
  const summary = t('aiToolEditText', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const oldText = String(input.old_text ?? '').trim()
  const newText = String(input.new_text ?? '')
    .replace(/\r\n/g, '\n')
    .trim()
  if (!oldText) return err('old_text must not be empty', summary)
  if (input.align !== undefined) {
    return err('edit_text cannot re-align a run; use edit_block with align', summary)
  }
  const occurrence = Math.max(1, Math.trunc(Number(input.occurrence ?? 1)) || 1)
  let newColor: [number, number, number] | undefined
  if (input.color !== undefined) {
    const hex = HEX_COLOR.exec(String(input.color))
    if (!hex) return err(`Invalid color "${String(input.color)}"; use #RRGGBB`, summary)
    const v = parseInt(hex[1]!, 16)
    newColor = [(v >> 16) & 255, (v >> 8) & 255, v & 255]
  }
  const newFontSize = input.font_size === undefined ? undefined : Number(input.font_size)
  if (newFontSize !== undefined && !(newFontSize > 0)) {
    return err('font_size must be a positive number', summary)
  }
  const newFont = input.font === undefined ? undefined : String(input.font)
  if (newFont !== undefined && !deps.editFonts().includes(newFont)) {
    const avail = deps.editFonts()
    return err(
      avail.length > 0
        ? `Font "${newFont}" is not available; choose one of: ${avail.join(', ')}`
        : 'No selectable fonts are available on this machine; omit the font parameter',
      summary,
    )
  }
  const indexPromise = deps.searchIndex()
  if (!indexPromise) return err('Document not ready', summary)
  const index = await indexPromise
  const entry = index[r.origIdx]
  const located = entry ? locateOccurrence(entry, oldText, occurrence) : null
  if (!entry || !located) {
    return err(
      `Occurrence ${occurrence} of "${oldText}" not found on page ${r.origIdx + 1}; use read_pages to verify the exact text`,
      summary,
    )
  }
  const reason = await deps.editText({
    pageIndex: r.origIdx,
    rect: located.rect,
    oldText: located.oldText,
    newText,
    fontSize: located.fontSize,
    newFontSize,
    newColor,
    newFont,
    newBold: input.bold === true ? true : undefined,
    newItalic: input.italic === true ? true : undefined,
  })
  if (reason) return err(`The edit could not be applied: ${reason}`, summary)
  deps.gotoPage(r.origIdx + 1)
  const total = countOccurrences(entry, oldText)
  const note =
    total > 1 && input.occurrence === undefined
      ? ` Note: the page has ${total} occurrences of this text and only the first was edited; pass occurrence to target another.`
      : ''
  const verb = newText ? 'Replaced' : 'Deleted'
  return {
    output: `${verb} occurrence ${occurrence} of "${oldText}" on page ${r.origIdx + 1} (unsaved; the user saves with ⌘S).${note}`,
    mutated: true,
    summary,
  }
}

/** The one clustered paragraph on a page containing `anchor` (whitespace-insensitive) */
async function locateBlock(
  deps: PdfAiDeps,
  origIdx: number,
  anchor: string,
): Promise<{ block: TextBlock } | { bad: string }> {
  const indexPromise = deps.searchIndex()
  if (!indexPromise) return { bad: 'Document not ready' }
  const entry = (await indexPromise)[origIdx]
  if (!entry) return { bad: `Page ${origIdx + 1} has no extractable text` }
  const squash = (s: string) => s.replace(/\s+/g, '')
  const key = squash(anchor)
  const hits = groupPageBlocks(entry).filter((b) =>
    squash(b.lines.map((l) => l.text).join('')).includes(key),
  )
  if (hits.length === 0) {
    return {
      bad: `No paragraph on page ${origIdx + 1} contains "${anchor}"; use read_pages to verify the exact text`,
    }
  }
  if (hits.length > 1) {
    return {
      bad: `${hits.length} paragraphs on page ${origIdx + 1} contain "${anchor}"; pass a longer, unique fragment`,
    }
  }
  return { block: hits[0]! }
}

/** Rewrite one clustered paragraph, reflowed within the block's original width */
async function editBlock(deps: PdfAiDeps, input: Record<string, unknown>): Promise<ToolExecution> {
  const summary = t('aiToolEditText', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const anchor = String(input.paragraph_text ?? '').trim()
  const newText = String(input.new_text ?? '')
    .replace(/\r\n/g, '\n')
    .trim()
  if (!anchor) return err('paragraph_text must not be empty', summary)
  let newColor: [number, number, number] | undefined
  if (input.color !== undefined) {
    const hex = HEX_COLOR.exec(String(input.color))
    if (!hex) return err(`Invalid color "${String(input.color)}"; use #RRGGBB`, summary)
    const v = parseInt(hex[1]!, 16)
    newColor = [(v >> 16) & 255, (v >> 8) & 255, v & 255]
  }
  const newFontSize = input.font_size === undefined ? undefined : Number(input.font_size)
  if (newFontSize !== undefined && !(newFontSize > 0)) {
    return err('font_size must be a positive number', summary)
  }
  const newFont = input.font === undefined ? undefined : String(input.font)
  if (newFont !== undefined && !deps.editFonts().includes(newFont)) {
    const avail = deps.editFonts()
    return err(
      avail.length > 0
        ? `Font "${newFont}" is not available; choose one of: ${avail.join(', ')}`
        : 'No selectable fonts are available on this machine; omit the font parameter',
      summary,
    )
  }
  const located = await locateBlock(deps, r.origIdx, anchor)
  if ('bad' in located) return err(located.bad, summary)
  const { block } = located
  const size = newFontSize ?? block.fontSize
  const widthPt = block.rect[2] - block.rect[0]
  // Measure with the face that will actually be embedded: an explicit font choice
  // wraps in that face's CSS family (same as the UI commit path), else the body font
  const cssFamily =
    (newFont ? EDIT_FONTS.find((f) => f.id === newFont)?.css : undefined) ??
    getComputedStyle(document.body).fontFamily
  const bold = input.bold === true ? true : undefined
  const italic = input.italic === true ? true : undefined
  const cssStyle = `${italic ? 'italic ' : ''}${bold ? 'bold' : ''}`.trim()
  const align = input.align === undefined ? block.align : String(input.align)
  if (align !== 'left' && align !== 'center' && align !== 'right') {
    return err('align must be one of left, center, right', summary)
  }
  const lines = newText
    .split('\n')
    .flatMap((p) => (p.trim() ? wrapText(p, widthPt, size, cssFamily, cssStyle) : []))
  const reason = await deps.editText({
    pageIndex: r.origIdx,
    rect: block.rect,
    oldText: joinBlockLines(block.lines.map((l) => l.text)),
    newText: lines.join('\n'),
    fontSize: block.fontSize,
    newFontSize,
    newColor,
    newFont,
    newBold: bold,
    newItalic: italic,
    origin: [block.rect[0], block.lines[0]!.y],
    lineLeading: block.lineHeight * (size / block.fontSize),
    lineXOffsets:
      align === 'left'
        ? undefined
        : lines.map((l) => {
            const slack = widthPt - measurePt(l, size, cssFamily, cssStyle)
            return Math.max(0, align === 'center' ? slack / 2 : slack)
          }),
    align: align === 'left' ? undefined : align,
    blockSource: newText,
  })
  if (reason) return err(`The edit could not be applied: ${reason}`, summary)
  deps.gotoPage(r.origIdx + 1)
  return {
    output: newText
      ? `Replaced the paragraph containing "${anchor}" on page ${r.origIdx + 1} with ${lines.length} reflowed line(s) (unsaved; the user saves with ⌘S).`
      : `Deleted the paragraph containing "${anchor}" on page ${r.origIdx + 1}; its area stays blank (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

async function addFormMarkTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const summary = t('aiToolAddFormMark', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const geom = deps.pageGeom(r.origIdx)
  if (!geom) return err('Document not ready', summary)
  const kind = String(input.kind ?? '')
  if (kind !== 'check' && kind !== 'cross') return err("kind must be 'check' or 'cross'", summary)
  const size = input.size === undefined ? STATIC_FORM_MARK_SIZE : Number(input.size)
  if (!(size > 0)) return err('size must be a positive number', summary)

  const disp = geomDispSize(geom)
  let tx: number
  let ty: number
  const anchor = String(input.anchor_text ?? '').trim()
  if (anchor) {
    const indexPromise = deps.searchIndex()
    if (!indexPromise) return err('Document not ready', summary)
    const entry = (await indexPromise)[r.origIdx]
    const located = entry ? locateOccurrence(entry, anchor, 1) : null
    if (!located) {
      return err(
        `"${anchor}" not found on page ${r.origIdx + 1}; use read_pages to verify the exact text`,
        summary,
      )
    }
    const a = dispBox(geom, located.rect)
    const placement = String(input.placement ?? 'right')
    if (placement === 'above') {
      tx = a.left + a.width / 2 - size / 2
      ty = a.top - TEXT_GAP_PT - size
    } else if (placement === 'below') {
      tx = a.left + a.width / 2 - size / 2
      ty = a.top + a.height + TEXT_GAP_PT
    } else if (placement === 'left') {
      tx = a.left - TEXT_GAP_PT - size
      ty = a.top + a.height / 2 - size / 2
    } else {
      tx = a.left + a.width + TEXT_GAP_PT
      ty = a.top + a.height / 2 - size / 2
    }
  } else if (input.x !== undefined || input.y !== undefined) {
    tx = Number(input.x ?? 0)
    ty = Number(input.y ?? 0)
    if (!Number.isFinite(tx) || !Number.isFinite(ty))
      return err('x and y must be numbers (points from the page top-left as displayed)', summary)
  } else {
    return err('Position the mark with anchor_text or x/y', summary)
  }
  tx = Math.min(Math.max(tx, 0), Math.max(disp.width - size, 0))
  ty = Math.min(Math.max(ty, 0), Math.max(disp.height - size, 0))

  if (signal?.aborted) return err('stopped by the user; nothing was changed', summary)
  deps.addFormMark(r.origIdx, kind, dispToPdfRect(geom, tx, ty, size, size))
  deps.gotoPage(r.origIdx + 1)
  return {
    output: `Placed a ${kind === 'check' ? 'check mark' : 'cross'} on page ${r.origIdx + 1}: ${fmt(size)} pt at x=${fmt(tx)}, y=${fmt(ty)} (unsaved; the user can drag/resize it, undo with ⌘Z, save with ⌘S).`,
    mutated: true,
    summary,
  }
}

/** Translate one clustered paragraph by a display-space offset (the AI counterpart of dragging a block) */
async function moveTextBlockTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const summary = t('aiToolMoveTextBlock', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const anchor = String(input.paragraph_text ?? '').trim()
  if (!anchor) return err('paragraph_text must not be empty', summary)
  const dx = Number(input.dx ?? 0)
  const dy = Number(input.dy ?? 0)
  if (!Number.isFinite(dx) || !Number.isFinite(dy)) return err('dx and dy must be numbers', summary)
  if (dx === 0 && dy === 0) return err('dx and dy are both 0; nothing to move', summary)
  const geom = deps.pageGeom(r.origIdx)
  if (!geom) return err('Document not ready', summary)
  const located = await locateBlock(deps, r.origIdx, anchor)
  if ('bad' in located) return err(located.bad, summary)
  const { block } = located
  const [ax, ay] = viewToPdf(geom, 0, 0)
  const [bx, by] = viewToPdf(geom, dx, dy)
  const delta: [number, number] = [bx - ax, by - ay]
  const res = await deps.moveTextBlock(r.origIdx, block, delta)
  if ('reason' in res) return err(`The paragraph could not be moved: ${res.reason}`, summary)
  deps.gotoPage(r.origIdx + 1)
  const [mx, my] = res.moveBy
  const moved = dispBox(geom, [
    block.rect[0] + mx,
    block.rect[1] + my,
    block.rect[2] + mx,
    block.rect[3] + my,
  ])
  return {
    output: `Moved the paragraph containing "${anchor}" on page ${r.origIdx + 1} by dx ${fmt(dx)}, dy ${fmt(dy)} pt; its top-left is now at x ${fmt(moved.left)}, y ${fmt(moved.top)} from the page top-left as displayed (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

/** Insert a brand-new text block (the blank-page counterpart of edit_text/edit_block) */
async function insertTextTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const summary = t('aiToolInsertText', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const geom = deps.pageGeom(r.origIdx)
  if (!geom) return err('Document not ready', summary)
  const text = String(input.text ?? '')
    .replace(/\r\n/g, '\n')
    .trim()
  if (!text) return err('text must not be empty', summary)

  const fontSize = input.font_size === undefined ? 14 : Number(input.font_size)
  if (!(fontSize > 0)) return err('font_size must be a positive number', summary)
  let color: [number, number, number] = [0, 0, 0]
  if (input.color !== undefined) {
    const hex = HEX_COLOR.exec(String(input.color))
    if (!hex) return err(`Invalid color "${String(input.color)}"; use #RRGGBB`, summary)
    const v = parseInt(hex[1]!, 16)
    color = [(v >> 16) & 255, (v >> 8) & 255, v & 255]
  }
  const font = input.font === undefined ? undefined : String(input.font)
  if (font !== undefined && !deps.editFonts().includes(font)) {
    const avail = deps.editFonts()
    return err(
      avail.length > 0
        ? `Font "${font}" is not available; choose one of: ${avail.join(', ')}`
        : 'No selectable fonts are available on this machine; omit the font parameter',
      summary,
    )
  }
  const align = input.align === undefined ? 'left' : String(input.align)
  if (align !== 'left' && align !== 'center' && align !== 'right')
    return err("align must be 'left', 'center' or 'right'", summary)
  const maxWidth = input.max_width === undefined ? undefined : Number(input.max_width)
  if (maxWidth !== undefined && !(maxWidth > 0))
    return err('max_width must be a positive number', summary)

  // measure/wrap with the face that will actually be embedded (same as edit_block)
  const bold = input.bold === true ? true : undefined
  const italic = input.italic === true ? true : undefined
  const { cssFamily, cssStyle } = faceCss(font, bold, italic)
  const lines = maxWidth
    ? text
        .split('\n')
        .flatMap((p) => (p.trim() ? wrapText(p, maxWidth, fontSize, cssFamily, cssStyle) : ['']))
    : text.split('\n')
  const lineWidths = lines.map((l) => measurePt(l, fontSize, cssFamily, cssStyle))
  const blockW = maxWidth ?? Math.max(...lineWidths, 1)
  const lineLeading = fontSize * 1.2
  const blockH = lineLeading * (lines.length - 1) + fontSize * 1.2

  // display-space position (top-left origin, points at scale 1), clamped to the page
  const size = geomDispSize(geom)
  let tx: number
  let ty: number
  const anchor = String(input.anchor_text ?? '').trim()
  if (anchor) {
    const indexPromise = deps.searchIndex()
    if (!indexPromise) return err('Document not ready', summary)
    const entry = (await indexPromise)[r.origIdx]
    const located = entry ? locateOccurrence(entry, anchor, 1) : null
    if (!located) {
      return err(
        `"${anchor}" not found on page ${r.origIdx + 1}; use read_pages to verify the exact text`,
        summary,
      )
    }
    const a = dispBox(geom, located.rect)
    const placement = String(input.placement ?? 'right')
    if (placement === 'above') {
      tx = a.left
      ty = a.top - TEXT_GAP_PT - blockH
    } else if (placement === 'below') {
      tx = a.left
      ty = a.top + a.height + TEXT_GAP_PT
    } else {
      // right: first baseline sits on the anchor's baseline (form-fill style)
      tx = a.left + a.width + TEXT_GAP_PT
      ty = a.top + a.height - fontSize * 1.1
    }
  } else {
    tx = input.x === undefined ? (size.width - blockW) / 2 : Number(input.x)
    ty = input.y === undefined ? 72 : Number(input.y)
    if (!Number.isFinite(tx) || !Number.isFinite(ty))
      return err('x and y must be numbers (points from the page top-left as displayed)', summary)
  }
  tx = Math.min(Math.max(tx, 0), Math.max(size.width - blockW, 0))
  ty = Math.min(Math.max(ty, 0), Math.max(size.height - blockH, 0))

  if (signal?.aborted) return err('stopped by the user; nothing was changed', summary)
  // origin.x is the alignment anchor, as the preview and the re-edit dialog read it:
  // center/right lines hang left of it by their own width
  const anchorX = align === 'center' ? tx + blockW / 2 : align === 'right' ? tx + blockW : tx
  const record: TextInsertInput = {
    pageIndex: r.origIdx,
    // TextInsertInput.origin is the first-line baseline in PDF user space
    origin: viewToPdf(geom, anchorX, ty + fontSize),
    text: lines.join('\n'),
    fontSize,
    color,
    font,
    bold,
    italic,
    lineLeading,
    lineXOffsets: align === 'left' ? undefined : alignOffsets(lineWidths, align),
    align: align === 'left' ? undefined : align,
    rotate: ((geom.rot % 360) + 360) % 360,
  }
  const id = deps.insertText(record)
  deps.gotoPage(r.origIdx + 1)
  // report the position list_inserted_text/move_inserted_text will use, not the wrap box
  const [px, py] = insertDispTopLeft(geom, { id, input: record })
  return {
    output: `Inserted ${lines.length} line(s) of text on page ${r.origIdx + 1} at x=${fmt(px)}, y=${fmt(py)}, ${fontSize} pt as T${id} (unsaved; the user can drag/edit it, undo with ⌘Z, save with ⌘S; edit_inserted_text / move_inserted_text / delete_inserted_text take this id until saved).`,
    mutated: true,
    summary,
  }
}

// ── Pending inserted text lifecycle ─────────────────────────────────

/** CSS face insert_text measures and previews with; the engine embeds the matching file */
function faceCss(font: string | undefined, bold: boolean | undefined, italic: boolean | undefined) {
  return {
    cssFamily:
      (font ? EDIT_FONTS.find((f) => f.id === font)?.css : undefined) ??
      getComputedStyle(document.body).fontFamily,
    cssStyle: `${italic ? 'italic ' : ''}${bold ? 'bold' : ''}`.trim(),
  }
}

/** origin.x is the align anchor: center/right lines hang left of it by their own width */
const alignOffsets = (widths: number[], align: TextAlign): number[] =>
  widths.map((w) => (align === 'left' ? 0 : align === 'center' ? -w / 2 : -w))

/** Offsets for a re-edited block, measured with the record's own face like insert_text did */
function remeasureOffsets(
  input: TextInsertInput,
  text: string,
  fontSize: number,
  align: TextAlign,
) {
  const { cssFamily, cssStyle } = faceCss(input.font, input.bold, input.italic)
  return alignOffsets(
    text.split('\n').map((l) => measurePt(l, fontSize, cssFamily, cssStyle)),
    align,
  )
}

const INSERT_ID_HINT =
  'call list_inserted_text for the current ids (only unsaved inserts qualify; saved text is page content — use edit_text)'

function resolveInsert(deps: PdfAiDeps, raw: unknown): { ins: LocalTextInsert } | { bad: string } {
  const key = String(raw ?? '').trim()
  if (!key) return { bad: `id must not be empty; ${INSERT_ID_HINT}` }
  const id = key.startsWith('T') ? key.slice(1) : key
  const ins = deps.textInserts().find((i) => i.id === id)
  return ins ? { ins } : { bad: `No pending inserted text "${key}"; ${INSERT_ID_HINT}` }
}

/** Displayed top-left (points, scale 1): origin.x is the align anchor, so center/right
    blocks extend left of it by their most negative line offset */
function insertDispTopLeft(geom: PageGeom, ins: LocalTextInsert): [number, number] {
  const [vx, vy] = pdfToView(geom, ins.input.origin[0], ins.input.origin[1])
  return [vx + Math.min(0, ...(ins.input.lineXOffsets ?? [])), vy - ins.input.fontSize]
}

const isAlign = (v: unknown): v is TextAlign => v === 'left' || v === 'center' || v === 'right'

function listInsertedTextTool(deps: PdfAiDeps, input: Record<string, unknown>): ToolExecution {
  const summary = t('aiToolListInsertedText')
  let inserts = deps.textInserts()
  let where = ''
  if (input.page !== undefined) {
    const r = resolvePage(deps, input.page)
    if ('bad' in r) return err(r.bad, summary)
    inserts = inserts.filter((i) => i.input.pageIndex === r.origIdx)
    where = ` on page ${r.origIdx + 1}`
  }
  if (inserts.length === 0) {
    return {
      output: `No pending inserted text blocks${where}. Text that was already saved is page content: use search_text and edit_text for it.`,
      summary,
    }
  }
  const lines = inserts.map((ins) => {
    const geom = deps.pageGeom(ins.input.pageIndex)
    const pos = geom ? insertDispTopLeft(geom, ins) : null
    const flat = ins.input.text.replace(/\n/g, ' ⏎ ')
    const preview = flat.length > 80 ? `${flat.slice(0, 80)}…` : flat
    return `T${ins.id}: page ${ins.input.pageIndex + 1}, ${
      pos ? `x=${fmt(pos[0])}, y=${fmt(pos[1])}` : 'position unavailable'
    }, ${ins.input.fontSize} pt, ${ins.input.align ?? 'left'}, ${
      ins.input.text.split('\n').length
    } line(s): "${preview}"`
  })
  return {
    output: `${inserts.length} pending inserted text block(s)${where} (unsaved; x/y = top-left in points as displayed):\n${lines.join('\n')}`,
    summary,
  }
}

function editInsertedTextTool(deps: PdfAiDeps, input: Record<string, unknown>): ToolExecution {
  const summary = t('aiToolEditInsertedText')
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolveInsert(deps, input.id)
  if ('bad' in r) return err(r.bad, summary)
  const cur = r.ins.input
  if (
    input.text === undefined &&
    input.font_size === undefined &&
    input.color === undefined &&
    input.align === undefined
  ) {
    return err('Pass at least one of text, font_size, color, align', summary)
  }
  const text =
    input.text === undefined ? cur.text : String(input.text).replace(/\r\n/g, '\n').trim()
  if (!text)
    return err('text must not be empty; use delete_inserted_text to remove the block', summary)
  const fontSize = input.font_size === undefined ? cur.fontSize : Number(input.font_size)
  if (!(fontSize > 0)) return err('font_size must be a positive number', summary)
  let color = cur.color
  if (input.color !== undefined) {
    const hex = HEX_COLOR.exec(String(input.color))
    if (!hex) return err(`Invalid color "${String(input.color)}"; use #RRGGBB`, summary)
    const v = parseInt(hex[1]!, 16)
    color = [(v >> 16) & 255, (v >> 8) & 255, v & 255]
  }
  const align = input.align === undefined ? (cur.align ?? 'left') : input.align
  if (!isAlign(align)) return err("align must be 'left', 'center' or 'right'", summary)
  const edit: TextInsertEdit = {}
  if (input.text !== undefined) edit.text = text
  if (input.font_size !== undefined) {
    edit.fontSize = fontSize
    edit.lineLeading = fontSize * 1.2
  }
  if (input.color !== undefined) edit.color = color
  if (input.align !== undefined) edit.align = align
  if (input.text !== undefined || input.font_size !== undefined || input.align !== undefined) {
    edit.lineXOffsets = remeasureOffsets(cur, text, fontSize, align)
  }
  deps.updateTextInsert(r.ins.id, edit)
  deps.gotoPage(cur.pageIndex + 1)
  return {
    output: `Updated inserted text T${r.ins.id} on page ${cur.pageIndex + 1}: ${text.split('\n').length} line(s), ${fontSize} pt, ${align} (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

function moveInsertedTextTool(deps: PdfAiDeps, input: Record<string, unknown>): ToolExecution {
  const summary = t('aiToolMoveInsertedText')
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolveInsert(deps, input.id)
  if ('bad' in r) return err(r.bad, summary)
  const { pageIndex, origin } = r.ins.input
  const geom = deps.pageGeom(pageIndex)
  if (!geom) return err('Document not ready', summary)
  const has = (k: string) => input[k] !== undefined
  const absolute = has('x') || has('y')
  const relative = has('dx') || has('dy')
  if (!absolute && !relative) {
    return err('Pass x/y (new top-left) or dx/dy (offset), in points as displayed', summary)
  }
  if (absolute && relative) return err('Pass either x/y or dx/dy, not both', summary)
  const [cx, cy] = insertDispTopLeft(geom, r.ins)
  let nx = has('x') ? Number(input.x) : cx + (has('dx') ? Number(input.dx) : 0)
  let ny = has('y') ? Number(input.y) : cy + (has('dy') ? Number(input.dy) : 0)
  if (!Number.isFinite(nx) || !Number.isFinite(ny))
    return err('x, y, dx, dy must be numbers', summary)
  const size = geomDispSize(geom)
  nx = Math.min(Math.max(nx, 0), size.width)
  ny = Math.min(Math.max(ny, 0), size.height)
  if (nx === cx && ny === cy) {
    return {
      output: `Inserted text T${r.ins.id} is already at x=${fmt(nx)}, y=${fmt(ny)}.`,
      summary,
    }
  }
  // same math as the drag handler: a display-space delta rotated into PDF user space
  const [ax, ay] = viewToPdf(geom, 0, 0)
  const [bx, by] = viewToPdf(geom, nx - cx, ny - cy)
  deps.moveTextInsert(r.ins.id, [origin[0] + (bx - ax), origin[1] + (by - ay)])
  deps.gotoPage(pageIndex + 1)
  return {
    output: `Moved inserted text T${r.ins.id} on page ${pageIndex + 1} to x=${fmt(nx)}, y=${fmt(ny)} (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

function deleteInsertedTextTool(deps: PdfAiDeps, input: Record<string, unknown>): ToolExecution {
  const summary = t('aiToolDeleteInsertedText')
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolveInsert(deps, input.id)
  if ('bad' in r) return err(r.bad, summary)
  deps.deleteTextInsert(r.ins.id)
  return {
    output: `Removed inserted text T${r.ins.id} from page ${r.ins.input.pageIndex + 1} (unsaved; undo with ⌘Z).`,
    mutated: true,
    summary,
  }
}

// ── Image tools ─────────────────────────────────────────────────────
// Model-facing coordinates are display-space points: what the user sees, with a
// top-left origin and page rotation applied (viewToPdf/pdfRectToCss, same as the
// UI edit path). Rects handed to deps stay in PDF user space (y-up, unrotated).

const IMAGE_GAP_PT = 8
const TEXT_GAP_PT = 4
const px2pt = (px: number) => (px * 72) / 96
const fmt = (n: number) => String(Math.round(n))

/** PDF-space rect → display-space box at scale 1 */
const dispBox = (geom: PageGeom, rect: readonly [number, number, number, number]) =>
  pdfRectToCss(geom, rect, 1)

/** Display-space box (top-left origin) → PDF user-space rect */
function dispToPdfRect(
  geom: PageGeom,
  left: number,
  top: number,
  w: number,
  h: number,
): [number, number, number, number] {
  const [ax, ay] = viewToPdf(geom, left, top)
  const [bx, by] = viewToPdf(geom, left + w, top + h)
  return [Math.min(ax, bx), Math.min(ay, by), Math.max(ax, bx), Math.max(ay, by)]
}

/** Stable per-page numbering: top-to-bottom, then left-to-right, as displayed */
const sortImageRefs = (refs: PageImageRef[], geom: PageGeom): PageImageRef[] =>
  [...refs].sort((a, b) => {
    const da = dispBox(geom, a.rect)
    const db = dispBox(geom, b.rect)
    return da.top - db.top || da.left - db.left
  })

function describeImage(ref: PageImageRef, geom: PageGeom, n: number, claimed: boolean): string {
  const d = dispBox(geom, ref.rect)
  const state = claimed ? ' — has a pending unsaved edit' : ''
  return `  image ${n}: ${fmt(d.width)} × ${fmt(d.height)} pt at x=${fmt(d.left)}, y=${fmt(d.top)}, ${
    ref.aboveText ? 'above' : 'below'
  } the text${state}`
}

async function listPageImagesTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const summary = t('aiToolListImages')
  let only: number | null = null
  if (input.page !== undefined) {
    const r = resolvePage(deps, input.page)
    if ('bad' in r) return err(r.bad, summary)
    only = r.origIdx
  }
  const refs = await deps.listImages()
  const byPage = new Map<number, PageImageRef[]>()
  for (const ref of refs) {
    if (deps.isDeleted(ref.pageIndex) || (only !== null && ref.pageIndex !== only)) continue
    byPage.set(ref.pageIndex, [...(byPage.get(ref.pageIndex) ?? []), ref])
  }
  const lines: string[] = []
  for (const [pageIdx, group] of [...byPage].sort((a, b) => a[0] - b[0])) {
    const geom = deps.pageGeom(pageIdx)
    if (!geom) continue
    const size = geomDispSize(geom)
    lines.push(`Page ${pageIdx + 1} (${fmt(size.width)} × ${fmt(size.height)} pt):`)
    sortImageRefs(group, geom).forEach((ref, i) =>
      lines.push(describeImage(ref, geom, i + 1, deps.isImageClaimed(ref))),
    )
  }
  return {
    output:
      lines.join('\n') ||
      (only !== null
        ? `Page ${only + 1} has no embedded images (unsaved inserts from this session are not listed)`
        : 'The document has no embedded images (unsaved inserts from this session are not listed)'),
    summary,
  }
}

/** page + 1-based listing number → the actual image, using the same ordering as the listing */
async function resolveImageRef(
  deps: PdfAiDeps,
  origIdx: number,
  geom: PageGeom,
  rawNumber: unknown,
): Promise<PageImageRef | string> {
  const n = Number(rawNumber)
  if (!Number.isInteger(n) || n < 1) return 'image_number must be a positive integer'
  const onPage = sortImageRefs(
    (await deps.listImages()).filter((ref) => ref.pageIndex === origIdx),
    geom,
  )
  if (onPage.length === 0)
    return `Page ${origIdx + 1} has no embedded images; note that unsaved inserts cannot be edited`
  const ref = onPage[n - 1]
  if (!ref)
    return `Page ${origIdx + 1} only has ${onPage.length} image(s); call list_page_images to see them`
  if (deps.isImageClaimed(ref))
    return `Image ${n} on page ${origIdx + 1} already has a pending unsaved edit; it cannot be edited again before the user saves`
  return ref
}

async function imageSearchTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const query = String(input.query ?? '').trim()
  const summary = t('aiToolImageSearch', { query })
  if (!query) return err('query must not be empty', summary)
  const r = await deps.searchImages(query, Number(input.max_results) || 8)
  // a backend failure must not read as an empty gallery — the model would fabricate image choices
  if (r.method === 'error') {
    return err(
      `image search failed (service error, not an empty result — you may retry): ${r.error ?? 'unknown error'}`,
      summary,
    )
  }
  const lines = r.images.map(
    (im, i) =>
      `${i + 1}. ${im.title || '(untitled)'} [${im.width ?? '?'}x${im.height ?? '?'}]\n   ${im.imageUrl}`,
  )
  return { output: lines.join('\n') || '(no images found)', summary }
}

async function generateImageTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const summary = t('aiToolGenImage')
  const prompt = String(input.prompt ?? '').trim()
  if (!prompt) return err('prompt must not be empty', summary)
  const r = await deps.generateImage({
    prompt,
    aspectRatio: input.aspect_ratio === undefined ? undefined : String(input.aspect_ratio),
  })
  if (!r.url) return err(`image generation failed: ${r.error ?? 'unknown error'}`, summary)
  return {
    output: `Generated image URL: ${r.url}\nInsert it into the document with insert_image.`,
    summary,
  }
}

async function insertImageTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const summary = t('aiToolInsertImage', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const url = String(input.url ?? '')
  if (!/^https?:\/\//.test(url))
    return err('invalid url; pass an imageUrl from image_search or generate_image', summary)
  const geom = deps.pageGeom(r.origIdx)
  if (!geom) return err('Document not ready', summary)
  const size = geomDispSize(geom)
  const fetched = await deps.fetchImage(url)
  // never write after the user hit stop (the download may resolve long after the abort)
  if (signal?.aborted) return err('stopped by the user; the image was not inserted', summary)
  if (!fetched)
    return err(
      'downloading the image failed (it may not be accessible); try another search result or regenerate',
      summary,
    )

  let w: number
  if (input.width !== undefined) {
    w = Number(input.width)
    if (!(w > 0)) return err('width must be a positive number', summary)
  } else {
    w = Math.min(px2pt(fetched.width), size.width / 2)
  }
  let h = (w * fetched.height) / fetched.width
  // never overflow the page box
  const fitK = Math.min(1, size.width / w, size.height / h)
  w *= fitK
  h *= fitK

  // top-left-origin position (tx from left, ty from page top)
  let tx: number
  let ty: number
  const anchor = String(input.anchor_text ?? '').trim()
  if (anchor) {
    const indexPromise = deps.searchIndex()
    if (!indexPromise) return err('Document not ready', summary)
    const entry = (await indexPromise)[r.origIdx]
    const located = entry ? locateOccurrence(entry, anchor, 1) : null
    if (!located) {
      return err(
        `"${anchor}" not found on page ${r.origIdx + 1}; use read_pages to verify the exact text`,
        summary,
      )
    }
    const a = dispBox(geom, located.rect)
    const placement = String(input.placement ?? 'below')
    if (placement === 'above') {
      tx = a.left
      ty = a.top - IMAGE_GAP_PT - h
    } else if (placement === 'right') {
      tx = a.left + a.width + IMAGE_GAP_PT
      ty = a.top + a.height / 2 - h / 2
    } else if (placement === 'left') {
      tx = a.left - IMAGE_GAP_PT - w
      ty = a.top + a.height / 2 - h / 2
    } else {
      tx = a.left
      ty = a.top + a.height + IMAGE_GAP_PT
    }
  } else if (input.x !== undefined || input.y !== undefined) {
    tx = Number(input.x ?? 0)
    ty = Number(input.y ?? 0)
    if (!Number.isFinite(tx) || !Number.isFinite(ty))
      return err('x and y must be numbers (points from the page top-left as displayed)', summary)
  } else {
    tx = (size.width - w) / 2
    ty = (size.height - h) / 2
  }
  tx = Math.min(Math.max(tx, 0), size.width - w)
  ty = Math.min(Math.max(ty, 0), size.height - h)

  const rect = dispToPdfRect(geom, tx, ty, w, h)
  const layer: ImageLayer = input.layer === 'above_text' ? 'aboveText' : 'belowText'
  deps.insertImage(r.origIdx, fetched.png, rect, layer)
  deps.gotoPage(r.origIdx + 1)
  return {
    output: `Inserted the image on page ${r.origIdx + 1}: ${fmt(w)} × ${fmt(h)} pt at x=${fmt(tx)}, y=${fmt(ty)}, ${
      layer === 'aboveText' ? 'above' : 'below'
    } the text (unsaved; the user can drag/resize it, undo with ⌘Z, save with ⌘S).`,
    mutated: true,
    summary,
  }
}

async function transformImageTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const summary = t('aiToolMoveImage', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const geom = deps.pageGeom(r.origIdx)
  if (!geom) return err('Document not ready', summary)
  const size = geomDispSize(geom)
  const ref = await resolveImageRef(deps, r.origIdx, geom, input.image_number)
  if (signal?.aborted) return err('stopped by the user; nothing was changed', summary)
  if (typeof ref === 'string') return err(ref, summary)
  const cur = dispBox(geom, ref.rect)

  let w = input.width === undefined ? undefined : Number(input.width)
  let h = input.height === undefined ? undefined : Number(input.height)
  if ((w !== undefined && !(w > 0)) || (h !== undefined && !(h > 0)))
    return err('width/height must be positive numbers', summary)
  if (w === undefined && h === undefined) {
    w = cur.width
    h = cur.height
  } else {
    // one side given → keep the aspect ratio
    w ??= (h! * cur.width) / cur.height
    h ??= (w * cur.height) / cur.width
  }
  const fitK = Math.min(1, size.width / w, size.height / h)
  w *= fitK
  h *= fitK

  let tx = input.x === undefined ? cur.left : Number(input.x)
  let ty = input.y === undefined ? cur.top : Number(input.y)
  if (!Number.isFinite(tx) || !Number.isFinite(ty))
    return err('x and y must be numbers (points from the page top-left as displayed)', summary)
  tx = Math.min(Math.max(tx, 0), size.width - w)
  ty = Math.min(Math.max(ty, 0), size.height - h)

  const layer: ImageLayer | undefined =
    input.layer === undefined ? undefined : input.layer === 'above_text' ? 'aboveText' : 'belowText'
  deps.transformImage(ref, dispToPdfRect(geom, tx, ty, w, h), layer)
  deps.gotoPage(r.origIdx + 1)
  return {
    output: `Adjusted image ${Number(input.image_number)} on page ${r.origIdx + 1}: now ${fmt(w)} × ${fmt(h)} pt at x=${fmt(tx)}, y=${fmt(ty)}${
      layer ? `, moved ${layer === 'aboveText' ? 'above' : 'below'} the text` : ''
    } (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

async function rotateImageTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const summary = t('aiToolRotateImage', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const geom = deps.pageGeom(r.origIdx)
  if (!geom) return err('Document not ready', summary)
  const ref = await resolveImageRef(deps, r.origIdx, geom, input.image_number)
  if (signal?.aborted) return err('stopped by the user; nothing was changed', summary)
  if (typeof ref === 'string') return err(ref, summary)
  const dir = input.direction === undefined ? 'cw' : String(input.direction)
  if (dir !== 'cw' && dir !== 'ccw' && dir !== '180')
    return err("direction must be 'cw', 'ccw' or '180'", summary)
  const turns = dir === 'cw' ? 1 : dir === 'ccw' ? 3 : 2
  const [x1, y1, x2, y2] = ref.rect
  const rect: [number, number, number, number] =
    turns % 2 === 1
      ? [
          (x1 + x2) / 2 - (y2 - y1) / 2,
          (y1 + y2) / 2 - (x2 - x1) / 2,
          (x1 + x2) / 2 + (y2 - y1) / 2,
          (y1 + y2) / 2 + (x2 - x1) / 2,
        ]
      : [x1, y1, x2, y2]
  deps.transformImage(ref, rect, undefined, turns)
  deps.gotoPage(r.origIdx + 1)
  return {
    output: `Rotated image ${Number(input.image_number)} on page ${r.origIdx + 1} ${
      dir === '180' ? '180°' : dir === 'cw' ? '90° clockwise' : '90° counter-clockwise'
    } about its center (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

async function replaceImageTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const summary = t('aiToolReplaceImage', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const geom = deps.pageGeom(r.origIdx)
  if (!geom) return err('Document not ready', summary)
  const ref = await resolveImageRef(deps, r.origIdx, geom, input.image_number)
  if (typeof ref === 'string') return err(ref, summary)
  const url = String(input.url ?? '')
  if (!/^https?:\/\//.test(url)) return err('url must be a direct http(s) image link', summary)
  const fetched = await deps.fetchImage(url)
  if (signal?.aborted) return err('stopped by the user; nothing was changed', summary)
  if (!fetched)
    return err(
      'The image could not be downloaded or decoded (the link may be inaccessible or not a raster image); pick another result or generate a new one.',
      summary,
    )
  deps.replaceImage(ref, fetched.png)
  deps.gotoPage(r.origIdx + 1)
  return {
    output: `Replaced the pixels of image ${Number(input.image_number)} on page ${r.origIdx + 1} in place; footprint and z-order kept (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

const BAKE_FAILED =
  'The image pixels could not be read, or another pending edit claimed this image meanwhile; nothing was changed'

/** Optional 0..1 fraction argument (undefined → 0) */
function fraction01(raw: unknown, name: string): number | string {
  if (raw === undefined || raw === null) return 0
  const v = Number(raw)
  if (!Number.isFinite(v) || v < 0 || v >= 1) return `${name} must be a fraction from 0 to below 1`
  return v
}

/** Kept region given as fractions of the DISPLAYED box → fractions of the image's
    object-space rect (what cropRect and the rendered pixels use); the two differ
    whenever the page is shown rotated */
function displayCropToObject(
  geom: PageGeom,
  rect: readonly [number, number, number, number],
  disp: CropFractions,
): CropFractions {
  const box = dispBox(geom, rect)
  const [kx1, ky1, kx2, ky2] = dispToPdfRect(
    geom,
    box.left + box.width * disp.l,
    box.top + box.height * disp.t,
    box.width * (disp.r - disp.l),
    box.height * (disp.b - disp.t),
  )
  const [x1, y1, x2, y2] = rect
  const w = x2 - x1 || 1
  const h = y2 - y1 || 1
  const unit = (v: number) => Math.min(1, Math.max(0, v))
  return {
    l: unit((kx1 - x1) / w),
    t: unit((y2 - ky2) / h),
    r: unit((kx2 - x1) / w),
    b: unit((y2 - ky1) / h),
  }
}

/** Parse the per-tool bake arguments (crop fractions still in display space); a string is
    the rejection reason */
function parseBakeOp(
  name: keyof typeof BAKE_SUMMARY_KEYS,
  input: Record<string, unknown>,
): { op: ImageBakeOp; what: string } | string {
  switch (name) {
    case 'flip_image': {
      const axis = String(input.axis ?? '')
      if (axis !== 'horizontal' && axis !== 'vertical')
        return "axis must be 'horizontal' or 'vertical'"
      return {
        op: { kind: 'flip', axis: axis === 'horizontal' ? 'h' : 'v' },
        what: `Flipped ${axis === 'horizontal' ? 'left-right' : 'top-bottom'}`,
      }
    }
    case 'set_image_opacity': {
      const alpha = Number(input.opacity)
      if (!Number.isFinite(alpha) || alpha < 0 || alpha > 1)
        return 'opacity must be a number from 0 (invisible) to 1 (opaque)'
      return {
        op: { kind: 'opacity', alpha },
        what: `Set the opacity to ${Math.round(alpha * 100)}%`,
      }
    }
    case 'crop_image': {
      const insets: number[] = []
      for (const side of ['left', 'top', 'right', 'bottom'] as const) {
        const v = fraction01(input[side], side)
        if (typeof v === 'string') return v
        insets.push(v)
      }
      const [l = 0, t = 0, r = 0, b = 0] = insets
      if (l + r >= 1 || t + b >= 1)
        return 'the insets leave no image: left + right and top + bottom must each stay below 1'
      if (!l && !t && !r && !b) return 'at least one inset must be greater than 0'
      return { op: { kind: 'crop', crop: { l, t, r: 1 - r, b: 1 - b } }, what: 'Cropped' }
    }
    case 'remove_image_background': {
      const tolerance =
        input.tolerance === undefined || input.tolerance === null
          ? DEFAULT_CUTOUT_TOLERANCE
          : Number(input.tolerance)
      if (!Number.isFinite(tolerance) || tolerance < 0 || tolerance > 100)
        return 'tolerance must be a number from 0 to 100'
      return {
        op: { kind: 'cutout', tolerance },
        what: `Removed the background (tolerance ${tolerance}) of`,
      }
    }
  }
}

const BAKE_SUMMARY_KEYS = {
  flip_image: 'aiToolFlipImage',
  set_image_opacity: 'aiToolSetImageOpacity',
  crop_image: 'aiToolCropImage',
  remove_image_background: 'aiToolRemoveImageBackground',
} as const

/** flip / opacity / crop / remove-background: one pixel bake on an existing image */
async function bakeImageTool(
  name: keyof typeof BAKE_SUMMARY_KEYS,
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const summary = t(BAKE_SUMMARY_KEYS[name], { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const geom = deps.pageGeom(r.origIdx)
  if (!geom) return err('Document not ready', summary)
  const parsed = parseBakeOp(name, input)
  if (typeof parsed === 'string') return err(parsed, summary)
  const ref = await resolveImageRef(deps, r.origIdx, geom, input.image_number)
  if (signal?.aborted) return err('stopped by the user; nothing was changed', summary)
  if (typeof ref === 'string') return err(ref, summary)
  // the bake runs on unrotated object-space pixels; the model speaks display space
  const quarterTurned = geom.rot % 180 !== 0
  const op: ImageBakeOp =
    parsed.op.kind === 'crop'
      ? { kind: 'crop', crop: displayCropToObject(geom, ref.rect, parsed.op.crop) }
      : parsed.op.kind === 'flip' && quarterTurned
        ? { kind: 'flip', axis: parsed.op.axis === 'h' ? 'v' : 'h' }
        : parsed.op
  const ok = await deps.bakeImage(ref, op, signal)
  if (signal?.aborted) return err('stopped by the user; nothing was changed', summary)
  if (!ok) return err(BAKE_FAILED, summary)
  deps.gotoPage(r.origIdx + 1)
  const where = `image ${Number(input.image_number)} on page ${r.origIdx + 1}`
  let detail = 'footprint and z-order kept'
  if (op.kind === 'crop') {
    const kept = dispBox(geom, cropRect(ref.rect, op.crop))
    detail = `now ${fmt(kept.width)} × ${fmt(kept.height)} pt at x=${fmt(kept.left)}, y=${fmt(kept.top)}; z-order kept`
  }
  return {
    output: `${parsed.what} ${where} in place; ${detail} (unsaved; the user can undo with ⌘Z and saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

async function deleteImageTool(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const summary = t('aiToolDeleteImage', { page: Number(input.page) })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const r = resolvePage(deps, input.page)
  if ('bad' in r) return err(r.bad, summary)
  const geom = deps.pageGeom(r.origIdx)
  if (!geom) return err('Document not ready', summary)
  const ref = await resolveImageRef(deps, r.origIdx, geom, input.image_number)
  if (signal?.aborted) return err('stopped by the user; nothing was changed', summary)
  if (typeof ref === 'string') return err(ref, summary)
  deps.deleteImage(ref)
  deps.gotoPage(r.origIdx + 1)
  return {
    output: `Deleted image ${Number(input.image_number)} on page ${r.origIdx + 1} (unsaved; the user can undo with ⌘Z and saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

/** Whole-document form field inventory (radios aggregate exportValue lists by field name) */
async function collectFields(
  doc: PDFDocumentProxy,
): Promise<Map<string, { kind: string; page: number; value: string; options: string[] }>> {
  const catalog = await buildFormCatalog(doc)
  const fields = new Map<string, { kind: string; page: number; value: string; options: string[] }>()
  for (const field of catalog.fields.values()) {
    if (field.readOnly || field.kind === 'signature') continue
    fields.set(field.name, {
      kind: field.kind,
      page: field.pageIndex + 1,
      value: field.kind === 'checkbox' ? String(field.checked) : field.value,
      options: field.options,
    })
  }
  return fields
}

async function listFormFields(deps: PdfAiDeps): Promise<ToolExecution> {
  const doc = deps.doc()
  if (!doc) return err('Document not ready', t('aiToolFields', { count: 0 }))
  const fields = await collectFields(doc)
  const edits = deps.formEdits()
  const lines = [...fields].map(([name, f]) => {
    const edit = edits.get(name)
    const value = edit
      ? edit.kind === 'checkbox'
        ? String(!!edit.checked)
        : (edit.value ?? '')
      : f.value
    const opts = f.options.length > 0 ? ` options[${f.options.join(', ')}]` : ''
    return `${name} (${f.kind}, page ${f.page})${opts} current value: ${value || '(empty)'}`
  })
  return {
    output: lines.join('\n') || 'The document has no form fields',
    summary: t('aiToolFields', { count: fields.size }),
  }
}

async function fillFormField(
  deps: PdfAiDeps,
  input: Record<string, unknown>,
): Promise<ToolExecution> {
  const name = String(input.name ?? '')
  const summary = t('aiToolFill', { name })
  if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
  const doc = deps.doc()
  if (!doc || !name) return err('Document not ready or name is empty', summary)
  const fields = await collectFields(doc)
  const field = fields.get(name)
  if (!field)
    return err(`No field named "${name}"; use list_form_fields to see the fields`, summary)
  let edit: FormValueInput
  if (field.kind === 'checkbox') {
    if (typeof input.checked !== 'boolean')
      return err('Checkbox requires the checked parameter', summary)
    edit = { name, kind: 'checkbox', checked: input.checked }
  } else {
    const value = String(input.value ?? '')
    if (field.kind !== 'text' && value && !field.options.includes(value)) {
      return err(
        `Value "${value}" is not among the options: [${field.options.join(', ')}]`,
        summary,
      )
    }
    edit = { name, kind: field.kind as 'text' | 'radio' | 'choice', value }
  }
  deps.applyFormEdit(edit)
  deps.gotoPage(field.page)
  return { output: `Filled ${name} (unsaved; the user saves with ⌘S)`, mutated: true, summary }
}

export async function executePdfTool(
  deps: PdfAiDeps,
  call: AgentToolCall,
  signal?: AbortSignal,
): Promise<ToolExecution> {
  const input = call.input
  switch (call.name) {
    case 'read_pages':
      return readPages(deps, input)
    case 'search_text':
      return searchText(deps, input)
    case 'goto_page': {
      const summary = t('aiToolGoto', { page: Number(input.page) })
      const r = resolvePage(deps, input.page)
      if ('bad' in r) return err(r.bad, summary)
      deps.gotoPage(r.origIdx + 1)
      return { output: `Jumped to page ${r.origIdx + 1}`, summary }
    }
    case 'markup_text':
      return markupText(deps, input)
    case 'edit_text':
      return editText(deps, input)
    case 'edit_block':
      return editBlock(deps, input)
    case 'move_text_block':
      return moveTextBlockTool(deps, input)
    case 'insert_text':
      return insertTextTool(deps, input, signal)
    case 'add_form_mark':
      return addFormMarkTool(deps, input, signal)
    case 'list_inserted_text':
      return listInsertedTextTool(deps, input)
    case 'edit_inserted_text':
      return editInsertedTextTool(deps, input)
    case 'move_inserted_text':
      return moveInsertedTextTool(deps, input)
    case 'delete_inserted_text':
      return deleteInsertedTextTool(deps, input)
    case 'read_annotations':
      return readAnnotations(deps, input)
    case 'add_note':
      return addNoteTool(deps, input, signal)
    case 'reply_note':
      return replyNoteTool(deps, input, signal)
    case 'edit_note':
      return editNoteTool(deps, input, signal)
    case 'delete_markup':
      return deleteMarkupTool(deps, input)
    case 'delete_note':
      return deleteNoteTool(deps, input)
    case 'image_search':
      return imageSearchTool(deps, input)
    case 'generate_image':
      return generateImageTool(deps, input)
    case 'list_page_images':
      return listPageImagesTool(deps, input)
    case 'insert_image':
      return insertImageTool(deps, input, signal)
    case 'transform_image':
      return transformImageTool(deps, input, signal)
    case 'rotate_image':
      return rotateImageTool(deps, input, signal)
    case 'replace_image':
      return replaceImageTool(deps, input, signal)
    case 'delete_image':
      return deleteImageTool(deps, input, signal)
    case 'flip_image':
    case 'set_image_opacity':
    case 'crop_image':
    case 'remove_image_background':
      return bakeImageTool(call.name, deps, input, signal)
    case 'list_form_fields':
      return listFormFields(deps)
    case 'fill_form_field':
      return fillFormField(deps, input)
    case 'rotate_page':
      return rotatePageTool(deps, input)
    case 'set_metadata':
      return setMetadataTool(deps, input)
    case 'delete_page': {
      const summary = t('aiToolDelete', { page: Number(input.page) })
      if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
      const r = resolvePage(deps, input.page)
      if ('bad' in r) return err(r.bad, summary)
      if (!deps.deletePage(r.origIdx)) return err('At least one page must remain', summary)
      return {
        output: `Deleted page ${r.origIdx + 1} (unsaved; can be undone)`,
        mutated: true,
        summary,
      }
    }
    case 'move_page':
      return movePageTool(deps, input)
    case 'reverse_pages': {
      const summary = t('aiToolReversePages')
      if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
      if (deps.pageOrder().length < 2) return err('The document has a single page', summary)
      deps.reversePages()
      return {
        output: `Reversed the page order (unsaved). ${describeOrder(deps.pageOrder())}`,
        mutated: true,
        summary,
      }
    }
    case 'insert_blank_page':
      return insertBlankPageTool(deps, input, signal)
    case 'set_page_size':
      return setPageSizeTool(deps, input, signal)
    case 'crop_pages':
      return cropPagesTool(deps, input, signal)
    case 'extract_pages':
      return extractPagesTool(deps, input, signal)
    case 'split_pdf':
      return splitPdfTool(deps, input, signal)
    case 'split_pages':
      return splitPagesTool(deps, input, signal)
    case 'merge_pages':
      return mergePagesTool(deps, input, signal)
    case 'replace_pages':
      return replacePagesTool(deps, input, signal)
    case 'set_watermark':
      return setWatermarkTool(deps, input)
    case 'set_header_footer':
      return setHeaderFooterTool(deps, input)
    case 'get_outline': {
      const outline = deps.outline()
      const lines: string[] = []
      const walk = (nodes: OutlineNode[], depth: number) => {
        for (const n of nodes) {
          lines.push(`${'  '.repeat(depth)}${n.title}`)
          if (n.items) walk(n.items, depth + 1)
        }
      }
      if (outline) walk(outline, 0)
      return {
        output: lines.join('\n') || 'The document has no outline',
        summary: t('aiToolOutline'),
      }
    }
    case 'create_document': {
      const typeRaw = input.type === undefined ? 'pdf' : String(input.type)
      const summary = t('aiToolCreateDocument')
      if (typeRaw !== 'pdf' && typeRaw !== 'docx' && typeRaw !== 'md')
        return err('type must be one of pdf/docx/md', summary)
      const type: CreateDocumentType = typeRaw
      const title = String(input.title ?? '').trim()
      if (!title) return err('title must not be empty', summary)
      const content = String(input.content ?? '')
      if (!content.trim()) return err('content must not be empty', summary)
      if (type !== 'md') {
        const echo = contentEchoError(content)
        if (echo) return err(echo, summary)
      }
      const r = await deps.createDocument({ type, title, content })
      if (!r.ok) return err(r.error ?? 'creating the document failed', summary)
      const name = `${title}.${type}`
      return {
        output: r.path
          ? `Created the new document at ${r.path} and opened it in a new tab.`
          : `Created the new document "${name}" in a new tab; it saves itself into the default folder.`,
        summary: t('aiToolCreatedDocument', { name }),
      }
    }
    default:
      return err(`Unknown tool: ${call.name}`, call.name)
  }
}
