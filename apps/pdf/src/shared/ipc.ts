import type { Lang } from '@genoffice/i18n'
import type { AiSettings, AiStreamChunk, AiStreamRequest } from '@genoffice/ai-provider'

export const PDF_CHANNELS = {
  consumePending: 'pdf:consume-pending',
  readFile: 'pdf:read-file',
  save: 'pdf:save',
  validateTextEdits: 'pdf:validate-text-edits',
  listEditFonts: 'pdf:list-edit-fonts',
  listPageImages: 'pdf:list-page-images',
  listStaticFormFills: 'pdf:list-static-form-fills',
  pageImagePng: 'pdf:page-image-png',
  pagePreviewPng: 'pdf:page-preview-png',
  extractPages: 'pdf:extract-pages',
  insertPdf: 'pdf:insert-pdf',
  insertBlankPage: 'pdf:insert-blank-page',
  splitPdf: 'pdf:split-pdf',
  mergePdf: 'pdf:merge-pdf',
  mergePages: 'pdf:merge-pages',
  replacePages: 'pdf:replace-pages',
  setPageSize: 'pdf:set-page-size',
  splitPages: 'pdf:split-pages',
  cropPages: 'pdf:crop-pages',
  exportImages: 'pdf:export-images',
  generateImage: 'pdf:generate-image',
  listSignatures: 'pdf:list-signatures',
  addSignature: 'pdf:add-signature',
  removeSignature: 'pdf:remove-signature',
  getUsername: 'pdf:get-username',
  dirtyChanged: 'pdf:dirty-changed',
  closeSaveRequest: 'pdf:close-save-request',
  closeSaveResult: 'pdf:close-save-result',
  saveAsRequest: 'pdf:save-as-request',
  saveAsResult: 'pdf:save-as-result',
  saveAsFlow: 'pdf:save-as-flow',
  printRequest: 'pdf:print-request',
  getLanguage: 'app:get-language',
  languageChanged: 'app:language-changed',
  getTheme: 'app:get-theme',
  themeChanged: 'app:theme-changed',
} as const

export const VISUAL_SIGNATURE_CONTENT_PREFIX = 'GenOffice visual signature field: '

/** Signature strokes: pad pixel coords, scaled proportionally and y-flipped when placed on the page */
export interface SignatureStrokes {
  paths: number[][]
  width: number
  height: number
}

/** Confirmed signature awaiting placement: hand strokes (Ink) or a bitmap (Stamp) */
export type SignatureData =
  | ({ kind: 'strokes' } & SignatureStrokes)
  | {
      kind: 'image'
      /** base64 PNG, without the data: prefix */
      image: string
      width: number
      height: number
    }

/** A reusable signature persisted in userData (shared across documents, WPS-style) */
export interface SavedSignature {
  id: string
  createdAt: number
  data: SignatureData
}

export type UiTheme = 'light' | 'dark' | 'system'

export type MarkupType = 'highlight' | 'underline' | 'strikeout'

/** A text markup to write; quads are 4-point groups in PDF coords (y up) [x1,yTop,x2,yTop,x1,yBottom,x2,yBottom] */
export interface MarkupInput {
  pageIndex: number
  type: MarkupType
  /** rgb normalized to 0-1 */
  color: [number, number, number]
  quads: number[][]
}

/** Delete an annotation already saved in the file. Object number is only a lookup
    hint; subtype + rect guard against (and recover from) object renumbering by a rewrite.
    'note' targets Text (comment) annotations. */
export interface AnnotDeleteInput {
  pageIndex: number
  /** PDF object number (pdf.js annotation id "123R" → 123) */
  objNum: number
  subtype: MarkupType | 'note'
  /** Annotation /Rect in PDF user space, for fallback matching */
  rect: [number, number, number, number]
  /** /Contents to match. Required identity for notes: every comment of a thread shares
      the root's rect, so rect alone would delete the wrong member. */
  contents?: string
}

/** Reply target: a note (Text) annotation saved in the file. The object number is only a
    lookup hint — pdfium stages that run before pdf-lib may renumber objects — so the
    parent is confirmed by rect + contents at write time. */
export interface NoteReplyTarget {
  objNum: number
  rect: [number, number, number, number]
  contents: string
}

/** Rewrite the /Contents of a note (Text) annotation saved in the file, in place —
    the object keeps its number so saved replies' /IRT chains stay intact (unlike
    delete + re-add). Identity matches like NoteReplyTarget: rect + current contents,
    with the object number as a tie-break hint. */
export interface NoteEditInput {
  pageIndex: number
  objNum: number
  rect: [number, number, number, number]
  /** /Contents currently in the file (identity match) */
  oldContents: string
  contents: string
}

/** Drawing annotations (all coords in PDF user space, y up).
    One union member per kind; a union-literal kind would break TS narrowing. */
interface DrawBase {
  pageIndex: number
  color: [number, number, number]
  width: number
  /** AcroForm field explicitly associated with a visual Ink signature. */
  formFieldName?: string
}

export type DrawingInput =
  | (DrawBase & {
      kind: 'ink'
      /** Each stroke as [x1,y1,x2,y2,...] */
      paths: number[][]
    })
  | (DrawBase & { kind: 'rect'; rect: [number, number, number, number] })
  | (DrawBase & { kind: 'ellipse'; rect: [number, number, number, number] })
  | (DrawBase & { kind: 'line'; from: [number, number]; to: [number, number] })
  | (DrawBase & { kind: 'arrow'; from: [number, number]; to: [number, number] })
  | {
      /** Image signature/stamp placed by the user; written as a Stamp annotation */
      kind: 'image'
      pageIndex: number
      /** base64 PNG, without the data: prefix */
      image: string
      /** PDF user space [x1,y1,x2,y2] */
      rect: [number, number, number, number]
      /** AcroForm field explicitly associated with a visual image signature. */
      formFieldName?: string
    }
  | {
      kind: 'note'
      pageIndex: number
      color: [number, number, number]
      at: [number, number]
      contents: string
      /** Annotation author (/T); omitted → 'GenOffice' */
      author?: string
      /** Creation time (ms since epoch) → /CreationDate and /M; omitted → save time */
      createdMs?: number
      /** Key for parenting replies inside the same request (parents sort before children) */
      localId?: string
      /** This note replies to a note saved in the file */
      replyToSaved?: NoteReplyTarget
      /** This note replies to another note in this request, by its localId */
      replyToLocalId?: string
    }

/**
 * Stamp layer (watermark/header/footer/page numbers all go through it).
 * The renderer rasterizes the bitmap via canvas (with rotation and fonts, bypassing
 * pdf-lib's lack of CJK support); the main process only embeds and positions it.
 */
export interface StampInput {
  pageIndex: number
  /** base64 PNG, without the data: prefix */
  image: string
  /** PDF user space [x1,y1,x2,y2] */
  rect: [number, number, number, number]
  opacity?: number
}

/** Document info; an empty string clears the field */
export interface MetadataInput {
  title?: string
  author?: string
  subject?: string
  keywords?: string
}

/**
 * Content-stream text replacement (pdfium). The renderer only knows pdf.js text items;
 * the main process re-locates the actual text objects by rect intersection and verifies
 * against oldText, because one pdf.js item may span many objects (CJK is often one per glyph).
 */
export interface TextEditInput {
  pageIndex: number
  /** PDF user space [x1,y1,x2,y2] of the run being replaced */
  rect: [number, number, number, number]
  /** Text currently displayed in rect; matching is whitespace-insensitive */
  oldText: string
  /** Replacement; '\n' splits it into stacked lines (one text object per line) */
  newText: string
  /** Original run's font size in PDF pt (fallback sizing for the rebuilt object) */
  fontSize: number
  /** User-chosen size in PDF pt; absent = keep the original run's size */
  newFontSize?: number
  /** User-chosen fill color (0-255 RGB); absent = keep the original run's color */
  newColor?: [number, number, number]
  /** Selection-level fill colors: [start,end) code-unit ranges into newText, ascending
      and non-overlapping. Chars outside every range keep newColor ?? the original run's
      color. The engine splits each rebuilt line into one text object per color.
      Legacy input form — new senders use styleRuns; still accepted by the engine. */
  colorRuns?: { start: number; end: number; color: [number, number, number] }[]
  /** Selection-level style overrides: [start,end) code-unit ranges into newText,
      ascending and non-overlapping. Each present field overrides the whole-edit value
      (newColor / newFont / newFontSize / newBold / newItalic) for the covered range;
      absent fields inherit it. The engine splits each rebuilt line into one text
      object per distinct style, drawing each with its own font face and size.
      Generalizes colorRuns. */
  styleRuns?: {
    start: number
    end: number
    color?: [number, number, number]
    font?: string
    size?: number
    bold?: boolean
    italic?: boolean
  }[]
  /** User-chosen rebuild font (EDIT_FONTS id); absent = keep the original font (embedded
      subset when it covers the replacement, else the same-named installed font, else
      fallback). An explicit id is honored only when the face's cmap covers the whole
      replacement (else fallback) */
  newFont?: string
  /** Baseline origin (PDF user space) for the rebuilt first line. Paragraph edits set
      this to the block's left edge / first baseline — the matched anchor object may be
      an indented or mid-paragraph run. Absent = the anchor object's own position.
      Ignored when the edit resolves to a fragment of a larger run. */
  origin?: [number, number]
  /** Baseline-to-baseline step between '\n' lines in PDF pt (paragraph edits pass the
      block's original leading); absent = fontSize × 1.2. Ignored for fragment matches. */
  lineLeading?: number
  /** Per-'\n'-line horizontal offset from origin.x in PDF pt — centered/right-aligned
      paragraph reflow places each line at its own x. Missing entries mean 0; ignored
      for fragment matches. */
  lineXOffsets?: number[]
  /** Paragraph alignment, renderer-side metadata (preview rendering and draft reopen);
      the engine positions lines purely from origin + lineXOffsets */
  align?: 'left' | 'center' | 'right'
  /** The paragraph editor's logical text as typed (hard line breaks intact),
      renderer-side metadata for reopening the draft — the wrapped newText can't
      tell a soft wrap from a deliberate break. The engine ignores it. */
  blockSource?: string
  /** Bold/italic for the rebuilt run, resolved via font variants: the chosen edit
      font's matching face file, or the original family's matching installed face.
      Degrades silently to the base face when no variant covers the replacement. */
  newBold?: boolean
  newItalic?: boolean
  /** Move the matched run as-is: every matched text object is translated by this
      PDF-user-space delta instead of being rewritten (fonts/kerning/colors survive
      untouched). Requires the match to cover the objects exactly (whole) — moving a
      fragment would drag the surrounding text along, so such edits are skipped.
      newText must be textually equivalent to oldText; nothing is rebuilt. */
  translate?: [number, number]
}

/** A new searchable/selectable text object inserted into a page content stream. */
export interface TextInsertInput {
  pageIndex: number
  /** First-line baseline origin in PDF user space. */
  origin: [number, number]
  /** Text to insert; '\n' creates stacked text objects. */
  text: string
  fontSize: number
  color: [number, number, number]
  font?: string
  bold?: boolean
  italic?: boolean
  lineLeading?: number
  lineXOffsets?: number[]
  align?: 'left' | 'center' | 'right'
  /** Final displayed page rotation; insertion matrix counter-rotates text into screen space. */
  rotate?: number
}

/** Dry-run match result for one pending text edit */
export interface TextEditValidation {
  /** null = the edit would apply; string = reason it would be skipped */
  reason: string | null
  /** Ink bounds [x1,y1,x2,y2] of the matched text objects (PDF user space). The renderer's
      edit rect is a pdf.js layout box (font metrics); glyph ink can poke out of it, so the
      preview cover must use the engine's real bounds or remnants of the old run show through */
  bounds?: [number, number, number, number]
  /** Fill colors the matched objects already draw with (earlier saved selection colors),
      as [start,end) ranges into oldText. Present only when the match spans differently
      colored objects. Editors seed their selection-color state from these so existing
      colors show while editing and survive the rebuild. */
  colorRuns?: { start: number; end: number; color: [number, number, number] }[]
  /** Base fill color of the matched run (its first object in reading order), for
      display while editing — a uniformly colored run has no colorRuns to seed from */
  baseColor?: [number, number, number]
}

/** Curated fonts selectable for rebuilt text runs. Single-face .ttf on every platform we
    ship — the subsetter cannot read .ttc collections, which rules out the macOS CJK faces.
    Availability is machine-dependent: the main process reports the usable subset. */
export const EDIT_FONTS = [
  { id: 'arial', label: 'Arial', css: "Arial, 'Helvetica Neue', sans-serif" },
  { id: 'times', label: 'Times New Roman', css: "'Times New Roman', Times, serif" },
  { id: 'courier', label: 'Courier New', css: "'Courier New', monospace" },
] as const

/** Target z-band for a content-stream image: under every text run (above the page
    background) or on top of the page's content */
export type ImageLayer = 'belowText' | 'aboveText'

/**
 * Content-stream image operations (pdfium). Existing images are addressed by their
 * bounds as reported by listPageImages — the main process re-matches the actual object
 * by rect at save time (fail-soft, like text edits), since object indices shift.
 */
export type ImageEditInput =
  | {
      kind: 'insertImage'
      pageIndex: number
      /** base64 PNG, without the data: prefix */
      image: string
      /** PDF user space [x1,y1,x2,y2] footprint */
      rect: [number, number, number, number]
      layer: ImageLayer
      /** Final display rotation of the page (0/90/180/270); the image content is
          counter-rotated so it shows upright on rotated pages */
      rotate?: number
    }
  | {
      kind: 'transformImage'
      pageIndex: number
      /** Bounds of the target image when listed (match key) */
      oldRect: [number, number, number, number]
      /** Final footprint (already width/height-swapped when quarterTurns is odd) */
      rect: [number, number, number, number]
      /** Present = also move the object to this z-band */
      layer?: ImageLayer
      /** Screen-clockwise quarter turns (0-3) applied about the image center.
          2D rotations commute with the page /Rotate mapping, so no adjustment
          for rotated pages is needed */
      quarterTurns?: number
    }
  | {
      kind: 'replaceImage'
      pageIndex: number
      /** Bounds of the target image when listed (match key) */
      oldRect: [number, number, number, number]
      /** Final footprint (drag/resize/rotate after the swap works like a transform) */
      rect: [number, number, number, number]
      /** base64 PNG replacing the image's pixels */
      image: string
      layer?: ImageLayer
      quarterTurns?: number
    }
  | { kind: 'deleteImage'; pageIndex: number; oldRect: [number, number, number, number] }

/** An existing content-stream image on a page */
export interface PageImageRef {
  pageIndex: number
  /** PDF user space [x1,y1,x2,y2] */
  rect: [number, number, number, number]
  /** Painted after the page's first text run (i.e. covers text it overlaps) */
  aboveText: boolean
}

/** Editable metadata for a GenOffice static form fill embedded as a page image. */
export interface StaticFormFillRecord {
  id: string
  kind: 'text' | 'check' | 'cross'
  pageIndex: number
  rect: [number, number, number, number]
  text?: string
  fontSize?: number
  color?: string
  align?: 'left' | 'center' | 'right'
}

/** Live-preview render request: a page region with some images removed */
export interface PagePreviewRequest {
  path: string
  pageIndex: number
  /** Images to erase, addressed by their listed rects (PDF user space) */
  excludeRects: [number, number, number, number][]
  /** Saved annotations to erase. Full identity data is required because object numbers
      may be stale after a save rewrites the PDF. */
  excludeAnnots?: AnnotDeleteInput[]
  /** Region to render, in display coords at scale 1 (after total rotation, y down) */
  clip: { x: number; y: number; width: number; height: number }
  /** Output bitmap width in px (height follows the clip aspect) */
  pxWidth: number
  /** Extra unsaved display rotation on top of the page's /Rotate: 0-3 quarter turns cw */
  rotate: number
}

/** An image edit that could not be matched to the document at save time and was skipped */
export interface ImageEditFailure {
  /** Position in the request's imageEdits array */
  editIndex: number
  pageIndex: number
  reason: string
}

export interface FormValueInput {
  name: string
  kind: 'text' | 'checkbox' | 'radio' | 'choice'
  /** For radio: selected exportValue; for choice: selected option; empty string clears selection */
  value?: string
  checked?: boolean
}

export interface SavePdfRequest {
  path: string
  /**
   * Save As destination. When set, `path` is only read (source bytes) and the edited PDF
   * is written to this path instead — the original file must never be mutated.
   * Must match the target granted to the view by the main process (save dialog pick).
   */
  targetPath?: string
  markups: MarkupInput[]
  /** Saved markup annotations to remove (applied before every other stage) */
  annotDeletes?: AnnotDeleteInput[]
  drawings: DrawingInput[]
  /** In-place content edits of saved note comments. Applied after `drawings`, so
      replies in the same request still match their /IRT parent by its old contents. */
  noteEdits?: NoteEditInput[]
  formValues: FormValueInput[]
  stamps: StampInput[]
  /** Applied to the source bytes before all annotation work — these rewrite page content streams */
  textEdits?: TextEditInput[]
  /** New searchable text objects, applied in the same content-stream stage as textEdits. */
  textInserts?: TextInsertInput[]
  /** Content-stream image operations, applied right after textEdits (same pdfium stage) */
  imageEdits?: ImageEditInput[]
  /** Complete resulting set; omitted means preserve existing embedded metadata. */
  staticFormFills?: StaticFormFillRecord[]
  /** Page rotation deltas (original page index → multiple of 90 clockwise) */
  rotations?: { pageIndex: number; delta: number }[]
  /** Pages to delete (original page indices) */
  deletedPages?: number[]
  /** New page order (array of original page indices, excluding deleted); omitted if unreordered */
  pageOrder?: number[]
  metadata?: MetadataInput
}

/** A text edit that could not be matched to the document at save time and was skipped */
export interface TextEditFailure {
  pageIndex: number
  oldText: string
  reason: string
}

export interface TextInsertFailure {
  editIndex: number
  pageIndex: number
  reason: string
}

export type SavePdfResult =
  | {
      ok: true
      skippedTextEdits?: TextEditFailure[]
      skippedTextInserts?: TextInsertFailure[]
      skippedImageEdits?: ImageEditFailure[]
    }
  | { ok: false; error: string }

/** Dry-run matching for pending text edits against the file on disk (no mutation) */
export interface ValidateTextEditsRequest {
  path: string
  edits: TextEditInput[]
}

/** Extract pages into a new PDF written to the GenOffice save dir and opened in a new tab */
export interface ExtractPagesRequest {
  path: string
  /** Original page indices */
  pages: number[]
  suggestedName: string
}

export type ExtractPagesResult = { ok: true; savedPath: string } | { ok: false; error: string }

/** Insert (merge) another PDF after a page of the current file: main process shows a picker and writes back immediately */
export interface InsertPdfRequest {
  path: string
  /** Insert after this original page index; -1 means front of the document */
  afterPageIndex: number
}

export type InsertPdfResult =
  { ok: true; insertedCount: number } | { ok: true; canceled: true } | { ok: false; error: string }

/** Insert a blank page (sized like the neighboring page) and write back immediately — no dialog */
export interface InsertBlankPageRequest {
  path: string
  /** Insert after this original page index; -1 means front of the document */
  afterPageIndex: number
}

export type InsertBlankPageResult = { ok: true } | { ok: false; error: string }

/** Split the file into chunks of chunkSize pages: main shows a folder picker and writes one PDF per chunk */
export interface SplitPdfRequest {
  path: string
  chunkSize: number
  /** Base for output file names: <baseName>-1.pdf, <baseName>-2.pdf, … */
  baseName: string
}

export type SplitPdfResult =
  | { ok: true; savedDir: string; count: number }
  | { ok: true; canceled: true }
  | { ok: false; error: string }

/** Merge the current file with other PDFs into a new file: main shows a multi-file picker then a save dialog */
export interface MergePdfRequest {
  path: string
  suggestedName: string
}

export type MergePdfResult =
  | { ok: true; savedPath: string; appendedCount: number }
  | { ok: true; canceled: true }
  | { ok: false; error: string }

/** N-up imposition (WPS "merge pages"): every perSheet pages onto one sheet, saved as a new file */
export interface MergePagesRequest {
  path: string
  /** Pages per sheet, 2-16 */
  perSheet: number
  /** Fill order: horizontal = left→right then down, vertical = top→bottom then right */
  direction: 'horizontal' | 'vertical'
  /** Draw hairline separators on the internal cell boundaries */
  separator: boolean
  suggestedName: string
}

export type MergePagesResult =
  { ok: true; savedPath: string } | { ok: true; canceled: true } | { ok: false; error: string }

/** Replace pages with all pages of another PDF (main shows a picker), written back in place */
export interface ReplacePagesRequest {
  path: string
  /** Original page indices to replace */
  pages: number[]
}

export type ReplacePagesResult =
  | { ok: true; removed: number; inserted: number }
  | { ok: true; canceled: true }
  | { ok: false; error: string }

/** Resize all pages to a target paper size (points), content scaled to fit; written back in place */
export interface SetPageSizeRequest {
  path: string
  width: number
  height: number
}

export type SetPageSizeResult = { ok: true } | { ok: false; error: string }

/** Split every page into a grid of pages (inverse of merge pages), written to the
 * GenOffice save dir and opened in a new tab */
export interface SplitPagesRequest {
  path: string
  perPage: 2 | 4 | 9
  suggestedName: string
}

export type SplitPagesResult = { ok: true; savedPath: string } | { ok: false; error: string }

/** Shrink the CropBox of the given pages to a fractional rect of the displayed page; in place */
export interface CropPagesRequest {
  path: string
  /** Original page indices */
  pages: number[]
  /** Fractions of the displayed page (after /Rotate), y down from the top */
  rect: { l: number; t: number; r: number; b: number }
}

export type CropPagesResult = { ok: true } | { ok: false; error: string }

/** Export pages as PNG: renderer rasterizes the bitmaps, main process shows a dialog and writes to disk */
export interface ExportImagesRequest {
  /** base64 PNGs (without the data: prefix), in page order */
  images: string[]
  /** 1-based page numbers, same length as images, used for file names */
  pageNumbers: number[]
  baseName: string
}

export type ExportImagesResult =
  | { ok: true; savedDir: string; count: number }
  | { ok: true; canceled: true }
  | { ok: false; error: string }

/** AI channels are app-wide shared ipcMain handlers (shell registers via docs-main registerAiIpc); pass-through only */
export const AI_CHANNELS = {
  getSettings: 'ai:get-settings',
  stream: 'ai:stream',
  streamChunk: 'ai:stream-chunk',
  streamCancel: 'ai:stream-cancel',
  imageSearch: 'ai:image-search',
  fetchImage: 'ai:fetch-image',
} as const

export interface ImageSearchResponse {
  images: Array<{
    title: string
    imageUrl: string
    sourceUrl: string
    source: string
    width?: number
    height?: number
  }>
  method: string
  /** failure reason when method === 'error' */
  error?: string
}

/** API exposed by preload to the renderer (window.pdfApi) */
export interface PdfApi {
  /** Take the pdf path pending for this view (queued at tab creation); null if none */
  consumePending(): Promise<string | null>
  /** Read pdf bytes. Only paths granted to this view are allowed */
  readFile(path: string): Promise<ArrayBuffer>
  /** Write markups/form values/page ops back to the original file (pdf-lib, content streams untouched); path grants same as readFile. With targetPath set (Save As), the original is only read and the result goes to targetPath */
  save(request: SavePdfRequest): Promise<SavePdfResult>
  /** Dry-run match of pending text edits against the file: reason null = would apply */
  validateTextEdits(request: ValidateTextEditsRequest): Promise<TextEditValidation[]>
  /** EDIT_FONTS ids whose font file exists on this machine */
  listEditFonts(): Promise<string[]>
  /** Enumerate the content-stream images of every page (for image edit mode) */
  listPageImages(path: string): Promise<PageImageRef[]>
  /** Read GenOffice static-fill metadata stored inside the PDF. */
  listStaticFormFills(path: string): Promise<StaticFormFillRecord[]>
  /** Render one existing image object to PNG (base64) for move/resize ghost previews; null if it can't be matched */
  pageImagePng(request: {
    path: string
    pageIndex: number
    rect: [number, number, number, number]
    /** Upscale factor over the on-page size (default 1); use >1 when the pixels feed a
        pixel edit (crop/flip/…) so the baked result keeps print-quality resolution */
    scale?: number
  }): Promise<string | null>
  /** Live-preview render of a page region with the given images removed (base64 PNG);
      the renderer patches it over the raster so touched images vanish before save */
  pagePreviewPng(request: PagePreviewRequest): Promise<string | null>
  extractPages(request: ExtractPagesRequest): Promise<ExtractPagesResult>
  insertPdf(request: InsertPdfRequest): Promise<InsertPdfResult>
  insertBlankPage(request: InsertBlankPageRequest): Promise<InsertBlankPageResult>
  splitPdf(request: SplitPdfRequest): Promise<SplitPdfResult>
  mergePdf(request: MergePdfRequest): Promise<MergePdfResult>
  mergePages(request: MergePagesRequest): Promise<MergePagesResult>
  replacePages(request: ReplacePagesRequest): Promise<ReplacePagesResult>
  setPageSize(request: SetPageSizeRequest): Promise<SetPageSizeResult>
  splitPages(request: SplitPagesRequest): Promise<SplitPagesResult>
  cropPages(request: CropPagesRequest): Promise<CropPagesResult>
  exportImages(request: ExportImagesRequest): Promise<ExportImagesResult>
  /** Web image search for AI tools (app-wide ai:image-search handler) */
  imageSearch(query: string, maxResults?: number): Promise<ImageSearchResponse>
  /** Download an image URL in the main process (SSRF-guarded, avoids CORS); null on failure */
  fetchImage(url: string): Promise<{ base64: string; mime: string } | null>
  /** AI image generation via Genspark (gsk); returns a downloadable URL or an error message */
  generateImage(op: { prompt: string; aspectRatio?: string }): Promise<{
    url?: string
    error?: string
  }>
  /** Saved signatures reusable across documents (persisted in userData), newest first */
  listSavedSignatures(): Promise<SavedSignature[]>
  /** Persist a signature for reuse; returns the updated list (capped, deduplicated) */
  addSavedSignature(data: SignatureData): Promise<SavedSignature[]>
  /** Delete one saved signature by id; returns the updated list */
  removeSavedSignature(id: string): Promise<SavedSignature[]>
  /** OS account name, used as the default author of new note comments; '' when unavailable */
  getUsername(): Promise<string>
  /** Mirror unsaved-changes state to the main process; drives the save prompt before closing a tab/window */
  setDirty(dirty: boolean): void
  /** Main process picked "Save" in the close prompt → renderer saves and replies via sendCloseSaveResult */
  onCloseSaveRequest(handler: () => void): () => void
  sendCloseSaveResult(ok: boolean): void
  /** Shell menu Save As → renderer writes pending edits to targetPath only (original untouched) and replies via sendSaveAsResult */
  onSaveAsRequest(handler: (targetPath: string) => void): () => void
  sendSaveAsResult(ok: boolean): void
  /** True while the shell's Save As flow (dialog included) is open — the renderer pauses autosave, since the dialog's window blur would otherwise trigger a save into the original */
  onSaveAsFlow(handler: (inFlight: boolean) => void): () => void
  /** Shell menu Print → renderer runs its print flow (save, rasterize, system dialog) */
  onPrintRequest(handler: () => void): () => void
  getLanguage(): Promise<Lang>
  onLanguageChanged(handler: (lang: Lang) => void): () => void
  getTheme(): Promise<UiTheme>
  onThemeChanged(handler: (theme: UiTheme) => void): () => void
  /** press on the shell chrome (tab strip is a sibling WebContentsView whose
   *  clicks produce no DOM event here) — dismiss open popovers */
  onChromePressed(handler: () => void): () => void
  getAiSettings(): Promise<AiSettings>
  aiStream(request: AiStreamRequest): Promise<void>
  aiStreamCancel(requestId: string): Promise<void>
  onAiStream(handler: (chunk: AiStreamChunk) => void): () => void
}
