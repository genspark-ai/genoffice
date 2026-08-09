import type { PDFDocumentProxy } from 'pdfjs-dist'
import type { AgentToolCall, AgentToolDef, ToolExecution } from '@genoffice/agent-core'
import type { OutlineNode } from '../OutlinePanel'
import type { PageEntry, SearchIndex } from '../search'
import { searchInIndex } from '../search'
import { geomDispSize, pdfRectToCss, viewToPdf } from '../annotations'
import type { PageGeom } from '../annotations'
import { EDIT_FONTS } from '../../shared/ipc'
import type {
  FormValueInput,
  ImageLayer,
  ImageSearchResponse,
  MarkupType,
  PageImageRef,
  TextEditInput,
} from '../../shared/ipc'
import { groupPageBlocks } from '../text-block'
import { joinBlockLines, measurePt, wrapText } from '../text-wrap'
import { t } from '../i18n/locale'

/** Text cap per read_pages fed back to the model (the payload is resent in full each turn, so volume must be limited) */
const READ_CHUNK_CHARS = 24_000

/** Capability surface App provides to AI tools; all getters, since the loop outlives render closures */
export interface PdfAiDeps {
  doc(): PDFDocumentProxy | null
  fileName(): string
  pageCount(): number
  /** Original page number of the currently visible page (1-based) */
  currentPage(): number
  readOnly(): boolean
  outline(): OutlineNode[] | null
  searchIndex(): Promise<SearchIndex> | null
  isDeleted(origIdx: number): boolean
  /** Original page number → scroll to that page; returns false if the page was deleted */
  gotoPage(origPage: number): boolean
  addMarkup(type: MarkupType, origIdx: number, rects: [number, number, number, number][]): void
  /** Queue a pending text edit (dry-run validated against the file when possible); null = accepted, string = rejection reason */
  editText(input: TextEditInput): Promise<string | null>
  /** Edit-font ids available on this machine (EDIT_FONTS subset) */
  editFonts(): string[]
  formEdits(): ReadonlyMap<string, FormValueInput>
  applyFormEdit(v: FormValueInput): void
  rotatePage(origIdx: number, dir: 90 | -90): void
  deletePage(origIdx: number): boolean
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
  /** Queue a pending delete of an existing image */
  deleteImage(ref: PageImageRef): void
  searchImages(query: string, maxResults: number): Promise<ImageSearchResponse>
  generateImage(op: { prompt: string; aspectRatio?: string }): Promise<{
    url?: string
    error?: string
  }>
  /** Download a URL (main-process, SSRF-guarded) and re-encode as PNG; null on failure */
  fetchImage(url: string): Promise<{ png: string; width: number; height: number } | null>
}

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
      },
      required: ['page', 'text', 'type'],
    },
  },
  {
    name: 'edit_text',
    description:
      'Replace a short text run on a page (rewrites the PDF content; takes effect on save). old_text must be a verbatim fragment that actually exists on that page (confirm with read_pages or search_text first); only the first occurrence on the page is edited unless occurrence is given. The replacement is drawn from the original position without reflowing the page, so keep it close to the original length; text cannot be deleted (new_text must not be empty).',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        old_text: { type: 'string', description: 'Verbatim text fragment currently on the page' },
        new_text: {
          type: 'string',
          description: 'Replacement text; "\\n" splits it into stacked lines',
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
      "Rewrite a whole paragraph on a page (rewrites the PDF content; takes effect on save). The paragraph is located by paragraph_text — a distinctive verbatim fragment of it. The ENTIRE paragraph is replaced by new_text, which is re-wrapped automatically within the paragraph's original width (the block grows downward when the text is longer). Use edit_text instead to change a few words without reflowing.",
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
            'Full replacement for the paragraph; "\\n" separates paragraphs within the block',
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
      },
      required: ['page', 'paragraph_text', 'new_text'],
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
    description: 'Rotate the given page 90 degrees clockwise or counterclockwise.',
    inputSchema: {
      type: 'object',
      properties: {
        page: { type: 'integer', description: 'Page number (1-based)' },
        direction: {
          type: 'string',
          enum: ['left', 'right'],
          description: 'left = counterclockwise, right = clockwise',
        },
      },
      required: ['page', 'direction'],
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
    name: 'get_outline',
    description:
      'Read the document outline (bookmarks) tree, including entry titles. Returns empty if the document has no outline.',
    inputSchema: { type: 'object', properties: {} },
  },
]

const READONLY_OUTPUT =
  'The document is encrypted and read-only; it cannot be modified. Inform the user.'

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
  for (const m of targets) deps.addMarkup(type, r.origIdx, m.rects)
  deps.gotoPage(r.origIdx + 1)
  return {
    output: `Marked ${targets.length} occurrence(s) on page ${r.origIdx + 1} (unsaved; the user saves with ⌘S)`,
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
  if (!newText) return err('new_text must not be empty (edit_text cannot delete text)', summary)
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
  return {
    output: `Replaced occurrence ${occurrence} of "${oldText}" on page ${r.origIdx + 1} (unsaved; the user saves with ⌘S).${note}`,
    mutated: true,
    summary,
  }
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
  if (!newText) return err('new_text must not be empty (edit_block cannot delete text)', summary)
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
  if (!entry) return err(`Page ${r.origIdx + 1} has no extractable text`, summary)
  const squash = (s: string) => s.replace(/\s+/g, '')
  const key = squash(anchor)
  const hits = groupPageBlocks(entry).filter((b) =>
    squash(b.lines.map((l) => l.text).join('')).includes(key),
  )
  if (hits.length === 0) {
    return err(
      `No paragraph on page ${r.origIdx + 1} contains "${anchor}"; use read_pages to verify the exact text`,
      summary,
    )
  }
  if (hits.length > 1) {
    return err(
      `${hits.length} paragraphs on page ${r.origIdx + 1} contain "${anchor}"; pass a longer, unique fragment`,
      summary,
    )
  }
  const block = hits[0]!
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
      block.align === 'left'
        ? undefined
        : lines.map((l) => {
            const slack = widthPt - measurePt(l, size, cssFamily, cssStyle)
            return Math.max(0, block.align === 'center' ? slack / 2 : slack)
          }),
    align: block.align === 'left' ? undefined : block.align,
  })
  if (reason) return err(`The edit could not be applied: ${reason}`, summary)
  deps.gotoPage(r.origIdx + 1)
  return {
    output: `Replaced the paragraph containing "${anchor}" on page ${r.origIdx + 1} with ${lines.length} reflowed line(s) (unsaved; the user saves with ⌘S).`,
    mutated: true,
    summary,
  }
}

// ── Image tools ─────────────────────────────────────────────────────
// Model-facing coordinates are display-space points: what the user sees, with a
// top-left origin and page rotation applied (viewToPdf/pdfRectToCss, same as the
// UI edit path). Rects handed to deps stay in PDF user space (y-up, unrotated).

const IMAGE_GAP_PT = 8
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

interface RawWidget {
  subtype?: string
  fieldType?: string
  fieldName?: string
  fieldValue?: unknown
  buttonValue?: string
  readOnly?: boolean
  checkBox?: boolean
  radioButton?: boolean
  options?: { exportValue?: unknown; displayValue?: unknown }[]
}

/** Whole-document form field inventory (radios aggregate exportValue lists by field name) */
async function collectFields(
  doc: PDFDocumentProxy,
): Promise<Map<string, { kind: string; page: number; value: string; options: string[] }>> {
  const fields = new Map<string, { kind: string; page: number; value: string; options: string[] }>()
  for (let n = 1; n <= doc.numPages; n++) {
    const page = await doc.getPage(n)
    const annots = (await page.getAnnotations()) as RawWidget[]
    for (const a of annots) {
      if (a.subtype !== 'Widget' || !a.fieldName || a.readOnly) continue
      const value = Array.isArray(a.fieldValue)
        ? String(a.fieldValue[0] ?? '')
        : String(a.fieldValue ?? '')
      if (a.fieldType === 'Tx') {
        fields.set(a.fieldName, { kind: 'text', page: n, value, options: [] })
      } else if (a.fieldType === 'Btn' && a.checkBox) {
        fields.set(a.fieldName, { kind: 'checkbox', page: n, value, options: [] })
      } else if (a.fieldType === 'Btn' && a.radioButton) {
        const cur = fields.get(a.fieldName) ?? { kind: 'radio', page: n, value, options: [] }
        if (typeof a.buttonValue === 'string' && !cur.options.includes(a.buttonValue))
          cur.options.push(a.buttonValue)
        fields.set(a.fieldName, cur)
      } else if (a.fieldType === 'Ch') {
        fields.set(a.fieldName, {
          kind: 'choice',
          page: n,
          value,
          options: (a.options ?? [])
            .map((o) => String(o.exportValue ?? o.displayValue ?? ''))
            .filter(Boolean),
        })
      }
    }
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
    case 'list_form_fields':
      return listFormFields(deps)
    case 'fill_form_field':
      return fillFormField(deps, input)
    case 'rotate_page': {
      const summary = t('aiToolRotate', { page: Number(input.page) })
      if (deps.readOnly()) return err(READONLY_OUTPUT, summary)
      const r = resolvePage(deps, input.page)
      if ('bad' in r) return err(r.bad, summary)
      deps.rotatePage(r.origIdx, input.direction === 'left' ? -90 : 90)
      deps.gotoPage(r.origIdx + 1)
      return { output: `Rotated page ${r.origIdx + 1} (unsaved)`, mutated: true, summary }
    }
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
    default:
      return err(`Unknown tool: ${call.name}`, call.name)
  }
}
