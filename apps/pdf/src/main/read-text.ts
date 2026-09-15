/**
 * Read-only PDF text extraction on pdfium, for consumers that only need the
 * text layer (the MCP `read_pdf` tool; a future search/preview path). Runs in
 * the same main-process WASM as the text editor (text-edit.ts) but never
 * mutates the document, so it bypasses the edit/save helpers and just walks
 * textpages: content-order text, per-page size and a char count that doubles
 * as the "does this page have a text layer at all" probe (scanned pages are 0).
 */
import { chainPdfium, loadPdfium, withDocument, type Pdfium } from './text-edit'

export interface PdfPageText {
  /** 1-based page number */
  page: number
  /** page size in PDF points */
  widthPt: number
  heightPt: number
  /** code units pdfium reports for this page's text layer (0 for scanned pages) */
  chars: number
  /** extracted text with \n line breaks; empty when the page has no text layer */
  text: string
  hasTextLayer: boolean
}

export interface PdfTextDoc {
  pageCount: number
  /** document info dictionary entries worth showing an agent ('' entries omitted) */
  info: { title?: string; author?: string }
  pages: PdfPageText[]
  /** the charBudget ran out before the requested span was fully extracted */
  truncated: boolean
}

export interface ReadPdfTextOptions {
  /** 1-based inclusive; default 1 */
  fromPage?: number
  /** 1-based inclusive; default pageCount */
  toPage?: number
  /** stop extracting once the accumulated text reaches this many chars */
  charBudget?: number
}

/**
 * Metadata read: the tag is an ASCII BYTESTRING, the value comes back UTF-16LE
 * with the byte length (NUL included) as the return value.
 */
function metaText(m: Pdfium, doc: number, tag: string): string {
  const tagBytes = Buffer.from(`${tag}\0`, 'latin1')
  const tagPtr = m._malloc(tagBytes.length)
  if (!tagPtr) return ''
  try {
    m.HEAPU8.set(tagBytes, tagPtr)
    const size = m._FPDF_GetMetaText(doc, tagPtr, 0, 0)
    if (size <= 2) return ''
    const buf = m._malloc(size)
    if (!buf) return ''
    try {
      const written = m._FPDF_GetMetaText(doc, tagPtr, buf, size)
      if (written <= 0) return ''
      return Buffer.from(m.HEAPU8.subarray(buf, buf + written))
        .toString('utf16le')
        .replace(/\0+$/, '')
    } finally {
      m._free(buf)
    }
  } finally {
    m._free(tagPtr)
  }
}

/** whole textpage text (UTF-16LE, pdfium uses \r\n line breaks) normalized to \n */
function pageText(m: Pdfium, textPage: number, count: number): string {
  if (count <= 0) return ''
  const buf = m._malloc(count * 2)
  if (!buf) return ''
  try {
    const written = m._FPDFText_GetText(textPage, 0, count, buf)
    if (written <= 0) return ''
    return Buffer.from(m.HEAPU8.subarray(buf, buf + written * 2))
      .toString('utf16le')
      .replace(/\0+$/, '')
      .replace(/\r\n/g, '\n')
      .replace(/\r/g, '\n')
      .trimEnd()
  } finally {
    m._free(buf)
  }
}

/**
 * Extract text from `bytes` (a whole PDF file). Page span is clamped to the
 * document; extraction stops early (with `truncated`) once charBudget is spent.
 */
export function readPdfText(
  bytes: Uint8Array,
  options: ReadPdfTextOptions = {},
): Promise<PdfTextDoc> {
  return chainPdfium(async () => {
    const m = await loadPdfium()
    return withDocument(m, bytes, async (doc) => {
      const pageCount = m._FPDF_GetPageCount(doc)
      const from = Math.max(1, Math.min(options.fromPage ?? 1, pageCount))
      const to = Math.max(from, Math.min(options.toPage ?? pageCount, pageCount))
      const budget = options.charBudget ?? Number.POSITIVE_INFINITY

      const info: { title?: string; author?: string } = {}
      const title = metaText(m, doc, 'Title')
      if (title) info.title = title
      const author = metaText(m, doc, 'Author')
      if (author) info.author = author

      const pages: PdfPageText[] = []
      let spent = 0
      let truncated = false
      for (let index = from; index <= to && !truncated; index++) {
        const page = m._FPDF_LoadPage(doc, index - 1)
        if (!page) {
          pages.push({
            page: index,
            widthPt: 0,
            heightPt: 0,
            chars: 0,
            text: '',
            hasTextLayer: false,
          })
          continue
        }
        try {
          const widthPt = m._FPDF_GetPageWidthF(page)
          const heightPt = m._FPDF_GetPageHeightF(page)
          const textPage = m._FPDFText_LoadPage(page)
          if (!textPage) {
            pages.push({
              page: index,
              widthPt,
              heightPt,
              chars: 0,
              text: '',
              hasTextLayer: false,
            })
            continue
          }
          try {
            const chars = m._FPDFText_CountChars(textPage)
            let text = spent < budget ? pageText(m, textPage, chars) : ''
            if (spent + text.length > budget) {
              text = text.slice(0, Math.max(0, budget - spent))
              truncated = true
            }
            spent += text.length
            pages.push({ page: index, widthPt, heightPt, chars, text, hasTextLayer: chars > 0 })
          } finally {
            m._FPDFText_ClosePage(textPage)
          }
        } finally {
          m._FPDF_ClosePage(page)
        }
      }
      return { pageCount, info, pages, truncated }
    })
  })
}
