/**
 * Shared PDF → IR pipeline (P25): the extract → furniture-dedup → analyze →
 * confidence-downgrade loop that used to live inside convertPdfToDocx, split
 * out so the PPTX exporter consumes the exact same IR without duplicating the
 * conversion policy. Pure and synchronous — rebuild layers stay per-format.
 */
import { analyzePage, PAGE_CONFIDENCE_MIN } from './analyze'
import { classifyPages } from './analyze/canvas'
import { detectFurniture, type FurnitureHf } from './analyze/furniture'
import {
  extractPage,
  PdfLoadError,
  readDocMetadata,
  renderPageByIndexPng,
  withPdfDocument,
} from './extract'
import type { ExtractedPage, PdfiumModule } from './extract'
import { coverageRatio, type Rect } from './geometry'
import type { IrPage, PageBlock, TextBlock } from './ir'
import { tryOcrScannedPage, type OcrEngine } from './ocr'

export interface ConvertOptions {
  /** initialized @embedpdf/pdfium module (see package header comment) */
  pdfium: PdfiumModule
  onProgress?: (page: number, total: number) => void
  /** raster scale for fallback page renders (pixels per point), default 2 */
  renderScale?: number
  /** user password for encrypted PDFs (P22); load failures throw PdfLoadError */
  password?: string
  /**
   * local OCR engine for scanned pages (platform-supplied; see src/ocr.ts).
   * Only pages with NO usable text layer enter the OCR path; every gate
   * failure keeps today's full-page-image fallback.
   */
  ocr?: OcrEngine
  /**
   * cell-data mode (xlsx exporter): the output carries data rows, not a page
   * facsimile, so bitmap fallbacks are pure loss. Keeps text covered by
   * vector-illustration regions (no rasterization) and ships text-bearing
   * pages instead of degrading them for layout-fidelity reasons.
   */
  cellData?: boolean
}

/** per-page conversion outcome (P4): lets callers surface degraded/scanned pages */
export interface PageResult {
  /** 1-based page number */
  page: number
  /** 'ok' = fully converted; 'degraded'/'scanned' = exported as a full-page
   * image; 'ocr' = scanned page recovered as editable text via local OCR */
  status: 'ok' | 'degraded' | 'scanned' | 'ocr'
  /** machine-readable degrade reason, e.g. 'bad-tounicode' | 'low-confidence' */
  reason?: string
  /** aggregated layout confidence (absent on scanned pages); on 'ocr' pages
   * this is the recognition confidence */
  confidence?: number
}

/** analyzed document IR plus the bookkeeping both rebuild layers need */
export interface IrDocument {
  irPages: IrPage[]
  warnings: string[]
  pageResults: PageResult[]
  furnitureHf: FurnitureHf[]
}

const DEGRADED_LABEL: Record<string, string> = {
  'bad-tounicode': 'unreliable text encoding (bad ToUnicode map)',
  rotated: 'rotated page',
  'vertical-text': 'vertical or rotated text',
  'low-confidence': 'low layout confidence',
  'content-lost': 'content could not be recovered',
  'graphics-lost': 'graphical content could not be recovered',
}

/** visible body chars a page must have before the empty-output guard applies */
const CONTENT_GUARD_MIN_CHARS = 10

/** cell-data mode: pages keeping at least this much IR text skip the
 * layout-fidelity degrades (low-confidence, graphics-lost) — data rows beat a
 * "not convertible" placeholder when there is real text to ship */
const CELL_DATA_KEEP_MIN_CHARS = 40

// ── graphics-loss guard (P29 E) ──
// Vector-art pages (book covers, full-page infographics) can sail through
// every text-level guard with high confidence while the IR keeps almost none
// of their ink: outline-glyph titles and logo art are ignored as stray vector
// paths. When the authored paint covers a real share of the page and the IR
// output covers almost none of it, the page ships as its bitmap instead.

/** authored ink must cover at least this page share before the guard applies */
const GRAPHICS_GUARD_MIN_AUTHORED = 0.3
/** emitted ink below this share of the authored ink degrades the page */
const GRAPHICS_GUARD_EMIT_SHARE = 0.35
/** every channel at/above this is paper-tone paint, not visible ink */
const GUARD_WHITE_MIN = 0xf2

function isNearWhiteHex(hex: string): boolean {
  if (!/^[0-9a-fA-F]{6}$/.test(hex)) return false
  return (
    parseInt(hex.slice(0, 2), 16) >= GUARD_WHITE_MIN &&
    parseInt(hex.slice(2, 4), 16) >= GUARD_WHITE_MIN &&
    parseInt(hex.slice(4, 6), 16) >= GUARD_WHITE_MIN
  )
}

/** authored visible-ink boxes: chars, images, and non-white opaque path fills */
function authoredInkBoxes(extracted: ExtractedPage): Rect[] {
  const boxes: Rect[] = []
  for (const c of extracted.chars) {
    if (!c.isGenerated && !c.invisible && c.text.trim() !== '') boxes.push(c.box)
  }
  for (const img of extracted.images) boxes.push(img.box)
  for (const p of extracted.paths) {
    if (!p.filled || (p.fillAlpha ?? 255) < 128 || isNearWhiteHex(p.fillColor)) continue
    for (const sub of p.subpaths) {
      if (sub.points.length < 3) continue
      const xs = sub.points.map((pt) => pt.x)
      const ys = sub.points.map((pt) => pt.y)
      boxes.push({
        x0: Math.min(...xs),
        y0: Math.min(...ys),
        x1: Math.max(...xs),
        y1: Math.max(...ys),
      })
    }
  }
  return boxes
}

/** ink boxes the rebuilt page actually paints. bgColor only MAYBE paints
 * (w:background needs a document-majority vote) and bgRender is only the
 * BACKGROUND stack — with the bake active the extracted paths/images already
 * exclude the baked objects, so neither earns page-wide credit here and the
 * guard still sees dropped FOREGROUND art on baked pages. */
function emittedInkBoxes(page: IrPage): Rect[] {
  if (page.render) return [{ x0: 0, y0: 0, x1: page.widthPt, y1: page.heightPt }]
  const boxes: Rect[] = []
  for (const b of page.blocks) boxes.push(b.box)
  for (const panel of page.bgPanels ?? []) boxes.push(panel.box)
  return boxes
}

/** non-whitespace text characters that actually made it into the page's IR */
function irTextCharCount(page: IrPage): number {
  let n = 0
  const addTextBlock = (b: { lines: Array<{ spans: Array<{ text: string }> }> }): void => {
    for (const line of b.lines) {
      for (const span of line.spans) n += span.text.replace(/\s+/g, '').length
    }
  }
  for (const block of page.blocks) {
    if (block.kind === 'text') addTextBlock(block)
    else if (block.kind === 'table') {
      for (const row of block.rows) {
        for (const cell of row) for (const tb of cell.blocks) addTextBlock(tb)
      }
    }
  }
  return n
}

// ── cross-page paragraph stitching (P32) ──
// A source page boundary that cuts MID-PARAGRAPH is flow evidence: the
// author's layout engine paginated a continuous paragraph. Rebuilding the two
// halves as separate paragraphs with an explicit page break amplifies any
// height drift on the earlier page into a stranded near-empty page (five
// spilled lines + w:pageBreakBefore = a page with five lines). Stitch the
// halves into one paragraph and let natural flow paginate; paragraph-aligned
// page starts keep their explicit break, re-syncing the drift.

/** the continuation's font size must stay within this ratio of the tail's */
const STITCH_SIZE_TOL = 1.1
/** left edges must agree within this many ems */
const STITCH_LEFT_TOL_EMS = 0.6
/** the tail line must reach this share of its block's widest line (unfinished) */
const STITCH_TAIL_FULL_RATIO = 0.9
/** the earlier page's content must span this share of the page height:
 * stitching an UNDERFULL page trades its (harmless) explicit break for a
 * page-count loss — natural flow packs the slack (8 source pages became 7) */
const STITCH_PREV_FILL_MIN = 0.72

const stitchableText = (b: PageBlock | undefined): TextBlock | null =>
  b !== undefined &&
  b.kind === 'text' &&
  b.list === undefined &&
  b.tocEntry === undefined &&
  b.cardId === undefined &&
  b.lines.length > 0
    ? b
    : null

const lineFontSizePt = (line: TextBlock['lines'][number]): number =>
  Math.max(...line.spans.map((s) => s.fontSize), 1)

const lineWidth = (line: TextBlock['lines'][number]): number => line.box.x1 - line.box.x0

export function stitchCrossPageParagraphs(pages: IrPage[]): void {
  for (let i = 1; i < pages.length; i++) {
    const prev = pages[i - 1]!
    const cur = pages[i]!
    if (prev.scanned || prev.degraded || prev.canvas) continue
    if (cur.scanned || cur.degraded || cur.canvas) continue
    // multi-column pages paginate per column — boundary evidence is ambiguous
    if ((prev.sections?.length ?? 0) > 1 || (cur.sections?.length ?? 0) > 1) continue
    if (prev.sections?.some((s) => s.columns.length > 1)) continue
    if (cur.sections?.some((s) => s.columns.length > 1)) continue
    const tail = stitchableText(prev.blocks[prev.blocks.length - 1])
    const head = stitchableText(cur.blocks[0])
    if (!tail || !head || tail.dir !== head.dir) continue
    // only near-full pages carry overflow risk worth trading the break for
    const tops = prev.blocks.map((b) => b.box.y1)
    const bottoms = prev.blocks.map((b) => b.box.y0)
    if (
      prev.heightPt <= 0 ||
      (Math.max(...tops) - Math.min(...bottoms)) / prev.heightPt < STITCH_PREV_FILL_MIN
    ) {
      continue
    }
    const tailLine = tail.lines[tail.lines.length - 1]!
    const headLine = head.lines[0]!
    const fontSize = lineFontSizePt(tailLine)
    const sizeRatio = fontSize / lineFontSizePt(headLine)
    if (sizeRatio > STITCH_SIZE_TOL || sizeRatio < 1 / STITCH_SIZE_TOL) continue
    // the tail line must look unfinished: as wide as the paragraph's widest line
    const widest = Math.max(...tail.lines.map(lineWidth))
    if (lineWidth(tailLine) < STITCH_TAIL_FULL_RATIO * widest) continue
    // the continuation starts flush with the tail paragraph's left edge
    if (Math.abs(tail.box.x0 - head.box.x0) > STITCH_LEFT_TOL_EMS * fontSize) continue
    if (head.firstLineIndentPt > STITCH_LEFT_TOL_EMS * fontSize) continue
    tail.stitchedFromLine = tail.lines.length
    tail.lines.push(...head.lines)
    cur.blocks.shift()
    // the sections tree references the same block objects — drop the head
    // there too, or the rebuild (which prefers sections) emits it twice
    for (const section of cur.sections ?? []) {
      for (const column of section.columns) {
        const at = column.blocks.indexOf(head)
        if (at >= 0) column.blocks.splice(at, 1)
      }
    }
    cur.flowsFromPrev = true
  }
}

/** most pages are scans → the caller should steer the user to an OCR flow */
export function isScannedDocument(pageResults: PageResult[], pages: number): boolean {
  const scannedPages = pageResults.filter((r) => r.status === 'scanned').length
  return pages > 0 && scannedPages > pages / 2
}

export function extractIrDocument(pdf: Uint8Array, opts: ConvertOptions): IrDocument {
  const m = opts.pdfium
  const irPages: IrPage[] = []
  const warnings: string[] = []
  const pageResults: PageResult[] = []
  let furnitureHf: FurnitureHf[] = []

  withPdfDocument(
    m,
    pdf,
    (doc) => {
      const total = m._FPDF_GetPageCount(doc)
      // a document that opens but exposes ZERO pages is unreadable in practice
      // (e.g. PDF 2.0 BrotliDecode page trees PDFium cannot decompress) — a
      // structured rejection beats silently writing an empty output file (P27)
      if (total <= 0) throw new PdfLoadError('corrupt', 0)
      const extractedPages: ExtractedPage[] = []
      for (let i = 0; i < total; i++) {
        extractedPages.push(
          extractPage(m, doc, i, {
            renderScale: opts.renderScale,
            rasterizeVectorRegions: !opts.cellData,
            cellData: opts.cellData,
          }),
        )
      }

      // P6: the source document renders its headers/footers/page numbers onto
      // every page — detected across pages and dropped (first occurrence kept)
      // so they do not repeat through the rebuilt body text
      const furniture = detectFurniture(
        extractedPages.map((p) => ({
          index: p.index,
          heightPt: p.heightPt,
          chars: p.chars,
          skip: p.scanned || p.degraded,
        })),
        // cell-data: repeated lines with no header/footer seat keep every
        // occurrence — deleting data breaks the xlsx promise
        { keepUnemitted: opts.cellData === true },
      )
      if (furniture.droppedLines > 0) {
        warnings.push(
          `headers/footers: ${furniture.droppedLines} repeated line(s) detected across pages and deduplicated`,
        )
      }
      furnitureHf = furniture.hf

      for (let i = 0; i < total; i++) {
        const extracted = extractedPages[i]!
        const dropSet = furniture.drop[i]!
        if (dropSet.size > 0) extracted.chars = extracted.chars.filter((c) => !dropSet.has(c))
        let page = analyzePage(extracted)

        // scanned page + an OCR engine: try to recover editable text; every
        // gate failure keeps the full-page-image fallback below unchanged.
        // Recognition reads a dedicated >=3 px/pt render (~216 dpi) — the
        // scale-2 fallback render costs measurable accuracy on small print.
        let ocrConfidence: number | undefined
        if (page.scanned && opts.ocr) {
          const hiRender =
            renderPageByIndexPng(m, doc, i, Math.max(opts.renderScale ?? 2, 3)) ?? undefined
          const recovered = tryOcrScannedPage(extracted, opts.ocr, hiRender)
          if (recovered) {
            page = recovered.page
            ocrConfidence = recovered.confidence
          }
        }

        const keepForCellData = (): boolean =>
          opts.cellData === true && irTextCharCount(page) >= CELL_DATA_KEEP_MIN_CHARS

        // P4 fidelity floor: a page whose aggregated confidence is too low ships
        // as its bitmap rather than as a garbled layout. OCR-recovered pages are
        // exempt like every other authored-ink guard (see IrPage.ocrRecovered):
        // the OCR engine applies its own acceptance gates, and re-degrading here
        // would clear the blocks to a bitmap while pageResults still reports the
        // page as editable 'ocr' text.
        if (
          !page.scanned &&
          !page.degraded &&
          !page.ocrRecovered &&
          page.confidence !== undefined &&
          page.confidence < PAGE_CONFIDENCE_MIN &&
          !keepForCellData()
        ) {
          page.degraded = true
          page.degradedReason = 'low-confidence'
          page.blocks = []
          page.sections = undefined
          page.shapes = undefined
          page.render = renderPageByIndexPng(m, doc, i, opts.renderScale ?? 2) ?? undefined
        }

        // P27 guard: a page whose visible body text was lost wholesale (in
        // extraction, or by analysis emitting zero text) must never ship as a
        // silently empty page — it degrades to its bitmap with a warning
        if (!page.scanned && !page.degraded) {
          const visibleBody = extracted.chars.filter(
            (c) => !c.isGenerated && !c.invisible && c.text.trim() !== '',
          ).length
          if (
            extracted.textLost ||
            (visibleBody >= CONTENT_GUARD_MIN_CHARS && irTextCharCount(page) === 0)
          ) {
            page.degraded = true
            page.degradedReason = 'content-lost'
            page.blocks = []
            page.sections = undefined
            page.shapes = undefined
            page.render = renderPageByIndexPng(m, doc, i, opts.renderScale ?? 2) ?? undefined
          }
        }

        // P29 E: vector-art pages whose ink the IR mostly dropped ship as bitmaps
        // (skipped for OCR-recovered pages: the "authored ink" is the scan
        // bitmap itself, which the recovered text intentionally supersedes)
        if (
          !page.scanned &&
          !page.degraded &&
          !page.ocrRecovered &&
          !extracted.ocrTextRecovered &&
          !keepForCellData()
        ) {
          const authoredCover = coverageRatio(
            authoredInkBoxes(extracted),
            page.widthPt,
            page.heightPt,
          )
          if (authoredCover >= GRAPHICS_GUARD_MIN_AUTHORED) {
            const emittedCover = coverageRatio(emittedInkBoxes(page), page.widthPt, page.heightPt)
            if (emittedCover < GRAPHICS_GUARD_EMIT_SHARE * authoredCover) {
              page.degraded = true
              page.degradedReason = 'graphics-lost'
              page.blocks = []
              page.sections = undefined
              page.shapes = undefined
              page.render = renderPageByIndexPng(m, doc, i, opts.renderScale ?? 2) ?? undefined
            }
          }
        }

        if (extracted.ocrTextRecovered) {
          warnings.push(
            extracted.ocrImageKept
              ? `page ${i + 1}: graphics-dominant searchable scan, exported as full-page image`
              : `page ${i + 1}: hidden OCR text layer recovered from scanned page image`,
          )
        }

        if (page.ocrRecovered) {
          warnings.push(
            `page ${i + 1}: scanned page, text recovered via local OCR (confidence ${(
              ocrConfidence ?? 0
            ).toFixed(2)})`,
          )
        } else if (page.scanned) {
          warnings.push(`page ${i + 1}: scanned page, exported as full-page image`)
        } else if (page.degraded) {
          const label =
            DEGRADED_LABEL[page.degradedReason ?? ''] ?? page.degradedReason ?? 'unknown'
          warnings.push(`page ${i + 1}: ${label}, exported as full-page image`)
        }
        if ((page.scanned || page.degraded) && !page.render) {
          warnings.push(`page ${i + 1}: fallback render failed, page content dropped`)
        }
        if (extracted.rotatedDropped) {
          warnings.push(
            `page ${i + 1}: ${extracted.rotatedDropped} rotated char(s) dropped (not representable in cells)`,
          )
        }
        if (extracted.vectorRegions.length > 0) {
          warnings.push(
            `page ${i + 1}: ${extracted.vectorRegions.length} vector illustration region(s) rendered as image`,
          )
        }
        if (page.shapes && page.shapes.ignoredPaths > 0) {
          warnings.push(
            `page ${i + 1}: ${page.shapes.ignoredPaths} stray vector path(s) ignored (curves/diagonals outside illustration regions)`,
          )
        }
        if (page.ignoredVerticalDecor) {
          warnings.push(
            `page ${i + 1}: ${page.ignoredVerticalDecor} decorative line(s) ignored (vertical or over the per-page cap)`,
          )
        }
        for (const w of page.warnings ?? []) warnings.push(`page ${i + 1}: ${w}`)

        pageResults.push({
          page: i + 1,
          status: page.ocrRecovered
            ? 'ocr'
            : page.scanned
              ? 'scanned'
              : page.degraded
                ? 'degraded'
                : 'ok',
          ...(page.degraded && page.degradedReason ? { reason: page.degradedReason } : {}),
          ...(page.ocrRecovered && ocrConfidence !== undefined
            ? { confidence: ocrConfidence }
            : page.confidence !== undefined
              ? { confidence: page.confidence }
              : {}),
        })
        irPages.push(page)
        opts.onProgress?.(i + 1, total)
      }

      // P19 canvas classifier: high-confidence slide pages leave the flow path
      // (their blocks emit as absolutely-positioned containers). Conservative
      // by design — document priors only lower the page gate, borderline pages
      // stay flow.
      classifyPages(irPages, readDocMetadata(m, doc))
      // cell-data (xlsx) promises page ↔ sheet: stitching would move a page's
      // continuation paragraph onto the previous sheet and leave its own empty
      if (!opts.cellData) stitchCrossPageParagraphs(irPages)
    },
    opts.password,
  )

  return { irPages, warnings, pageResults, furnitureHf }
}
