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
import type { IrPage } from './ir'

export interface ConvertOptions {
  /** initialized @embedpdf/pdfium module (see package header comment) */
  pdfium: PdfiumModule
  onProgress?: (page: number, total: number) => void
  /** raster scale for fallback page renders (pixels per point), default 2 */
  renderScale?: number
  /** user password for encrypted PDFs (P22); load failures throw PdfLoadError */
  password?: string
}

/** per-page conversion outcome (P4): lets callers surface degraded/scanned pages */
export interface PageResult {
  /** 1-based page number */
  page: number
  /** 'ok' = fully converted; 'degraded'/'scanned' = exported as a full-page image */
  status: 'ok' | 'degraded' | 'scanned'
  /** machine-readable degrade reason, e.g. 'bad-tounicode' | 'low-confidence' */
  reason?: string
  /** aggregated layout confidence (absent on scanned pages) */
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
        extractedPages.push(extractPage(m, doc, i, { renderScale: opts.renderScale }))
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
        const page = analyzePage(extracted)

        // P4 fidelity floor: a page whose aggregated confidence is too low ships
        // as its bitmap rather than as a garbled layout
        if (
          !page.scanned &&
          !page.degraded &&
          page.confidence !== undefined &&
          page.confidence < PAGE_CONFIDENCE_MIN
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
        if (!page.scanned && !page.degraded && !extracted.ocrTextRecovered) {
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

        if (page.scanned) {
          warnings.push(`page ${i + 1}: scanned page, exported as full-page image`)
        } else if (page.degraded) {
          const label =
            DEGRADED_LABEL[page.degradedReason ?? ''] ?? page.degradedReason ?? 'unknown'
          warnings.push(`page ${i + 1}: ${label}, exported as full-page image`)
        }
        if ((page.scanned || page.degraded) && !page.render) {
          warnings.push(`page ${i + 1}: fallback render failed, page content dropped`)
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
          status: page.scanned ? 'scanned' : page.degraded ? 'degraded' : 'ok',
          ...(page.degraded && page.degradedReason ? { reason: page.degradedReason } : {}),
          ...(page.confidence !== undefined ? { confidence: page.confidence } : {}),
        })
        irPages.push(page)
        opts.onProgress?.(i + 1, total)
      }

      // P19 canvas classifier: high-confidence slide pages leave the flow path
      // (their blocks emit as absolutely-positioned containers). Conservative
      // by design — document priors only lower the page gate, borderline pages
      // stay flow.
      classifyPages(irPages, readDocMetadata(m, doc))
    },
    opts.password,
  )

  return { irPages, warnings, pageResults, furnitureHf }
}
