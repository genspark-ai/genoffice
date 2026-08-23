/**
 * Page furniture detection (P6) — document-level pass over the EXTRACTED
 * pages, before per-page analysis. PDF pages repeat the source document's
 * headers/footers and page numbers on every page; converted back to body
 * text they would appear N times where the source has them (at most) once.
 *
 * Two rules, both restricted to the page-top/page-bottom bands:
 *  - repeated lines: same normalized text at the same edge distance on
 *    enough pages → header/footer; the FIRST occurrence stays (content is
 *    never fully dropped), the repeats go;
 *  - page numbers: digit/roman-only lines (with decorations, "Page N of M"
 *    style included) at a shared edge distance on ≥2 pages → all go.
 *
 * Text appearing at the top of only one or two pages is content, not a
 * header — the repeat threshold scales with the page count.
 */
import type { PdfChar } from '../ir'
import { groupIntoLines, isSpaceCode } from './lines'

/** a line is header/footer material only within this distance of the page edge */
const BAND_PT = 96
/** repeated lines need max(this, REPEAT_RATIO × text pages) occurrences */
const REPEAT_MIN_PAGES = 3
const REPEAT_RATIO = 0.2
/** page-number lines need this many pages with one at the same edge distance */
const PAGENO_MIN_PAGES = 2
/** edge distances within this tolerance (pt) are the same header/footer slot */
const EDGE_TOL_PT = 6
/**
 * lines set below this size are hidden micro text (SEO/watermark junk baked
 * invisibly into every page of template decks), not headers or footers — a
 * reader never sees them, so "repeated furniture" reasoning does not apply;
 * they stay in the body, as invisible in the DOCX as in the source (P11 C)
 */
const FURNITURE_MIN_FONT_PT = 2.5

/** the caller's page shape (subset of ExtractedPage; scanned/degraded pages skip) */
export interface FurniturePage {
  index: number
  heightPt: number
  chars: readonly PdfChar[]
  skip?: boolean
}

const isVisible = (c: PdfChar): boolean => !isSpaceCode(c.code) && c.code > 0x1f

/**
 * Repeat keys keep their digits: Japanese "Form No.1" and "Form No.2" captions are DIFFERENT
 * content lines, not one repeating header (digit-normalized keys merged them
 * and dropped real text). True headers repeat verbatim; headers with varying
 * numbers in them are the page-number rule's business.
 */
const normalize = (text: string): string => text.toLowerCase().replace(/\s+/g, ' ').trim()

/** "5" / "- 12 -" / "— 7 —" / "iv" (bare number + decorations) */
const PAGENO_RE = /^[-—–‒―−‐·.()[\]/\s]*(?:[0-9０-９]+|[ivxlcdm]+)[-—–‒―−‐·.()[\]/\s]*$/i
/** …or "Page 3 of 20" and its CJK "page X of Y" equivalents: digits stripped, only marker words left */
const PAGENO_WORDS_RE = /^(?:page|of|pg|p|第|页|頁|共|ページ|стр|из)+$/i
const NON_LETTER_RE = /[^\p{L}]+/gu

const isPageNumberLine = (text: string): boolean => {
  const trimmed = text.trim()
  if (PAGENO_RE.test(trimmed)) return true
  const letters = trimmed.replace(/[0-9０-９]+/g, '').replace(NON_LETTER_RE, '')
  return letters.length > 0 && PAGENO_WORDS_RE.test(letters)
}

interface BandLine {
  page: number
  /** distance from the page edge the band hugs */
  edgeDist: number
  chars: PdfChar[]
  norm: string
  pageNo: boolean
}

interface Slot {
  lines: BandLine[]
  pages: Set<number>
}

/** group band lines by normalized text, then cluster by edge distance */
function slotsOf(lines: BandLine[]): Slot[] {
  const byText = new Map<string, BandLine[]>()
  for (const line of lines) {
    let list = byText.get(line.norm)
    if (!list) byText.set(line.norm, (list = []))
    list.push(line)
  }
  const slots: Slot[] = []
  for (const group of byText.values()) {
    const sorted = [...group].sort((a, b) => a.edgeDist - b.edgeDist)
    let current: BandLine[] = []
    for (const line of sorted) {
      const anchor = current[0]
      if (anchor && line.edgeDist - anchor.edgeDist > EDGE_TOL_PT) {
        slots.push({ lines: current, pages: new Set(current.map((l) => l.page)) })
        current = []
      }
      current.push(line)
    }
    if (current.length > 0) {
      slots.push({ lines: current, pages: new Set(current.map((l) => l.page)) })
    }
  }
  return slots
}

/** docx PAGE-field placeholder (mirrors docx-engine PAGE_MARK '') */
export const HF_PAGE_MARK = '\uE001'
/** an emitted slot's text must stay a sane header line */
const HF_MAX_TEXT_LEN = 120
/** at most this many re-emitted lines per band */
const HF_MAX_PER_BAND = 2

/**
 * A dropped furniture slot re-emittable as a real docx header/footer line
 * (P17). Page-number slots carry HF_PAGE_MARK where the running number sat.
 */
export interface FurnitureHf {
  band: 'top' | 'bottom'
  /** representative text; HF_PAGE_MARK marks the PAGE field position */
  text: string
  pageNo: boolean
  fontSizePt: number
  fontFamily: string
  bold: boolean
  italic: boolean
  /** hex RRGGBB */
  color: string
  /** representative ink span (pt, page coords) */
  x0: number
  x1: number
  /** baseline distance from the hugged page edge (pt) */
  edgeDistPt: number
  /** the slot repeats on the document's first text page too */
  coversFirstPage: boolean
}

export interface FurnitureResult {
  /** chars to drop, per input page (same order); empty sets where nothing matched */
  drop: Array<Set<PdfChar>>
  /** dropped line count (for the conversion warning) */
  droppedLines: number
  /** slots re-emittable as real docx headers/footers (P17) */
  hf: FurnitureHf[]
}

/** raw line text including inner spaces, trimmed at the ends */
const lineText = (chars: readonly PdfChar[]): string =>
  chars
    .map((c) => c.text)
    .join('')
    .replace(/\s+/g, ' ')
    .trim()

/** digit runs (ASCII value, index, length) of a line's text; fullwidth folded */
function digitRuns(text: string): Array<{ value: number; index: number; length: number }> {
  const out: Array<{ value: number; index: number; length: number }> = []
  for (const m of text.matchAll(/[0-9０-９]+/g)) {
    const ascii = [...m[0]]
      .map((d) => String.fromCharCode(d.charCodeAt(0) - (d >= '０' ? 0xfee0 : 0)))
      .join('')
    out.push({ value: parseInt(ascii, 10), index: m.index!, length: m[0].length })
  }
  return out
}

const HAS_RTL_RE = /[\u0590-\u08FF\uFB1D-\uFDFF\uFE70-\uFEFF]/

/** style + geometry of a representative band line → FurnitureHf shell */
function hfShell(
  band: 'top' | 'bottom',
  line: BandLine,
  coversFirstPage: boolean,
): Omit<FurnitureHf, 'text' | 'pageNo'> | null {
  const visible = line.chars.filter(isVisible)
  const first = visible[0]
  if (!first) return null
  return {
    band,
    fontSizePt: first.fontSize,
    fontFamily: first.fontFamily,
    bold: first.fontWeight >= 600,
    italic: first.italic,
    color: first.color,
    x0: Math.min(...visible.map((c) => c.box.x0)),
    x1: Math.max(...visible.map((c) => c.box.x1)),
    edgeDistPt: line.edgeDist,
    coversFirstPage,
  }
}

/**
 * A page-number slot as an HF entry: the digit run that counts the physical
 * page (value === page+1 on every probe) becomes HF_PAGE_MARK. Slots whose
 * numbers do not track the physical page (offset numbering, roman) stay
 * un-emittable — reproducing them wrong is worse than omitting them.
 */
function pageNoHf(band: 'top' | 'bottom', slot: Slot, firstTextPage: number): FurnitureHf | null {
  const samples = [...slot.lines].sort((a, b) => a.page - b.page)
  const probes = [
    samples[0]!,
    samples[Math.floor(samples.length / 2)]!,
    samples[samples.length - 1]!,
  ]
  const runsPer = probes.map((l) => ({
    l,
    text: lineText(l.chars),
    runs: digitRuns(lineText(l.chars)),
  }))
  const n = runsPer[0]!.runs.length
  if (n === 0 || runsPer.some((r) => r.runs.length !== n)) return null
  let k = -1
  for (let i = 0; i < n; i++) {
    if (runsPer.every(({ l, runs }) => runs[i]!.value === l.page + 1)) {
      k = i
      break
    }
  }
  if (k < 0) return null
  const rep = runsPer[0]!
  if (rep.text.length > HF_MAX_TEXT_LEN || HAS_RTL_RE.test(rep.text)) return null
  const run = rep.runs[k]!
  const text = rep.text.slice(0, run.index) + HF_PAGE_MARK + rep.text.slice(run.index + run.length)
  const shell = hfShell(band, rep.l, slot.pages.has(firstTextPage))
  return shell ? { ...shell, text, pageNo: true } : null
}

/** a verbatim-repeated slot as an HF entry (representative = second occurrence) */
function textHf(band: 'top' | 'bottom', slot: Slot, firstTextPage: number): FurnitureHf | null {
  const ordered = [...slot.lines].sort((a, b) => a.page - b.page)
  const rep = ordered[1] ?? ordered[0]!
  const text = lineText(rep.chars)
  if (!text || text.length > HF_MAX_TEXT_LEN || HAS_RTL_RE.test(text)) return null
  const shell = hfShell(band, rep, slot.pages.has(firstTextPage))
  return shell ? { ...shell, text, pageNo: false } : null
}

/** Detect repeated headers/footers + page numbers across a document's pages. */
export function detectFurniture(pages: readonly FurniturePage[]): FurnitureResult {
  const drop: Array<Set<PdfChar>> = pages.map(() => new Set())
  const textPages = pages
    .map((page, at) => ({ page, at }))
    .filter(({ page }) => !page.skip && page.chars.length > 0)
  let droppedLines = 0
  const hf: FurnitureHf[] = []
  if (textPages.length < PAGENO_MIN_PAGES) return { drop, droppedLines, hf }

  const top: BandLine[] = []
  const bottom: BandLine[] = []
  for (const { page, at } of textPages) {
    for (const raw of groupIntoLines(page.chars)) {
      const text = raw.chars
        .filter(isVisible)
        .map((c) => c.text)
        .join('')
      const norm = normalize(text)
      if (!norm) continue
      const visible = raw.chars.filter(isVisible)
      if (Math.max(...visible.map((c) => c.fontSize)) < FURNITURE_MIN_FONT_PT) continue
      const line = { page: at, chars: raw.chars, norm, pageNo: isPageNumberLine(text) }
      const fromTop = page.heightPt - raw.baseline
      if (fromTop <= BAND_PT) top.push({ ...line, edgeDist: fromTop })
      else if (raw.baseline <= BAND_PT) bottom.push({ ...line, edgeDist: raw.baseline })
    }
  }

  const repeatMin = Math.max(REPEAT_MIN_PAGES, Math.ceil(REPEAT_RATIO * textPages.length))
  const firstTextPage = textPages[0]!.at
  for (const [bandName, band] of [
    ['top', top],
    ['bottom', bottom],
  ] as const) {
    let emitted = 0
    // page-number slots first: all page numbers share one normalized key ('#')
    // only per decoration style, so cluster them together regardless of text
    const numbered = band.filter((l) => l.pageNo)
    for (const slot of slotsOf(numbered.map((l) => ({ ...l, norm: '#pageno' })))) {
      if (slot.pages.size < PAGENO_MIN_PAGES) continue
      for (const line of slot.lines) {
        for (const c of line.chars) drop[line.page]!.add(c)
        droppedLines++
      }
      // a slot only re-emits when it covers most pages — a two-page fluke
      // must not stamp a page number onto the whole document
      if (slot.pages.size >= repeatMin && emitted < HF_MAX_PER_BAND) {
        const entry = pageNoHf(bandName, slot, firstTextPage)
        if (entry) {
          hf.push(entry)
          emitted++
        }
      }
    }
    for (const slot of slotsOf(band.filter((l) => !l.pageNo))) {
      if (slot.pages.size < repeatMin) continue
      const entry = emitted < HF_MAX_PER_BAND ? textHf(bandName, slot, firstTextPage) : null
      const ordered = [...slot.lines].sort((a, b) => a.page - b.page)
      if (entry) {
        // the slot re-renders as a real header/footer on EVERY page — the
        // first occurrence must go too, or its page shows the line twice
        hf.push(entry)
        emitted++
        for (const line of ordered) {
          for (const c of line.chars) drop[line.page]!.add(c)
          droppedLines++
        }
      } else {
        // keep the first occurrence — headers usually carry real content once
        for (const line of ordered.slice(1)) {
          for (const c of line.chars) drop[line.page]!.add(c)
          droppedLines++
        }
      }
    }
  }
  return { drop, droppedLines, hf }
}
