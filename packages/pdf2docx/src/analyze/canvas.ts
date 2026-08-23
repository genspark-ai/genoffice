/**
 * P19 page classifier: flow | canvas | scan. Canvas pages (slides and other
 * absolutely-positioned designs) bypass the flow rebuild — their text blocks
 * emit as page-anchored containers (w:framePr) instead of stacking into a
 * paragraph flow the source never had.
 *
 * Decision discipline (Leaf-approved): the cost of a false canvas is an
 * editing disaster (a normal document turned into a pile of text boxes), the
 * cost of a miss is only the status quo — so document-level priors merely
 * LOWER the page-level gate, every borderline page stays flow, and non-slide
 * page geometry is a hard veto in Phase 1.
 */
import type { IrPage, PageBlock, TextBlock } from '../ir'
import { rectHeight, rectWidth } from '../geometry'

/** PDF Info dictionary fields relevant to the slide prior */
export interface DocMeta {
  producer?: string
  creator?: string
}

/** document-level priors: they tune the page gate, they never decide alone */
export interface CanvasDocPrior {
  slideProducer: boolean
  /** share of pages with typical slide geometry (16:9 / 4:3 landscape) */
  slideSizeShare: number
  /** median per-page character count (whitespace excluded) */
  medianPageChars: number
  /** share of page joints that read like one flowing paragraph */
  continuityShare: number
  /** page-level canvas points required (lower = stronger prior) */
  pointsNeeded: number
}

/** slide-app exports — a hit is the strongest document prior there is */
const SLIDE_PRODUCER_RE = /powerpoint|keynote|google slides|impress|wps\s*演示|wps\s+presentation/i

export function isSlideProducer(meta: DocMeta): boolean {
  return SLIDE_PRODUCER_RE.test(`${meta.producer ?? ''} ${meta.creator ?? ''}`)
}

/** typical slide aspect ratios; letter (1.294) and A4 (1.415) landscape stay outside */
const SLIDE_ASPECTS = [16 / 9, 4 / 3]
const SLIDE_ASPECT_TOL = 0.02
/** slides are at least 8in wide (PowerPoint's smallest common export) */
const SLIDE_MIN_WIDTH_PT = 576

export function isSlideSizedPage(widthPt: number, heightPt: number): boolean {
  if (widthPt <= heightPt || widthPt < SLIDE_MIN_WIDTH_PT) return false
  const aspect = widthPt / heightPt
  return SLIDE_ASPECTS.some((a) => Math.abs(aspect - a) / a <= SLIDE_ASPECT_TOL)
}

/** non-whitespace character count across the page's text AND table blocks */
function pageCharCount(page: IrPage): number {
  let n = 0
  const countText = (b: TextBlock) => {
    for (const line of b.lines) for (const s of line.spans) n += s.text.replace(/\s/g, '').length
  }
  for (const b of page.blocks) {
    if (b.kind === 'text') countText(b)
    else if (b.kind === 'table') {
      for (const row of b.rows) for (const cell of row) cell.blocks.forEach(countText)
    }
  }
  return n
}

/** char-weighted median font size over the page's text blocks (0 = no text) */
function weightedMedianFontPt(page: IrPage): number {
  const buckets: Array<{ size: number; chars: number }> = []
  for (const b of page.blocks) {
    if (b.kind !== 'text') continue
    for (const line of b.lines) {
      for (const s of line.spans) {
        const chars = s.text.replace(/\s/g, '').length
        if (chars > 0) buckets.push({ size: s.fontSize, chars })
      }
    }
  }
  if (buckets.length === 0) return 0
  buckets.sort((a, b) => a.size - b.size)
  const half = buckets.reduce((t, b) => t + b.chars, 0) / 2
  let acc = 0
  for (const b of buckets) {
    acc += b.chars
    if (acc >= half) return b.size
  }
  return buckets[buckets.length - 1]!.size
}

const isBehindFloat = (b: PageBlock): boolean => b.kind === 'image' && b.float !== undefined

/** rough union area of the page's background layers (grid-sampled, 0..1 of page) */
function backgroundCoverShare(page: IrPage): number {
  if (page.bgRender) return 1
  const boxes = [
    ...(page.bgPanels ?? []).map((p) => p.box),
    ...page.blocks.filter(isBehindFloat).map((b) => b.box),
  ]
  if (boxes.length === 0) return 0
  const N = 24
  let covered = 0
  for (let gy = 0; gy < N; gy++) {
    for (let gx = 0; gx < N; gx++) {
      const x = ((gx + 0.5) / N) * page.widthPt
      const y = ((gy + 0.5) / N) * page.heightPt
      if (boxes.some((b) => x >= b.x0 && x <= b.x1 && y >= b.y0 && y <= b.y1)) covered++
    }
  }
  return covered / (N * N)
}

const overlapArea = (a: TextBlock['box'], b: TextBlock['box']): number =>
  Math.max(0, Math.min(a.x1, b.x1) - Math.max(a.x0, b.x0)) *
  Math.max(0, Math.min(a.y1, b.y1) - Math.max(a.y0, b.y0))

// ── page-level gates ──
/** hard vetoes: any page beyond these is body text, never canvas */
const VETO_MAX_CHARS = 1000
const VETO_MAX_TEXT_LINES = 40
/** canvas points */
const SPARSE_MAX_CHARS = 600
const DISPLAY_MEDIAN_FONT_PT = 13.5
const DISPLAY_MAX_FONT_PT = 26
const BG_COVER_MIN_SHARE = 0.3
const ISLANDS_MIN_BLOCKS = 3
const ISLANDS_MAX_INK_SHARE = 0.35
const OVERLAP_MIN_SHARE = 0.3

/** page-level canvas evidence points (0..5) */
function canvasPoints(page: IrPage): number {
  const textBlocks = page.blocks.filter((b): b is TextBlock => b.kind === 'text')
  const chars = pageCharCount(page)
  let points = 0
  if (chars <= SPARSE_MAX_CHARS) points++
  const median = weightedMedianFontPt(page)
  const maxFont = Math.max(
    0,
    ...textBlocks.flatMap((b) => b.lines.flatMap((l) => l.spans.map((s) => s.fontSize))),
  )
  if (median >= DISPLAY_MEDIAN_FONT_PT || maxFont >= DISPLAY_MAX_FONT_PT) points++
  if (page.bgColor !== undefined || backgroundCoverShare(page) >= BG_COVER_MIN_SHARE) {
    points++
  }
  const pageArea = page.widthPt * page.heightPt
  const inkShare =
    textBlocks.reduce((t, b) => t + rectWidth(b.box) * rectHeight(b.box), 0) / pageArea
  if (textBlocks.length >= ISLANDS_MIN_BLOCKS && inkShare <= ISLANDS_MAX_INK_SHARE) points++
  const overlays = page.blocks.filter(isBehindFloat).map((b) => b.box)
  const overlapsArt = textBlocks.some((b) => {
    const own = rectWidth(b.box) * rectHeight(b.box)
    return own > 0 && overlays.some((o) => overlapArea(b.box, o) >= OVERLAP_MIN_SHARE * own)
  })
  if (overlapsArt || (page.warnings ?? []).some((w) => w.startsWith('overlapping blocks'))) {
    points++
  }
  return points
}

// ── cross-page paragraph continuity (a flow fingerprint slides never show) ──
/** a joint "continues" when the page-ending line still fills its block and stops mid-sentence */
const CONTINUITY_MIN_LINE_FILL = 0.8
const CONTINUITY_MAX_FONT_PT = 14.5
const TERMINAL_PUNCT_RE = /[.。．!！?？…;；:：]["'"』」)】〕]?\s*$/

function jointContinues(prev: IrPage, next: IrPage): boolean {
  const lastText = [...prev.blocks].reverse().find((b): b is TextBlock => b.kind === 'text')
  const firstText = next.blocks.find((b): b is TextBlock => b.kind === 'text')
  if (!lastText || !firstText) return false
  if (lastText.lines.length < 2) return false
  const lastLine = lastText.lines[lastText.lines.length - 1]!
  const lastSpan = lastLine.spans[lastLine.spans.length - 1]
  if (!lastSpan || lastSpan.fontSize > CONTINUITY_MAX_FONT_PT) return false
  if (TERMINAL_PUNCT_RE.test(lastSpan.text)) return false
  const fill = rectWidth(lastLine.box) / Math.max(1, rectWidth(lastText.box))
  if (fill < CONTINUITY_MIN_LINE_FILL) return false
  const nextSpan = firstText.lines[0]?.spans[0]
  if (!nextSpan) return false
  // the continuation must look like the same body text, not a fresh heading
  return Math.abs(nextSpan.fontSize - lastSpan.fontSize) <= 2
}

// ── document prior → page gate threshold ──
const POINTS_STRONG = 2
const POINTS_WEAK = 3
const POINTS_NONE = 4
/** geometry-only prior needs a doc that reads like slides overall */
const WEAK_PRIOR_SIZE_SHARE = 0.9
const WEAK_PRIOR_MAX_MEDIAN_CHARS = 700
const WEAK_PRIOR_MAX_CONTINUITY = 0.25
/** flowing joints beyond this poison even a producer prior */
const STRONG_PRIOR_MAX_CONTINUITY = 0.3

export function computeCanvasPrior(pages: IrPage[], meta: DocMeta): CanvasDocPrior {
  const flowPages = pages.filter((p) => !p.scanned && !p.degraded)
  const slideSized = flowPages.filter((p) => isSlideSizedPage(p.widthPt, p.heightPt))
  const slideSizeShare = flowPages.length > 0 ? slideSized.length / flowPages.length : 0
  const counts = flowPages.map(pageCharCount).sort((a, b) => a - b)
  const medianPageChars = counts.length > 0 ? counts[Math.floor(counts.length / 2)]! : 0
  let joints = 0
  let continuing = 0
  for (let i = 0; i + 1 < pages.length; i++) {
    const prev = pages[i]!
    const next = pages[i + 1]!
    if (prev.scanned || prev.degraded || next.scanned || next.degraded) continue
    joints++
    if (jointContinues(prev, next)) continuing++
  }
  const continuityShare = joints > 0 ? continuing / joints : 0
  const slideProducer = isSlideProducer(meta)

  let pointsNeeded = POINTS_NONE
  if (slideProducer) {
    pointsNeeded = continuityShare > STRONG_PRIOR_MAX_CONTINUITY ? POINTS_WEAK : POINTS_STRONG
  } else if (
    slideSizeShare >= WEAK_PRIOR_SIZE_SHARE &&
    medianPageChars <= WEAK_PRIOR_MAX_MEDIAN_CHARS &&
    continuityShare <= WEAK_PRIOR_MAX_CONTINUITY
  ) {
    pointsNeeded = POINTS_WEAK
  }
  return { slideProducer, slideSizeShare, medianPageChars, continuityShare, pointsNeeded }
}

export type PageLayoutClass = 'flow' | 'canvas' | 'scan'

export function classifyPage(page: IrPage, prior: CanvasDocPrior): PageLayoutClass {
  if (page.scanned || page.degraded) return 'scan'
  // non-slide geometry: only the document-gated newsletter path (P20 C,
  // classifyPages) may admit such pages — never the slide points system
  if (!isSlideSizedPage(page.widthPt, page.heightPt)) return 'flow'
  if (pageCharCount(page) > VETO_MAX_CHARS) return 'flow'
  const textLines = page.blocks.reduce((t, b) => t + (b.kind === 'text' ? b.lines.length : 0), 0)
  if (textLines > VETO_MAX_TEXT_LINES) return 'flow'
  return canvasPoints(page) >= prior.pointsNeeded ? 'canvas' : 'flow'
}

// ── newsletter pages (P20 C) ──
// Full-page magazine/bulletin layouts: several absolutely-placed regions,
// non-linear reading order. The give-away the flow model cannot fake is the
// cross-SECTION spacing chain running BACKWARDS by a large share of the page
// (the next region's content starts far ABOVE the previous region's bottom —
// XY-cut had to interleave bands to cover the layout). Admission is
// document-gated: one STRONG page unlocks canvas for the document's other
// region-structured pages, and ordinary documents (reports / gov papers / forms /
// contracts) never produce a strong page — the worst cross-section backtrack
// over the whole 129-sample corpus outside lo-tSC is 135pt ≈ 0.16 of its
// page, far under the 0.4 gate.

/** strong page: backtrack at least this share of the page height */
const NEWSLETTER_CROSS_RATIO = 0.4
const NEWSLETTER_MIN_SECTIONS = 2
/** weak page: total columns across sections (region grid, not one text river) */
const NEWSLETTER_MIN_COLUMNS = 4
/** real content pages only — image/crop test artifacts have almost no text */
const NEWSLETTER_MIN_CHARS = 100
/** …and a page denser than this reads as a report, not a bulletin */
const NEWSLETTER_MAX_CHARS = 4000
const NEWSLETTER_MAX_TEXT_LINES = 80
/** a lone strong page in a long document is noise, not a bulletin */
const NEWSLETTER_MIN_STRONG_SHARE = 0.2

/** worst cross-section backtrack (pt): how far a section's first block starts
 * ABOVE the previous section's bottom (mirrors applySpacingChain's joints) */
function crossSectionBacktrackPt(page: IrPage): number {
  let worst = 0
  let prevBottom: number | null = null
  for (const s of page.sections ?? []) {
    const flow = s.columns[0]?.blocks.filter((b) => !isBehindFloat(b)) ?? []
    const first = flow[0]
    if (prevBottom !== null && first !== undefined) {
      worst = Math.max(worst, first.box.y1 - prevBottom)
    }
    prevBottom = s.box.y0
  }
  return worst
}

/** shared newsletter-page shape: regioned, mixed with tables, bulletin-dense */
function newsletterPageShape(page: IrPage): boolean {
  if (page.scanned || page.degraded) return false
  if (isSlideSizedPage(page.widthPt, page.heightPt)) return false
  const secs = page.sections ?? []
  if (secs.length < NEWSLETTER_MIN_SECTIONS) return false
  if (secs.reduce((n, s) => n + s.columns.length, 0) < NEWSLETTER_MIN_COLUMNS) return false
  if (!page.blocks.some((b) => b.kind === 'table')) return false
  const chars = pageCharCount(page)
  if (chars < NEWSLETTER_MIN_CHARS || chars > NEWSLETTER_MAX_CHARS) return false
  const textLines = page.blocks.reduce((t, b) => t + (b.kind === 'text' ? b.lines.length : 0), 0)
  return textLines <= NEWSLETTER_MAX_TEXT_LINES
}

/** strong page: region shape AND the layout-model-collapsing backtrack */
export function isNewsletterStrongPage(page: IrPage): boolean {
  if (!newsletterPageShape(page)) return false
  if (!(page.sections ?? []).some((s) => s.columns.length >= 2)) return false
  return crossSectionBacktrackPt(page) >= NEWSLETTER_CROSS_RATIO * page.heightPt
}

/** classify every page of one document, marking canvas pages on the IR */
export function classifyPages(pages: IrPage[], meta: DocMeta): CanvasDocPrior {
  const prior = computeCanvasPrior(pages, meta)
  for (const page of pages) {
    if (classifyPage(page, prior) === 'canvas') page.canvas = true
  }
  // newsletter admission (P20 C): a strong page unlocks the document's other
  // region-structured pages; without one, nothing changes
  const live = pages.filter((p) => !p.scanned && !p.degraded)
  const strong = live.filter(isNewsletterStrongPage)
  if (strong.length > 0 && strong.length >= NEWSLETTER_MIN_STRONG_SHARE * live.length) {
    for (const page of live) {
      if (page.canvas !== true && newsletterPageShape(page)) page.canvas = true
    }
  }
  return prior
}
