/**
 * XY-Cut column / section detection — pdf2docx rule 4, extended past 2
 * columns. The page (minus header/footer candidate bands, tables and floats)
 * is swept top → bottom in y-slabs; a section is a maximal run of slabs
 * sharing at least one INTERNAL x-projection valley (a gutter). Full-width
 * content kills every gutter and closes the section; a fresh internal gap
 * after a gutterless run opens one (title-above-columns pages). Sections
 * without a surviving gutter are single-column.
 *
 * Column count is free (1/2/3+), widths need not be equal. Reading order is
 * top → bottom inside a column; across columns LTR pages go left → right and
 * RTL-dominant sections right → left (first-strong, like P2 paragraphs).
 */
import type { Interval, Rect } from '../geometry'
import {
  complementIntervals,
  intersectIntervals,
  median,
  mergeIntervals,
  rectUnion,
  rectUnionAll,
} from '../geometry'
import type { Dir, ImageBlock, TableBlock, TextBlock } from '../ir'
import { isRtlScript } from '../script'
import type { LineUnit } from './units'

// ── tunable gates ──

/** a gutter (persistent internal valley) must be at least this wide (pt)… */
const GUTTER_MIN_PT = 9
/**
 * …and much wider on landscape (slide) pages: design spacing between display
 * elements (a letter-spaced title's inner gap, a badge's air) runs 10–20pt
 * there, while real slide column gutters are generous (~50pt). The document
 * threshold minted micro-columns out of data-viz slides.
 */
const GUTTER_MIN_PT_LANDSCAPE = 24
/** …and at least this many ems of the section's median font size */
const GUTTER_MIN_EMS = 0.6
/** top/bottom share of the page height scanned for header/footer candidates */
const HF_BAND_RATIO = 0.08
/**
 * …wider on landscape (slide) pages: deck furniture (chapter labels, page
 * numbers) sits deeper than a document running head, and an unpeeled label
 * donates flank cover that turns the body's margins into phantom gutters
 */
const HF_BAND_RATIO_LANDSCAPE = 0.12
/** band content only peels off when this many median line heights above/below the body */
const HF_GAP_LINE_HEIGHTS = 1.5
/** a multi-column section needs at least this many elements per column */
const MIN_COLUMN_ELEMENTS = 3
/**
 * …but a column separated by a VAST gutter (this many × gutterMin) is spaced
 * apart on purpose and stands alone with just 2 elements (P15 B): a photo
 * caption right of a slide's stats row sits behind a ~250pt gap — merging it
 * into the nearest stats column runs the caption into that column's label
 * text. Single-element bands (a card's "02" badge, P14 A) still merge.
 */
const STANDALONE_GUTTER_FACTOR = 3
const STANDALONE_MIN_ELEMENTS = 2
/**
 * …and every column's CONTENT must be at least this wide (pt). Bullet glyphs
 * separated from their item text by the word gap look like a 3pt-wide
 * "column" of markers — a sliver column is never a real page column.
 */
const MIN_COLUMN_WIDTH_PT = 24

/** one flow element: a text unit, or a table/image block placed in the flow */
export interface SectionElement {
  box: Rect
  unit?: LineUnit
  block?: TableBlock | ImageBlock | TextBlock
}

export interface LayoutColumn {
  box: Rect
  elements: SectionElement[]
}

export interface LayoutSection {
  box: Rect
  /** columns in LAYOUT order (left → right); reading order comes from `dir` */
  columns: LayoutColumn[]
  /** surviving internal gutters, layout order */
  gutters: Interval[]
  /** first-strong direction over the section's text (column reading order) */
  dir: Dir
}

interface Slab {
  top: number
  bottom: number
  cover: Interval[]
  gaps: Interval[]
}

/** gaps flanked by coverage on both sides (candidate gutters, not margins) */
function internalGaps(gaps: readonly Interval[], cover: readonly Interval[]): Interval[] {
  return gaps.filter(
    (g) => cover.some((c) => c.hi <= g.lo + 0.01) && cover.some((c) => c.lo >= g.hi - 0.01),
  )
}

function slabsOf(elements: readonly SectionElement[], bodyLeft: number, bodyRight: number): Slab[] {
  const ys = [...new Set(elements.flatMap((e) => [e.box.y0, e.box.y1]))].sort((a, b) => b - a)
  const slabs: Slab[] = []
  for (let i = 1; i < ys.length; i++) {
    const top = ys[i - 1]!
    const bottom = ys[i]!
    const active = elements.filter((e) => e.box.y1 > bottom && e.box.y0 < top)
    if (active.length === 0) continue // pure whitespace constrains nothing
    const cover = mergeIntervals(active.map((e) => ({ lo: e.box.x0, hi: e.box.x1 })))
    slabs.push({ top, bottom, cover, gaps: complementIntervals(cover, bodyLeft, bodyRight) })
  }
  return slabs
}

interface OpenSection {
  slabs: Slab[]
  gaps: Interval[]
  cover: Interval[]
}

/**
 * a gutterless run at least this many line heights tall is a standalone
 * block (title/quote stack), not the staggered head of a column — fresh
 * column structure appearing beside it closes the run instead of merging
 */
const FRESH_STRUCTURE_MIN_RUN_LINES = 2.5

/** sweep slabs into maximal gutter-coherent runs (see module header) */
function sweepSections(
  slabs: Slab[],
  gutterMin: number,
  minRunHeightPt: number,
): Array<{ top: number; bottom: number }> {
  const ranges: Array<{ top: number; bottom: number }> = []
  let open: OpenSection | null = null

  const close = (o: OpenSection): void => {
    ranges.push({ top: o.slabs[0]!.top, bottom: o.slabs[o.slabs.length - 1]!.bottom })
  }

  for (const slab of slabs) {
    if (!open) {
      open = { slabs: [slab], gaps: slab.gaps, cover: slab.cover }
      continue
    }
    const tryGaps = intersectIntervals(open.gaps, slab.gaps).filter((g) => g.hi - g.lo >= gutterMin)
    const combinedCover = mergeIntervals([...open.cover, ...slab.cover])
    const internalTry = internalGaps(tryGaps, combinedCover)
    const openInternal = internalGaps(
      open.gaps.filter((g) => g.hi - g.lo >= gutterMin),
      open.cover,
    )
    const slabInternal = internalGaps(
      slab.gaps.filter((g) => g.hi - g.lo >= gutterMin),
      slab.cover,
    )
    // fresh column structure beside a TALL gutterless run (P14 A): the shared
    // gap counts as internal only because the INCOMING slab supplies the far
    // flank (a title/quote stack ending at the gutter's left edge, card
    // columns starting below). Merging glues the stack into one column and
    // steals rows from the other; close instead. Thin runs (staggered column
    // heads, a lone title line's top sliver) keep merging.
    const runHeightPt = open.slabs[0]!.top - open.slabs[open.slabs.length - 1]!.bottom
    const freshStructure =
      internalTry.length > 0 &&
      internalGaps(internalTry, open.cover).length === 0 &&
      runHeightPt >= minRunHeightPt
    // a slab that KILLS one of the run's internal gutters closes the section
    // even when another gutter survives (P14 A): a full-width closing line
    // that happens to start just right of a number-column's narrow pseudo-
    // gutter must not ride into the section on that sliver and void the real
    // column split. Narrowed gutters still count as surviving.
    const killedGutter = openInternal.some(
      (g) => !internalTry.some((t) => t.lo < g.hi && t.hi > g.lo),
    )
    // keep merging while a shared gutter survives, or while both sides are
    // plain single-column; a lost gutter or a fresh one closes the section
    if (
      !freshStructure &&
      !killedGutter &&
      (internalTry.length > 0 || (openInternal.length === 0 && slabInternal.length === 0))
    ) {
      open.slabs.push(slab)
      open.gaps = intersectIntervals(open.gaps, slab.gaps)
      open.cover = combinedCover
    } else {
      close(open)
      open = { slabs: [slab], gaps: slab.gaps, cover: slab.cover }
    }
  }
  if (open) close(open)
  return ranges
}

/**
 * Section direction = the DOMINANT script (strong-char majority). First-strong
 * (P2's paragraph rule) is position-biased on multi-column pages — the
 * top-left element wins even when the page is mostly RTL — so column reading
 * order votes instead.
 */
function sectionDir(elements: readonly SectionElement[]): Dir {
  let rtl = 0
  let ltr = 0
  for (const el of elements) {
    for (const c of el.unit?.chars ?? []) {
      if (isRtlScript(c.script)) rtl++
      else if (c.script !== 'common') ltr++
    }
  }
  return rtl > ltr ? 'rtl' : 'ltr'
}

/** assign elements to the coverage bands between consecutive gutters */
function splitByGutters(
  elements: readonly SectionElement[],
  box: Rect,
  gutters: readonly Interval[],
): LayoutColumn[] {
  const bounds = [box.x0, ...gutters.map((g) => (g.lo + g.hi) / 2), box.x1]
  const columns: LayoutColumn[] = []
  for (let i = 1; i < bounds.length; i++) {
    columns.push({
      box: { x0: bounds[i - 1]!, x1: bounds[i]!, y0: box.y0, y1: box.y1 },
      elements: [],
    })
  }
  for (const el of elements) {
    const cx = (el.box.x0 + el.box.x1) / 2
    let col = columns.findIndex((c) => cx >= c.box.x0 && cx <= c.box.x1)
    if (col < 0) col = cx < box.x0 ? 0 : columns.length - 1
    columns[col]!.elements.push(el)
  }
  return columns
}

/** column content width; empty columns count as zero-width (always weak) */
const contentWidthOf = (col: LayoutColumn): number =>
  col.elements.length === 0
    ? 0
    : Math.max(...col.elements.map((e) => e.box.x1)) -
      Math.min(...col.elements.map((e) => e.box.x0))

/** solve one section's columns; degrades to a single column when gates fail */
function buildSection(elements: SectionElement[], gutterMin: number): LayoutSection {
  const box = rectUnionAll(elements.map((e) => e.box))
  const dir = sectionDir(elements)
  const single = (): LayoutSection => ({
    box,
    columns: [{ box, elements: [...elements].sort(byReadingOrder) }],
    gutters: [],
    dir,
  })

  const cover = mergeIntervals(elements.map((e) => ({ lo: e.box.x0, hi: e.box.x1 })))
  // persistent gutters = gaps present in every slab of the section
  let gaps: Interval[] | null = null
  for (const slab of slabsOf(elements, box.x0, box.x1)) {
    gaps = gaps === null ? slab.gaps : intersectIntervals(gaps, slab.gaps)
  }
  let gutters = internalGaps(
    (gaps ?? []).filter((g) => g.hi - g.lo >= gutterMin),
    cover,
  )

  // a weak column (too few elements / sliver-narrow content) must not degrade
  // the whole section (P14 A): a card's "02" number seeds its own pseudo-gutter
  // and the one-element band it minted used to void the REAL column split.
  // Merge the weak column into a neighbor by dropping its narrower flanking
  // gutter and re-split; single-column only when no gutter survives.
  while (gutters.length > 0) {
    const columns = splitByGutters(elements, box, gutters)
    const isStrong = (c: LayoutColumn): boolean =>
      (c.elements.length >= MIN_COLUMN_ELEMENTS ||
        // a detected table is a heavily-gated aggregate of dozens of units —
        // the column it anchors is real even standing alone (P16 E)
        c.elements.some((e) => e.block?.kind === 'table')) &&
      contentWidthOf(c) >= MIN_COLUMN_WIDTH_PT
    // standalone weak columns (P15 B) only exist beside real columns — an
    // all-weak split (an author-name pair) still degrades to a single column
    const anchored = columns.some(isStrong)
    const weak = columns.findIndex((c, i) => {
      if (isStrong(c)) return false
      // deliberately-set-apart column (P15 B): every flanking gutter is vast
      const flanks = [gutters[i - 1], gutters[i]].filter((g) => g !== undefined)
      const standalone =
        anchored &&
        c.elements.length >= STANDALONE_MIN_ELEMENTS &&
        contentWidthOf(c) >= MIN_COLUMN_WIDTH_PT &&
        flanks.every((g) => g.hi - g.lo >= STANDALONE_GUTTER_FACTOR * gutterMin)
      return !standalone
    })
    if (weak < 0) {
      for (const col of columns) {
        col.elements.sort(byReadingOrder)
        // hug the column box to its content (the gutter split is generous)
        const extent = rectUnionAll(col.elements.map((e) => e.box))
        col.box = { ...col.box, x0: extent.x0, x1: extent.x1 }
      }
      return { box, columns, gutters, dir }
    }
    const left = weak > 0 ? gutters[weak - 1]! : undefined
    const right = weak < gutters.length ? gutters[weak] : undefined
    const dropAt =
      left === undefined
        ? weak
        : right === undefined || right.hi - right.lo >= left.hi - left.lo
          ? weak - 1
          : weak
    gutters = gutters.filter((_, i) => i !== dropAt)
  }
  return single()
}

/** top → bottom, then left → right (layout, not logical, order) */
function byReadingOrder(a: SectionElement, b: SectionElement): number {
  return b.box.y1 - a.box.y1 || a.box.x0 - b.box.x0
}

/**
 * Peel isolated header/footer band content into leading/trailing single-column
 * groups so a page number or running head never distorts the column solve.
 */
function peelHeaderFooter(
  elements: SectionElement[],
  pageHeightPt: number,
  bandRatio: number,
): { header: SectionElement[]; body: SectionElement[]; footer: SectionElement[] } {
  const lineH = median(elements.map((e) => e.box.y1 - e.box.y0)) || 12
  const minGap = HF_GAP_LINE_HEIGHTS * lineH
  const topBand = pageHeightPt * (1 - bandRatio)
  const bottomBand = pageHeightPt * bandRatio

  const header = elements.filter((e) => e.box.y0 >= topBand)
  const footer = elements.filter((e) => e.box.y1 <= bottomBand)
  const body = elements.filter((e) => !header.includes(e) && !footer.includes(e))

  const headerOk =
    header.length > 0 &&
    body.length > 0 &&
    Math.min(...header.map((e) => e.box.y0)) - Math.max(...body.map((e) => e.box.y1)) >= minGap
  const footerOk =
    footer.length > 0 &&
    body.length > 0 &&
    Math.min(...body.map((e) => e.box.y0)) - Math.max(...footer.map((e) => e.box.y1)) >= minGap

  return {
    header: headerOk ? header : [],
    body:
      headerOk && footerOk
        ? body
        : headerOk
          ? [...body, ...footer]
          : footerOk
            ? [...header, ...body]
            : elements,
    footer: footerOk ? footer : [],
  }
}

/** Full section/column pass over one page's flow elements. */
/**
 * Merge directly-adjacent multi-column sections with the SAME split (P22 D):
 * a hair of slab noise (a wrapped heading, a list marker crossing the
 * gutter) closes a 2-column run mid-page, and LibreOffice renders every
 * extra continuous-section transition with its own column balancing — a
 * torn body page grows stubs and ghost pages. Same column count + pairwise
 * overlapping gutters = one structure; text order improves too (the tear
 * interleaved half-columns).
 */
export function mergeTwinSections(sections: LayoutSection[]): LayoutSection[] {
  const out: LayoutSection[] = []
  for (const s of sections) {
    const prev = out[out.length - 1]
    const twins =
      prev !== undefined &&
      prev.columns.length >= 2 &&
      prev.columns.length === s.columns.length &&
      prev.gutters.length === s.gutters.length &&
      prev.gutters.every((g, i) => {
        const h = s.gutters[i]!
        return g.lo < h.hi && h.lo < g.hi
      })
    if (!twins) {
      out.push({ ...s, columns: s.columns.map((c) => ({ ...c, elements: [...c.elements] })) })
      continue
    }
    prev!.box = rectUnion(prev!.box, s.box)
    prev!.gutters = prev!.gutters.map((g, i) => ({
      lo: Math.max(g.lo, s.gutters[i]!.lo),
      hi: Math.min(g.hi, s.gutters[i]!.hi),
    }))
    for (const [i, col] of s.columns.entries()) {
      const pc = prev!.columns[i]!
      pc.box = rectUnion(pc.box, col.box)
      pc.elements.push(...col.elements)
    }
  }
  return out
}

export function detectSections(
  elements: readonly SectionElement[],
  pageHeightPt: number,
  pageWidthPt?: number,
): LayoutSection[] {
  if (elements.length === 0) return []
  const landscape = pageWidthPt !== undefined && pageWidthPt > pageHeightPt
  const { header, body, footer } = peelHeaderFooter(
    [...elements],
    pageHeightPt,
    landscape ? HF_BAND_RATIO_LANDSCAPE : HF_BAND_RATIO,
  )

  const sections: LayoutSection[] = []
  if (header.length > 0) sections.push(buildSectionSingle(header))

  if (body.length > 0) {
    const bodyBox = rectUnionAll(body.map((e) => e.box))
    const fontSize = median(body.filter((e) => e.unit).map((e) => e.unit!.fontSize)) || 12
    const gutterMin = Math.max(
      landscape ? GUTTER_MIN_PT_LANDSCAPE : GUTTER_MIN_PT,
      GUTTER_MIN_EMS * fontSize,
    )
    const lineH = median(body.map((e) => e.box.y1 - e.box.y0)) || 12
    const slabs = slabsOf(body, bodyBox.x0, bodyBox.x1)
    for (const range of sweepSections(slabs, gutterMin, FRESH_STRUCTURE_MIN_RUN_LINES * lineH)) {
      const inRange = body.filter((e) => {
        const cy = (e.box.y0 + e.box.y1) / 2
        return cy <= range.top + 0.01 && cy >= range.bottom - 0.01
      })
      if (inRange.length > 0) sections.push(buildSection(inRange, gutterMin))
    }
    // elements whose center escaped every range (edge overlap) — append as single
    const placed = new Set(sections.flatMap((s) => s.columns.flatMap((c) => c.elements)))
    const stray = body.filter((e) => !placed.has(e))
    if (stray.length > 0) sections.push(buildSectionSingle(stray))
  }

  if (footer.length > 0) sections.push(buildSectionSingle(footer))
  sections.sort((a, b) => b.box.y1 - a.box.y1)
  return sections
}

function buildSectionSingle(elements: SectionElement[]): LayoutSection {
  const box = rectUnionAll(elements.map((e) => e.box))
  return {
    box,
    columns: [{ box, elements: [...elements].sort(byReadingOrder) }],
    gutters: [],
    dir: sectionDir(elements),
  }
}
