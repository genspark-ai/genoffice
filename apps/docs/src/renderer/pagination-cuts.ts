/**
 * PageCut: the page-level facts one layout pass derives from the slicing,
 * built once after sliceWithLineSplit from the measured blocks and the
 * section geometry. Pure (no DOM), so jsdom tests can run it. Consumers
 * (gap builder, preview, header/footer strips, TOC, status bar) switch to it
 * stage by stage; until then the canvas shadows it (`pageCutMismatch`).
 */
import { BlockIndex } from './pagination-index'
import { formatPageNumber, hfVariantOf, hfVariantReservedPx, type HfVariant } from './pagination-hf'
import {
  applyLeadSpills,
  applyLiftTops,
  markTableSeamSlices,
  pageCapacityPx,
  pageNoAfterOf,
  pageOfY,
  pinnedFloatPage,
  sectionFirstPages,
} from './pagination-slices'
import type {
  BlockBox,
  FloatBox,
  FloatSplitPatch,
  PageRegion,
  PageSlice,
  SectionGeom,
  SectionHfHeights,
} from './pagination-types'

/** matches the gap builder's DOM row match (`Math.abs(off - cutOff) < 1.5`) */
export const LEAD_ROW_TOL_PX = 1.5

/** Plan-less fallback of BlockPlan.dropAtPageTop: Word drops the space-before
 *  of a block that opens a page after a natural break too (probe 2026-10-20: a
 *  before=360 heading pushed by overflow or by its keepNext chain lands at the
 *  plain first-line position, compat 14 and 15); only a section's first page
 *  keeps it */
export function leadDropsSpaceBefore(blocks: { section?: number }[], i: number): boolean {
  return i > 0 && (blocks[i].section ?? 0) === (blocks[i - 1].section ?? 0)
}

export type PageLeadKind = 'block' | 'row' | 'cell' | 'line' | 'tail'

/** Where the page's first content starts: a whole block, a table row, a line
 *  inside a cell, a line inside a paragraph, or nothing (trailing blank page) */
export interface PageLead {
  /** block index; -1 for `tail` */
  block: number
  kind: PageLeadKind
  /** `row` / `cell`: index into the table's measured rows (`cell` = the row the cut falls inside) */
  row?: number
  /** `line` / `cell`: cut offset from the block's DOM top (px), as the line anchors measure it */
  lineOffset?: number
}

export interface PageCut {
  start: number
  end: number
  section: number
  regions?: PageRegion[]
  repeatHeader?: PageSlice['repeatHeader']
  continuedSection?: true
  physHeight?: number
  liftTop?: number
  leadSpill?: number
  leadTable?: true
  cutTable?: true
  tailTable?: true
  lead: PageLead
  /** slicer capacity of this page (pageCapacityPx) */
  capacityPx: number
  /** flow span used, including a repeated table header */
  usedPx: number
  /** physical height the page's content occupies (regioned pages stack their
   *  column regions; otherwise usedPx) */
  physUsedPx: number
  /** unused body height below the content (vertical text: the body height,
   *  not the sideways fill extent the capacity measures) */
  padPx: number
  hf: { variant: HfVariant; headerPx: number; footerPx: number }
  /** a block lead loses its space before: an explicit break opened the page
   *  (Word probe 2026-09-23) or the plan drops it at a natural lead; only a
   *  section's first page keeps it */
  leadDropsSpaceBefore: boolean
  /** the lead's explicit break keeps this much of its space before (px): its
   *  margin-top below the gap, instead of the declared (collapsed) one */
  leadMtPx?: number
  pageNo: number
  display: string
  /** zero-height slice: a deliberate blank page (double w:br, even/odd parity) */
  blank: boolean
  /** page-pinned floated blocks whose anchor lands on this page */
  floats: number[]
  /** anchor-shifted blocks (data-anchor-dy) whose flow slot lands on this page */
  anchored: number[]
  /** floating boxes (`PageCutOptions.floats` indices) whose anchor lands on this page */
  boxes: number[]
  /** the gap widget that opens this page (applyGapLedger); absent on the first
   *  page, a parity twin and a boundary the gap builder dropped */
  gap?: PageGapFacts
  /** cumulative height of the page-turn gaps above this page's content (px at
   *  zoom 1): screen Y of a flow y on this page = origin + (y + gapOffsetPx) * zoom */
  gapOffsetPx: number
  /** cumulative flow displacement above this page's content: the gaps outside
   *  floating tables plus their repeated-header clones (a split floating table's
   *  in-table gaps move only the float; its carry spacer moves the flow, see
   *  flowOffsetAt) */
  flowOffsetPx: number
}

/** page-turn widget kinds: a zero-height `cut` marker (multi-cell row cut,
 *  nested-table overlay) is a boundary without a band */
export type PageGapKind = 'block' | 'inline' | 'table' | 'cell' | 'cut'

/** What the gap builder emitted for a page turn, in px at zoom 1 */
export interface PageGapFacts {
  kind: PageGapKind
  /** rendered box height: bottom margin + band + top margin (0 for a cut marker) */
  heightPx: number
  /** previous page's bottom margin incl. the footnote area and the pad */
  marginBottomPx: number
  /** next page's top margin after the header push (`--gap-mt`) */
  marginTopPx: number
  /** next page's section left margin on the paper (`--gap-ml`) */
  marginLeftPx: number
  /** header push past the section's top margin (page-edge V offsets count from the pgMar top) */
  topPushPx?: number
  /** mixed-column page above: the gap is pulled up over the vacated stacked space */
  pullUpPx: number
  /** repeated table header clones rendered right below a table gap */
  repeatHeaderPx: number
  /** in-table gap of a split floating table: displaces the float, not the flow */
  floatGap?: true
  /** flow y where the widget landed (learnGapFlowTops): the slice boundary plus
   *  the trailing space the previous block keeps in the DOM (Word drops it at
   *  the page bottom), minus the pull-up */
  flowTopPx?: number
  /** next page's left edge and width on the shared paper (differing-width documents) */
  pageBox?: { left: number; width: number }
}

/** flow spacer that carries a split floating table's anchor paragraph past the
 *  table's in-table gaps: sits at flow y `y`, `px` tall */
export interface FlowCarry {
  y: number
  px: number
}

/** one page-turn entry of the gap builder's output */
export interface GapLedgerEntry extends PageGapFacts {
  /** index of the page the gap opens */
  page: number
}

/**
 * Stamp the gap facts the builder emitted onto the cuts and accumulate the two
 * offsets. Pages without an entry (first page, parity twins, dropped
 * boundaries) inherit the offsets of the page above.
 */
export function applyGapLedger(cuts: PageCut[], ledger: readonly GapLedgerEntry[]): void {
  const byPage = new Map<number, PageGapFacts>()
  for (const { page, ...facts } of ledger) byPage.set(page, facts)
  let gapOffsetPx = 0
  let flowOffsetPx = 0
  cuts.forEach((cut, i) => {
    const gap = byPage.get(i)
    if (gap) {
      cut.gap = gap
      gapOffsetPx += gap.heightPx - gap.pullUpPx
      if (!gap.floatGap) flowOffsetPx += gap.heightPx - gap.pullUpPx + gap.repeatHeaderPx
    } else delete cut.gap
    cut.gapOffsetPx = gapOffsetPx
    cut.flowOffsetPx = flowOffsetPx
  })
}

/** flow displacement (px at zoom 1) the DOM shows above flow y: the page's
 *  flowOffsetPx plus the carry spacers at or above y */
export function flowOffsetAt(
  cuts: readonly PageCut[],
  carries: readonly FlowCarry[],
  y: number,
  tol = 0,
): number {
  let off = 0
  for (const cut of cuts) {
    const gap = cut.gap
    if (!gap || gap.floatGap || (gap.flowTopPx ?? cut.start) > y + tol) continue
    off += gap.heightPx - gap.pullUpPx + gap.repeatHeaderPx
  }
  return carries.reduce((s, c) => (c.y <= y + tol ? s + c.px : s), off)
}

/** Stamp the page-turn widgets' landed flow tops (DOM order = page order; px at
 *  zoom 1 from the first page's content top) onto the ledger; a count mismatch
 *  leaves the slice boundaries in place. */
export function learnGapFlowTops(cuts: readonly PageCut[], widgetTops: readonly number[]): boolean {
  const turns = cuts.filter((c, i) => i > 0 && c.gap && c.gap.kind !== 'cut')
  if (turns.length !== widgetTops.length) return false
  turns.forEach((cut, i) => {
    const prev = cuts[cuts.indexOf(cut) - 1]
    cut.gap!.flowTopPx = widgetTops[i] - prev.gapOffsetPx
  })
  return true
}

/** the nearest page at or above `page` whose gap knows its header push (cut
 *  markers carry none, like the DOM markers before them) */
export function topPushAt(cuts: readonly PageCut[], page: number): number | undefined {
  for (let p = page; p >= 0; p--) {
    const push = cuts[p].gap?.topPushPx
    if (push !== undefined) return push
  }
  return undefined
}

/** Paper edges of every page a gap opens, in flow px from the editor box top
 *  (zoom 1): where its paper starts, where the previous paper ends, and the gap
 *  box itself */
export interface PagePaperEdge {
  page: number
  gap: PageGapFacts
  /** this page's paper top (gap bottom minus the top margin) */
  top: number
  /** previous page's paper bottom (gap top plus the bottom margin) */
  prevBottom: number
  gapTop: number
  gapBottom: number
}

export function pagePaperEdges(cuts: readonly PageCut[], firstMarginTop: number): PagePaperEdge[] {
  const out: PagePaperEdge[] = []
  cuts.forEach((cut, page) => {
    const gap = cut.gap
    if (!gap || page === 0) return
    const gapTop =
      gap.flowTopPx === undefined
        ? firstMarginTop + cut.start + cut.gapOffsetPx - gap.heightPx
        : firstMarginTop + gap.flowTopPx + cuts[page - 1].gapOffsetPx
    const gapBottom = gapTop + gap.heightPx
    out.push({
      page,
      gap,
      top: gapBottom - gap.marginTopPx,
      prevBottom: gapTop + gap.marginBottomPx,
      gapTop,
      gapBottom,
    })
  })
  return out
}

/** The blocks that open a page behind a block-kind gap widget, with the page
 *  index: the wrap-obstacle pass clears its rectangles in front of them (a
 *  page turn inside a table or a paragraph has no top-level widget and resets
 *  nothing, as before). */
export function pageLeadsOf(
  cuts: readonly PageCut[],
  blocks: readonly BlockBox[],
): Map<Element, number> {
  const out = new Map<Element, number>()
  cuts.forEach((c, i) => {
    const gap = c.gap
    if (!gap) return
    let el: HTMLElement | undefined
    if (c.lead.kind === 'block') {
      if (gap.kind === 'block') el = blocks[c.lead.block]?.el
    } else if (c.lead.kind !== 'tail' && !gap.floatGap) {
      // the turn sits inside a table or paragraph: Word scopes wrap obstacles
      // to the page, so the first block after the straddling one opens the
      // new set (a block straddling several pages resets once, at its last turn)
      for (let b = c.lead.block + 1; b < blocks.length && !el; b++) el = blocks[b].el
    }
    if (el) out.set(el, i)
  })
  return out
}

/** page a floated block was pinned to (`floats`), or undefined */
export function floatPageOf(cuts: readonly PageCut[], block: number): number | undefined {
  const i = cuts.findIndex((c) => c.floats.includes(block))
  return i >= 0 ? i : undefined
}

/** owning page by block index (`floats` and `anchored`) and by floating-box index (`boxes`) */
export function pageOwners(cuts: readonly PageCut[]): {
  blocks: ReadonlyMap<number, number>
  boxes: ReadonlyMap<number, number>
} {
  const blocks = new Map<number, number>()
  const boxes = new Map<number, number>()
  cuts.forEach((c, i) => {
    for (const b of c.floats) blocks.set(b, i)
    for (const b of c.anchored) blocks.set(b, i)
    for (const f of c.boxes) boxes.set(f, i)
  })
  return { blocks, boxes }
}

/**
 * Transition shadow (T4.5): every page reader compares the page it derived
 * from the DOM gap widgets with the PageCut-derived page; disagreements are
 * counted here and published as `__pageDebug.pageReaderMismatch`. Cumulative
 * like the other counters; `enabled` is set per pass (dev or `__pageReaderShadow`).
 */
export const pageReaderShadow = {
  enabled: false,
  count: 0,
  diag: [] as string[],
  report(reader: string, dom: unknown, cut: unknown, detail = ''): void {
    this.count++
    if (this.diag.length < 40)
      this.diag.push(
        `${reader}: dom ${String(dom)} vs cut ${String(cut)}${detail ? ` ${detail}` : ''}`,
      )
  },
}

export interface PageCutOptions {
  evenOddHf?: boolean
  titlePgOf?: (section: number) => boolean | undefined
  pageNumberFmtOf?: (section: number) => string | undefined
  floatSplits?: FloatSplitPatch[]
  /** the measured floating boxes; each cut lists the ones its page owns (`boxes`) */
  floats?: readonly FloatBox[]
}

/** Page-window facts the sheets apply around a cut: table seams, the part of
 *  a lifted block hanging above the page start, the last line's spilled
 *  leading. Computed once here for the canvas and the preview. */
export function applyPageWindows(cuts: PageCut[], blocks: BlockBox[], geoms: SectionGeom[]): void {
  markTableSeamSlices(cuts, blocks)
  applyLiftTops(cuts, blocks)
  applyLeadSpills(cuts, blocks, geoms)
}

/**
 * Flow window owning an anchor on its page and the hoisted box's content-area
 * Y: the page window on plain pages, the anchor's column on regioned pages
 * (column-relative X boxes in a multi-column region are left to the column).
 */
export function hoistWindow(
  slice: Pick<PageSlice, 'start' | 'end' | 'regions'>,
  anchorTop: number,
): { start: number; end: number; dy: number; multiCol: boolean } | null {
  if (!slice.regions) {
    return { start: slice.start, end: slice.end, dy: anchorTop - slice.start, multiCol: false }
  }
  for (const region of slice.regions) {
    for (const col of region.columns) {
      if (anchorTop >= col.start && anchorTop < col.end) {
        return {
          start: col.start,
          end: col.end,
          dy: region.top + anchorTop - col.start,
          multiCol: region.columns.length > 1,
        }
      }
    }
  }
  return null
}

/** page number of each heading element: its block by element identity, else
 *  by data-idx (a pagination pass can restamp the node, like the caret-page
 *  lookup); a heading with no block counts as a miss and takes the first page */
export function headingPagesOf(
  blocks: readonly BlockBox[],
  cuts: readonly PageCut[],
  headings: ReadonlyArray<HTMLElement | null>,
): { pages: number[]; misses: number } {
  const byEl = new Map<HTMLElement, BlockBox>()
  const byIdx = new Map<string, BlockBox>()
  for (const b of blocks) {
    if (!b.el) continue
    byEl.set(b.el, b)
    if (b.docxIndex !== undefined && !byIdx.has(String(b.docxIndex)))
      byIdx.set(String(b.docxIndex), b)
  }
  let misses = 0
  const pages = headings.map((el) => {
    const idx = el?.dataset.idx
    const block = el
      ? (byEl.get(el) ?? (idx !== undefined ? byIdx.get(idx) : undefined))
      : undefined
    if (!block) {
      misses++
      return cuts[0].pageNo
    }
    const i = pageOfY(cuts as PageCut[], block.top)
    return cuts[Math.min(Math.max(i, 0), cuts.length - 1)].pageNo
  })
  return { pages, misses }
}

const px1 = (v: number | undefined) => (v === undefined ? '' : v.toFixed(1))

/** one line per page: what a page sheet depends on; equal strings = same pages */
export function pageCutSignature(cuts: readonly PageCut[]): string {
  return cuts
    .map((c) =>
      [
        c.start.toFixed(1),
        c.end.toFixed(1),
        c.section,
        `${c.lead.kind}${c.lead.block}${c.lead.row ?? ''}`,
        c.hf.variant,
        c.pageNo,
        c.regions?.length ?? 0,
        px1(c.repeatHeader?.height),
        px1(c.liftTop),
        px1(c.leadSpill),
        `${c.leadTable ? 'L' : ''}${c.cutTable ? 'C' : ''}${c.tailTable ? 'T' : ''}`,
      ].join(':'),
    )
    .join('|')
}

/** where two cut lists of one document disagree (0.5 px on lengths), one line per fact */
const sameList = (a: readonly number[], b: readonly number[]) =>
  a.length === b.length && a.every((v, i) => v === b[i])

export function pageCutDiffs(a: readonly PageCut[], b: readonly PageCut[]): string[] {
  const out: string[] = []
  if (a.length !== b.length) out.push(`pages ${a.length} vs ${b.length}`)
  const n = Math.min(a.length, b.length)
  const near = (x: number | undefined, y: number | undefined) => Math.abs((x ?? 0) - (y ?? 0)) < 0.5
  for (let i = 0; i < n; i++) {
    const x = a[i]
    const y = b[i]
    const facts: Array<[string, boolean, unknown, unknown]> = [
      ['start', near(x.start, y.start), x.start, y.start],
      ['end', near(x.end, y.end), x.end, y.end],
      ['section', x.section === y.section, x.section, y.section],
      [
        'lead',
        x.lead.kind === y.lead.kind && x.lead.block === y.lead.block && x.lead.row === y.lead.row,
        `${x.lead.kind}${x.lead.block}${x.lead.row ?? ''}`,
        `${y.lead.kind}${y.lead.block}${y.lead.row ?? ''}`,
      ],
      ['capacity', near(x.capacityPx, y.capacityPx), x.capacityPx, y.capacityPx],
      ['hf', x.hf.variant === y.hf.variant, x.hf.variant, y.hf.variant],
      ['pageNo', x.pageNo === y.pageNo, x.pageNo, y.pageNo],
      [
        'regions',
        (x.regions?.length ?? 0) === (y.regions?.length ?? 0),
        x.regions?.length,
        y.regions?.length,
      ],
      [
        'repeatHeader',
        near(x.repeatHeader?.height, y.repeatHeader?.height),
        x.repeatHeader?.height,
        y.repeatHeader?.height,
      ],
      ['liftTop', near(x.liftTop, y.liftTop), x.liftTop, y.liftTop],
      ['leadSpill', near(x.leadSpill, y.leadSpill), x.leadSpill, y.leadSpill],
      ['leadTable', !!x.leadTable === !!y.leadTable, x.leadTable, y.leadTable],
      ['cutTable', !!x.cutTable === !!y.cutTable, x.cutTable, y.cutTable],
      ['tailTable', !!x.tailTable === !!y.tailTable, x.tailTable, y.tailTable],
      ['display', x.display === y.display, x.display, y.display],
      ['floats', sameList(x.floats, y.floats), x.floats.join(','), y.floats.join(',')],
      ['anchored', sameList(x.anchored, y.anchored), x.anchored.join(','), y.anchored.join(',')],
      ['boxes', sameList(x.boxes, y.boxes), x.boxes.join(','), y.boxes.join(',')],
    ]
    for (const [name, same, va, vb] of facts)
      if (!same) out.push(`p${i + 1} ${name} ${String(va)} vs ${String(vb)}`)
  }
  return out
}

function leadOf(
  blocks: BlockBox[],
  index: BlockIndex,
  y: number,
  floatSplits: FloatSplitPatch[] | undefined,
): PageLead {
  const fsplit = floatSplits?.find((f) => f.cutYs.some((c) => Math.abs(c - y) < 0.5))
  const fsIdx = fsplit ? index.firstAtTop(fsplit.blockTop, (bb) => !!bb.el && !!bb.floatTable) : -1
  const fsBlock = fsIdx >= 0 ? blocks[fsIdx] : undefined
  const i = fsBlock ? -1 : index.firstAtTop(y)
  if (i >= 0 && blocks[i].el) return { block: i, kind: 'block' }
  const b = fsBlock ?? index.containing(y, (bb) => !!bb.el)
  if (!b?.el) return { block: -1, kind: 'tail' }
  const bi = blocks.indexOf(b)
  const cutOff = y - b.top - (b.spaceBeforePx ?? 0) - (fsBlock ? fsplit!.dyPx : 0)
  const rows = b.tableRows
  if (rows && rows.length > 0) {
    let top = 0
    let inside = -1
    for (let r = 0; r < rows.length; r++) {
      if (top <= cutOff + 0.5) inside = r
      if (Math.abs(top - cutOff) < LEAD_ROW_TOL_PX) return { block: bi, kind: 'row', row: r }
      top += rows[r].height
    }
    return { block: bi, kind: 'cell', row: inside, lineOffset: cutOff }
  }
  return { block: bi, kind: 'line', lineOffset: cutOff }
}

export function buildPageCuts(
  blocks: BlockBox[],
  slices: PageSlice[],
  geoms: SectionGeom[],
  hfHeights: SectionHfHeights[],
  opts: PageCutOptions = {},
): PageCut[] {
  const index = new BlockIndex(blocks)
  const firsts = sectionFirstPages(slices)
  const geomOf = (s: number) => geoms[Math.max(0, Math.min(s, geoms.length - 1))]
  const floats: number[][] = slices.map(() => [])
  const anchored: number[][] = slices.map(() => [])
  blocks.forEach((b, i) => {
    if (slices.length === 0 || !b.el) return
    if (b.floated || b.floatTable) floats[pinnedFloatPage(slices, b.top)].push(i)
    if (b.anchorShifted) anchored[pinnedFloatPage(slices, b.top)].push(i)
  })
  const boxes: number[][] = slices.map(() => [])
  if (slices.length > 0)
    opts.floats?.forEach((f, i) => boxes[pinnedFloatPage(slices, f.anchorTop)].push(i))
  let prevNo = 0
  let prevSection = -1
  const cuts: PageCut[] = slices.map((s, i) => {
    const g = geomOf(s.section)
    const pageNo = pageNoAfterOf(geomOf, prevNo, prevSection, s.section)
    prevNo = pageNo
    prevSection = s.section
    const capacityPx = pageCapacityPx(g, firsts[i], pageNo)
    const usedPx = s.end - s.start + (s.repeatHeader?.height ?? 0)
    const physUsedPx = s.regions ? (s.physHeight ?? usedPx) : usedPx
    const bodyPx = g.vertical
      ? firsts[i] && g.firstBodyHeight !== undefined
        ? g.firstBodyHeight
        : g.evenBodyHeight !== undefined && pageNo % 2 === 0
          ? g.evenBodyHeight
          : (g.contentWidth ?? capacityPx)
      : capacityPx
    const lead = leadOf(blocks, index, s.start, opts.floatSplits)
    const lb = lead.kind === 'block' ? blocks[lead.block] : undefined
    const leadKeepsPx = lb?.leadMtPx
    const leadDrops =
      !!lb &&
      (!!blocks[lead.block - 1]?.breakAfter ||
        (leadKeepsPx === undefined &&
          (!!lb.breakBefore || (lb.dropAtPageTop ?? leadDropsSpaceBefore(blocks, lead.block)))))
    const variant = hfVariantOf(opts.titlePgOf?.(s.section), firsts[i], !!opts.evenOddHf, pageNo)
    const h = hfHeights[Math.max(0, Math.min(s.section, hfHeights.length - 1))] ?? {
      headerPx: 0,
      footerPx: 0,
    }
    const hf = hfVariantReservedPx(h, variant)
    return {
      start: s.start,
      end: s.end,
      section: s.section,
      ...(s.regions ? { regions: s.regions } : {}),
      ...(s.repeatHeader ? { repeatHeader: s.repeatHeader } : {}),
      ...(s.continuedSection ? { continuedSection: true } : {}),
      ...(s.physHeight !== undefined ? { physHeight: s.physHeight } : {}),
      lead,
      capacityPx,
      usedPx,
      physUsedPx,
      padPx: Math.max(0, bodyPx - physUsedPx),
      hf: { variant, ...hf },
      leadDropsSpaceBefore: leadDrops,
      ...(!leadDrops && leadKeepsPx !== undefined ? { leadMtPx: leadKeepsPx } : {}),
      pageNo,
      display: formatPageNumber(pageNo, opts.pageNumberFmtOf?.(s.section)),
      blank: s.end === s.start,
      floats: floats[i],
      anchored: anchored[i],
      boxes: boxes[i],
      gapOffsetPx: 0,
      flowOffsetPx: 0,
    }
  })
  applyPageWindows(cuts, blocks, geoms)
  return cuts
}
