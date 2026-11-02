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
  pageNo: number
  display: string
  /** zero-height slice: a deliberate blank page (double w:br, even/odd parity) */
  blank: boolean
  /** page-pinned floated blocks whose anchor lands on this page */
  floats: number[]
}

export interface PageCutOptions {
  evenOddHf?: boolean
  titlePgOf?: (section: number) => boolean | undefined
  pageNumberFmtOf?: (section: number) => string | undefined
  floatSplits?: FloatSplitPatch[]
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
  blocks.forEach((b, i) => {
    if (b.el && (b.floated || b.floatTable) && slices.length > 0)
      floats[pinnedFloatPage(slices, b.top)].push(i)
  })
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
    const leadDrops =
      !!lb &&
      (!!lb.breakBefore ||
        !!blocks[lead.block - 1]?.breakAfter ||
        (lb.dropAtPageTop ?? leadDropsSpaceBefore(blocks, lead.block)))
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
      pageNo,
      display: formatPageNumber(pageNo, opts.pageNumberFmtOf?.(s.section)),
      blank: s.end === s.start,
      floats: floats[i],
    }
  })
  applyPageWindows(cuts, blocks, geoms)
  return cuts
}
