/**
 * The slicing pipeline between `measureBlocks` and the PageCut list, shared by
 * the canvas pass (App.tsx) and the pagination preview (its no-pass fallback
 * and its shadow re-slice). Pure over its inputs: the canvas and the preview
 * differ only in the context they hand in (header/footer resolver, resume),
 * never in the steps, so a step added for one side runs for the other.
 */
import type { SectionInfo, SectionSettings } from '@genoffice/docx-engine'
import { buildPageCuts, type PageCut } from './pagination-cuts'
import { sliceWithLineSplit, type PassResume } from './pagination-lines'
import { appendEndnotesBlock, appendFloatSpillBlock } from './pagination-measure'
import { assignSections, liveSections, sectionGeoms } from './pagination-sections'
import type {
  BlockBox,
  BlockMetaOf,
  FloatBox,
  FloatVShiftPatch,
  PageNoteItem,
  PageSlice,
  SectionGeom,
  SectionHfHeights,
  SliceOutputs,
} from './pagination-types'

const twipsToPx = (twips: number) => (twips / 1440) * 96

/** `measureBlocks` output */
export interface PassInputs {
  blocks: BlockBox[]
  totalHeight: number
  floats: FloatBox[]
  sectBreaks: Set<number>
}

export interface PassContext {
  /** parsed sections (empty = one section described by `section`) */
  sections: SectionInfo[]
  section: SectionSettings
  delSectBreaks?: Set<number>
  endnoteItems: PageNoteItem[]
  endnoteSeparatorPx: number
  /** per-section header/footer variant heights (multi-section documents) */
  hfHeightsOf: (secList: SectionInfo[]) => SectionHfHeights[]
  /** the lone section's heights and geometry (no `sections`) */
  singleHfPx: SectionHfHeights
  singleGeom: SectionGeom
  evenOddHf: boolean
  titlePg: boolean
  /** measure as a full-width single flow: the geometry drops column flow to match */
  dropCols: boolean
  blockMetaOf?: BlockMetaOf
  /** vertical-text sections: a block's inline extras (indents the writing-mode
   *  decoration moves off the physical margins), captured while measuring */
  inlineExtraOf?: (el: HTMLElement) => number
  factor: number
  /** canvas only: resume the previous slicing above the dirty block */
  resumeFor?: (blocks: BlockBox[], secSig: string) => PassResume | undefined
  out: SliceOutputs
}

/** the flow after sections, endnotes and float spill are in place; input of `sliceFlow` */
export interface PreparedFlow {
  blocks: BlockBox[]
  floats: FloatBox[]
  secList: SectionInfo[] | null
  secSig: string
  hfHeights: SectionHfHeights[]
  geoms: SectionGeom[]
  flowH: number
  /** flow height before the float spill block */
  flowBase: number
  spillOverhangPx: number
  hasSpill: boolean
  endnotesTop: number | null
}

export interface SlicedFlow {
  slices: PageSlice[]
  cuts: PageCut[]
  resumed: boolean
}

export function prepareFlow(inputs: PassInputs, ctx: PassContext): PreparedFlow {
  const { blocks, totalHeight, floats, sectBreaks } = inputs
  // liveSections: when a section-break block is deleted, that section merges
  // into the next in real time (effective before saving)
  const secList =
    ctx.sections.length > 0
      ? liveSections(ctx.sections, blocks, sectBreaks, ctx.delSectBreaks)
      : null
  if (secList) assignSections(blocks, secList)
  if (ctx.inlineExtraOf)
    for (const b of blocks) if (b.el && !b.floated) b.inlineExtraPx = ctx.inlineExtraOf(b.el)
  // the endnote area takes part in page slicing (placed together at the
  // document end; overflows continue on later pages)
  const withEndnotes = appendEndnotesBlock(
    blocks,
    totalHeight,
    ctx.endnoteItems,
    ctx.endnoteSeparatorPx,
  )
  const hfHeights = secList ? ctx.hfHeightsOf(secList) : [ctx.singleHfPx]
  const geoms: SectionGeom[] = secList
    ? sectionGeoms(secList, hfHeights, ctx.evenOddHf)
    : [ctx.singleGeom]
  if (ctx.dropCols) for (const g of geoms) if (g.cols) g.cols = undefined
  // floating boxes below the flow end still need pages to land on; boxes
  // reaching only into the last page's bottom margin stay there (Word draws
  // anchored objects over the margin instead of opening a page)
  const lastSec = secList?.[secList.length - 1]?.settings ?? ctx.section
  const spillOverhangPx =
    geoms[geoms.length - 1]?.marginBottomPx ?? (lastSec ? twipsToPx(lastSec.marginBottom) : 0)
  const flowBase = withEndnotes?.totalHeight ?? totalHeight
  const flowWithFloats = appendFloatSpillBlock(blocks, flowBase, floats, spillOverhangPx)
  const secSig = secList
    ? secList.map((s) => `${s.firstBlockIndex}:${s.lastBlockIndex}:${s.startType ?? ''}`).join('|')
    : ''
  return {
    blocks,
    floats,
    secList,
    secSig,
    hfHeights,
    geoms,
    flowH: flowWithFloats ?? flowBase,
    flowBase,
    spillOverhangPx,
    hasSpill: flowWithFloats !== null,
    endnotesTop: withEndnotes?.top ?? null,
  }
}

export function sliceFlow(flow: PreparedFlow, ctx: PassContext): SlicedFlow {
  const { blocks, floats, geoms, secList } = flow
  const out = ctx.out
  const resume = ctx.resumeFor?.(blocks, flow.secSig)
  let slices = sliceWithLineSplit(
    blocks,
    geoms,
    flow.flowH,
    ctx.factor,
    ctx.blockMetaOf,
    out,
    resume,
  )
  const floatTops = floats.map((f) => f.top)
  const pullFloats = () => {
    floats.forEach((f, i) => (f.top = floatTops[i]))
    applyFloatPullUps(floats, out.floatPullUps ?? [])
  }
  pullFloats()
  // the spill block was sized from unpulled boxes: a box pulled up to the
  // paper edge may no longer need the page it opened
  if (flow.hasSpill && out.floatPullUps?.length) {
    const spillAt = blocks.findIndex((b) => b.isFloatSpill)
    if (spillAt >= 0) blocks.splice(spillAt, 1)
    const again = appendFloatSpillBlock(blocks, flow.flowBase, floats, flow.spillOverhangPx)
    if ((again ?? flow.flowBase) < flow.flowH - 0.5) {
      // the published flow must carry the height the cuts were sliced from
      flow.flowH = again ?? flow.flowBase
      flow.hasSpill = again !== null
      slices = sliceWithLineSplit(blocks, geoms, flow.flowH, ctx.factor, ctx.blockMetaOf, out)
      pullFloats()
    }
  }
  const cuts = buildPageCuts(blocks, slices, geoms, flow.hfHeights, {
    evenOddHf: ctx.evenOddHf,
    // a lone section follows the live Different First Page toggle
    titlePgOf:
      secList && secList.length > 1
        ? (si) => secList[Math.min(si, secList.length - 1)].titlePg
        : () => ctx.titlePg,
    pageNumberFmtOf: secList
      ? (si) => secList[Math.min(si, secList.length - 1)]?.pageNumberFmt
      : undefined,
    floatSplits: out.floatSplits,
    floats,
  })
  return { slices, cuts, resumed: !!resume }
}

/** Move the boxes of a pulled-up anchor paragraph (floatPullUps) with it: the
 *  preview applies the canvas pass's list to its own float snapshot, as
 *  sliceFlow does to the pass's floats */
export function applyFloatPullUps(
  floats: readonly FloatBox[],
  pullUps: readonly FloatVShiftPatch[],
): void {
  for (const p of pullUps)
    for (const f of floats)
      if (!f.pinned && !f.pageRelV && Math.abs(f.anchorTop - p.blockTop) < 0.5) f.top += p.dyPx
}

/** `prepareFlow` + `sliceFlow` */
export function slicePass(inputs: PassInputs, ctx: PassContext): PreparedFlow & SlicedFlow {
  const flow = prepareFlow(inputs, ctx)
  const sliced = sliceFlow(flow, ctx)
  return { ...flow, ...sliced }
}
