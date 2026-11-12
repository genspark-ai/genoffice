/**
 * The preview prints the canvas pass's PageCut list against its own snapshot
 * of the blocks. Both sides now run one pipeline (pagination-pass); this suite
 * is the unit test the screenshot gate never had: over every corpus document,
 * the preview-mode run (block meta applied up front, shadow on a twin, no
 * resume) yields the canvas cuts, and an independent applyBlockMeta over a
 * second measure lands on the pass's block facts (blockFacts). The synthetic
 * case replays the footnote-band regression: a preview that skips
 * applyBlockMeta diverges in the block facts instead of a DEV-only shadow
 * hiding it. With a pass the preview measures nothing: the sheets position
 * against the pass's own blocks and floating boxes; the shadow twin starts
 * from the boxes as measured (pull-ups undone) because sliceFlow pulls them.
 */
import { readdirSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'
import {
  parseDocx,
  readSections,
  type ParsedDoc,
  type SectionSettings,
} from '@genoffice/docx-engine'
import {
  applyBlockMeta,
  pageCutDiffs,
  pageCutSignature,
  pageOfY,
  type BlockBox,
  type SectionGeom,
  type SliceOutputs,
} from '../src/renderer/pagination'
import {
  applyFloatPullUps,
  prepareFlow,
  sliceFlow,
  slicePass,
  type PassContext,
  type PassInputs,
} from '../src/renderer/pagination-pass'
import { buildCorpusBlocks } from './helpers/corpus-blocks'

const CORPUS_DIR = join(dirname(fileURLToPath(import.meta.url)), 'pagination-corpus/docx')
const stems = readdirSync(CORPUS_DIR)
  .filter((f) => f.endsWith('.docx'))
  .map((f) => f.slice(0, -'.docx'.length))
  .sort()

const freshOut = (): SliceOutputs => ({
  rowFills: [],
  rowSplits: [],
  floatVShifts: [],
  floatPullUps: [],
  floatCarries: [],
  floatFlows: [],
  floatSplits: [],
  oversizeClips: [],
})

const NO_HF = { headerPx: 0, footerPx: 0 }

function corpusCtx(parsed: ParsedDoc, geoms: SectionGeom[]): PassContext {
  const sections = readSections(parsed)
  return {
    sections,
    section: sections[0].settings,
    endnoteItems: [],
    endnoteSeparatorPx: 0,
    hfHeightsOf: (secs) => secs.map(() => NO_HF),
    singleHfPx: NO_HF,
    singleGeom: geoms[0],
    evenOddHf: false,
    titlePg: false,
    dropCols: false,
    factor: 1,
    out: freshOut(),
  }
}

/** the block facts the cuts were sliced from (the retired previewBlockDiffs, kept for the corpus proof) */
const blockFacts = (blocks: readonly BlockBox[]) =>
  blocks
    .filter((b) => !b.isFloatSpill)
    .map((b) => [
      b.el?.dataset.idx,
      b.docxIndex,
      Math.round(b.top * 2),
      Math.round(b.height * 2),
      Math.round((b.spaceBeforePx ?? 0) * 2),
      !!b.metaApplied,
      !!b.anchorShifted,
      b.noteBands?.length ?? 0,
    ])

const inputsOf = (blocks: BlockBox[], totalHeight: number): PassInputs => ({
  blocks,
  totalHeight,
  floats: [],
  sectBreaks: new Set(),
})

describe('canvas pass == preview pass over the corpus', () => {
  for (const stem of stems) {
    it(stem, async () => {
      const parsed = await parseDocx(new Uint8Array(readFileSync(join(CORPUS_DIR, `${stem}.docx`))))
      // canvas: one pipeline call over freshly measured blocks
      const canvasSide = buildCorpusBlocks(parsed)
      const canvas = slicePass(
        inputsOf(canvasSide.blocks, canvasSide.totalHeight),
        corpusCtx(parsed, canvasSide.geoms),
      )
      // preview: its own measure, block meta first, the shadow slices a twin
      const previewSide = buildCorpusBlocks(parsed)
      const ctx = corpusCtx(parsed, previewSide.geoms)
      const flow = prepareFlow(inputsOf(previewSide.blocks, previewSide.totalHeight), ctx)
      applyBlockMeta(flow.blocks, undefined, 1)
      const twin = { ...flow, blocks: [...flow.blocks], floats: flow.floats.map((f) => ({ ...f })) }
      const shadow = sliceFlow(twin, ctx)
      expect(pageCutDiffs(canvas.cuts, shadow.cuts), 'previewCutMismatch').toEqual([])
      expect(pageCutSignature(shadow.cuts)).toBe(pageCutSignature(canvas.cuts))
      expect(blockFacts(flow.blocks), 'block facts').toEqual(blockFacts(canvas.blocks))
      // pass.pageNotes: the pass places footnotes by pageOfY over its slices; the
      // preview used to redo it over the cuts, which share the slice starts
      for (const b of canvas.blocks)
        expect(pageOfY(canvas.cuts, b.top)).toBe(pageOfY(canvas.slices, b.top))
      // the pipeline is a fixed point: slicing the same blocks again moves nothing
      const again = sliceFlow(flow, { ...ctx, out: freshOut() })
      expect(pageCutDiffs(canvas.cuts, again.cuts)).toEqual([])
      // the preview reuses the pass's prepared flow
      // as is; the shadow twin over it yields the same cuts and leaves the pass's
      // blocks untouched
      const geomBefore = JSON.stringify(
        canvas.blocks.map((b) => [b.top, b.height, b.spaceBeforePx, b.leadFoldPx]),
      )
      const reuseCtx = corpusCtx(parsed, canvasSide.geoms)
      const twinOfPass = { ...canvas, blocks: [...canvas.blocks], floats: [] }
      expect(pageCutDiffs(canvas.cuts, sliceFlow(twinOfPass, reuseCtx).cuts)).toEqual([])
      expect(
        JSON.stringify(canvas.blocks.map((b) => [b.top, b.height, b.spaceBeforePx, b.leadFoldPx])),
      ).toBe(geomBefore)
    })
  }
})

/** a paragraph whose two footnote markers sit on different lines (marker boxes mocked) */
function noteRefPara(top: number): HTMLElement {
  const el = document.createElement('p')
  el.getBoundingClientRect = () => ({ top, height: 80 }) as DOMRect
  for (const dy of [10, 52]) {
    const sup = document.createElement('sup')
    sup.className = 'doc-note-ref'
    sup.setAttribute('data-note-ref', String(dy))
    sup.setAttribute('data-note-kind', 'footnote')
    sup.getBoundingClientRect = () =>
      ({ top: top + dy, bottom: top + dy + 8, height: 8 }) as DOMRect
    el.appendChild(sup)
  }
  return el
}

describe('block facts replay the skipped-block-meta regression', () => {
  const bands = [{ heightPx: 20 }, { heightPx: 40 }]
  const metaOf = (i: number) =>
    i === 1 ? { footnoteExtraPx: 60, footnoteBands: bands } : undefined
  // one canvas element: the canvas pass and the preview snapshot measure the same DOM
  const el = noteRefPara(100)
  const blocksOf = (): BlockBox[] => [
    { top: 0, height: 100, docxIndex: 0 },
    { top: 100, height: 80, docxIndex: 1, el },
    { top: 180, height: 100, docxIndex: 2 },
  ]
  const ctx = (): PassContext => ({
    sections: [],
    section: { pageHeight: 15840, marginBottom: 1440 } as SectionSettings,
    endnoteItems: [],
    endnoteSeparatorPx: 0,
    hfHeightsOf: () => [NO_HF],
    singleHfPx: NO_HF,
    singleGeom: { contentHeight: 200, forceBreak: false },
    evenOddHf: false,
    titlePg: false,
    dropCols: false,
    blockMetaOf: metaOf,
    factor: 1,
    out: freshOut(),
  })

  it('a preview that applies block meta up front matches the canvas blocks and cuts', () => {
    const canvas = slicePass(inputsOf(blocksOf(), 280), ctx())
    expect(canvas.blocks[1]).toMatchObject({ height: 140, footnoteExtraPx: 60 })
    const c = ctx()
    const flow = prepareFlow(inputsOf(blocksOf(), 280), c)
    applyBlockMeta(flow.blocks, metaOf, 1)
    expect(blockFacts(flow.blocks)).toEqual(blockFacts(canvas.blocks))
    expect(pageCutDiffs(canvas.cuts, sliceFlow(flow, c).cuts)).toEqual([])
    expect(flow.blocks[1].noteBands).toEqual([
      { offset: 14, height: 20 },
      { offset: 56, height: 40 },
    ])
  })

  it('the cuts index the pass floats the preview now reads', () => {
    const box = {
      el: document.createElement('div'),
      top: 150,
      height: 10,
      anchorTop: 150,
      pinned: false,
      pageRelV: false,
    }
    const canvas = slicePass({ ...inputsOf(blocksOf(), 280), floats: [box] }, ctx())
    expect(canvas.cuts.flatMap((c) => c.boxes)).toEqual([0])
    expect(canvas.floats[0]).toBe(box)
  })

  it('captures vertical-text inline extras through the shared hook on both sides', () => {
    const extraOf = (el: HTMLElement) => (el.tagName === 'P' ? 12 : 0)
    const canvas = slicePass(inputsOf(blocksOf(), 280), { ...ctx(), inlineExtraOf: extraOf })
    const flow = prepareFlow(inputsOf(blocksOf(), 280), { ...ctx(), inlineExtraOf: extraOf })
    expect(canvas.blocks.map((b) => b.inlineExtraPx)).toEqual([undefined, 12, undefined])
    expect(flow.blocks.map((b) => b.inlineExtraPx)).toEqual([undefined, 12, undefined])
    const plain = prepareFlow(inputsOf(blocksOf(), 280), ctx())
    expect(plain.blocks.every((b) => b.inlineExtraPx === undefined)).toBe(true)
  })

  it('a preview that skips block meta (the former production path) diverges', () => {
    const canvas = slicePass(inputsOf(blocksOf(), 280), ctx())
    const flow = prepareFlow(inputsOf(blocksOf(), 280), ctx())
    expect(blockFacts(flow.blocks)).not.toEqual(blockFacts(canvas.blocks))
    expect(flow.blocks[1]).toMatchObject({ height: 80 })
    expect(flow.blocks[1].metaApplied).toBeUndefined()
  })
})

describe('an empty document still gets a page', () => {
  // the preview prints nothing only without a pass or with zero cuts: a new
  // document (one empty paragraph, or no block at all) must yield one cut
  const ctx = (): PassContext => ({
    sections: [],
    section: { pageHeight: 15840, marginBottom: 1440 } as SectionSettings,
    endnoteItems: [],
    endnoteSeparatorPx: 0,
    hfHeightsOf: () => [NO_HF],
    singleHfPx: NO_HF,
    singleGeom: { contentHeight: 800, forceBreak: false },
    evenOddHf: false,
    titlePg: false,
    dropCols: false,
    factor: 1,
    out: freshOut(),
  })

  it('one empty paragraph slices to one page', () => {
    const one = slicePass(
      {
        blocks: [{ top: 0, height: 18, emptyPara: true }],
        totalHeight: 18,
        floats: [],
        sectBreaks: new Set(),
      },
      ctx(),
    )
    expect(one.cuts.map((c) => [c.start, c.end])).toEqual([[0, 18]])
  })

  it('an empty flow slices to one blank page', () => {
    const none = slicePass({ blocks: [], totalHeight: 0, floats: [], sectBreaks: new Set() }, ctx())
    expect(none.cuts.map((c) => [c.start, c.end])).toEqual([[0, 0]])
  })
})

describe('the preview pulls its floats up like the pass (P24)', () => {
  // the legacy text-anchor fixture of pagination.test: the drawing passes the
  // body bottom and is pulled up to the paper edge (probe 2026-10-13)
  const geom: SectionGeom = {
    contentHeight: 800,
    contentWidth: 600,
    marginBottomPx: 60,
    forceBreak: false,
  }
  const blocksOf = (): BlockBox[] => [
    { top: 0, height: 700 },
    { top: 700, height: 20, floatKeepWhole: true, floatBottom: 1115, floatMaxH: 400 },
    { top: 720, height: 20 },
  ]
  // the drawing reaches below the last page's bottom margin: a spill block opens
  // a page until the pull-up (-255) brings it back onto the first page
  const floatsOf = () => [
    {
      el: document.createElement('div'),
      top: 715,
      height: 400,
      anchorTop: 700,
      pinned: false,
      pageRelV: false,
    },
  ]
  const ctx = (): PassContext => ({
    sections: [],
    section: { pageHeight: 15840, marginBottom: 1440 } as SectionSettings,
    endnoteItems: [],
    endnoteSeparatorPx: 0,
    hfHeightsOf: () => [NO_HF],
    singleHfPx: NO_HF,
    singleGeom: geom,
    evenOddHf: false,
    titlePg: false,
    dropCols: false,
    factor: 1,
    out: freshOut(),
  })

  it('applies the pass floatPullUps to the snapshot floats; without them the box stays low', () => {
    const c = ctx()
    const canvas = slicePass(
      { blocks: blocksOf(), totalHeight: 740, floats: floatsOf(), sectBreaks: new Set() },
      c,
    )
    expect(c.out.floatPullUps).toEqual([{ blockTop: 700, dyPx: -255 }])
    expect(canvas.floats[0].top).toBe(460)
    const preview = floatsOf()
    expect(preview[0].top).toBe(715)
    applyFloatPullUps(preview, c.out.floatPullUps!)
    expect(preview[0].top).toBe(canvas.floats[0].top)
    // pinned / page-relative boxes never move with the anchor
    const pinned = [
      { ...floatsOf()[0], pinned: true },
      { ...floatsOf()[0], pageRelV: true },
    ]
    applyFloatPullUps(pinned, c.out.floatPullUps!)
    expect(pinned.map((f) => f.top)).toEqual([715, 715])
  })

  it('a twin of the pass flow re-slices to the same cuts', () => {
    const inputs = () => ({
      blocks: blocksOf(),
      totalHeight: 740,
      floats: floatsOf(),
      sectBreaks: new Set<number>(),
    })
    const prepared = prepareFlow(inputs(), ctx())
    const canvas = slicePass(inputs(), ctx())
    expect(prepared.flowH).toBeGreaterThan(canvas.flowH)
    expect(canvas.cuts).toHaveLength(1)
    const twin = { ...canvas, blocks: [...canvas.blocks], floats: floatsOf() }
    const c2 = ctx()
    expect(pageCutDiffs(canvas.cuts, sliceFlow(twin, c2).cuts)).toEqual([])
  })

  it('the shadow twin undoes the pass pull-ups on copies of the pass floats first', () => {
    const c = ctx()
    const canvas = slicePass(
      { blocks: blocksOf(), totalHeight: 740, floats: floatsOf(), sectBreaks: new Set() },
      c,
    )
    expect(canvas.floats[0].top).toBe(460)
    const twinFloats = canvas.floats.map((f) => ({ ...f }))
    applyFloatPullUps(
      twinFloats,
      c.out.floatPullUps!.map((p) => ({ ...p, dyPx: -p.dyPx })),
    )
    expect(twinFloats[0].top).toBe(715)
    const twin = { ...canvas, blocks: [...canvas.blocks], floats: twinFloats }
    expect(pageCutDiffs(canvas.cuts, sliceFlow(twin, ctx()).cuts)).toEqual([])
    expect(twinFloats[0].top).toBe(460)
    // the pass floats themselves were not touched
    expect(canvas.floats[0].top).toBe(460)
  })
})
