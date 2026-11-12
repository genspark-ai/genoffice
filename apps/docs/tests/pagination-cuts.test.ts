// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import {
  applyGapLedger,
  buildPageCuts,
  flowOffsetAt,
  learnGapFlowTops,
  headingPagesOf,
  pageCutDiffs,
  pageCutSignature,
  pageOwners,
  pageLeadsOf,
  pagePaperEdges,
  topPushAt,
  type GapLedgerEntry,
} from '../src/renderer/pagination-cuts'
import type { BlockBox, PageSlice, SectionGeom } from '../src/renderer/pagination-types'

const para = (top: number, height: number, extra: Partial<BlockBox> = {}): BlockBox => ({
  top,
  height,
  el: document.createElement('p'),
  ...extra,
})
const table = (top: number, heights: number[], extra: Partial<BlockBox> = {}): BlockBox => ({
  top,
  height: heights.reduce((a, b) => a + b, 0),
  el: document.createElement('table'),
  tableRows: heights.map((h) => ({ height: h })),
  ...extra,
})
const slice = (start: number, end: number, section = 0): PageSlice => ({ start, end, section })
const geom = (contentHeight: number, extra: Partial<SectionGeom> = {}): SectionGeom => ({
  contentHeight,
  forceBreak: false,
  ...extra,
})

describe('buildPageCuts', () => {
  it('classifies the lead of every page', () => {
    const blocks = [
      para(0, 300),
      para(300, 500),
      table(800, [100, 100, 100]),
      para(1100, 400),
      para(1500, 100),
    ]
    const slices = [
      slice(0, 300),
      slice(300, 800),
      slice(800, 950),
      slice(950, 1300),
      slice(1300, 1600),
    ]
    const cuts = buildPageCuts(blocks, slices, [geom(500)], [{ headerPx: 0, footerPx: 0 }])
    expect(cuts.map((c) => c.lead)).toEqual([
      { block: 0, kind: 'block' },
      { block: 1, kind: 'block' },
      { block: 2, kind: 'block' },
      { block: 2, kind: 'cell', row: 1, lineOffset: 150 },
      { block: 3, kind: 'line', lineOffset: 200 },
    ])
    expect(cuts.map((c) => c.pageNo)).toEqual([1, 2, 3, 4, 5])
    expect(cuts.map((c) => c.padPx)).toEqual([200, 0, 350, 150, 200])
  })

  it('row lead tolerates the DOM row match window and respects folded space-before', () => {
    const blocks = [para(0, 100), table(100, [200, 200], { spaceBeforePx: 10 })]
    const cuts = buildPageCuts(
      blocks,
      [slice(0, 311), slice(311, 500)],
      [geom(400)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    expect(cuts[1].lead).toEqual({ block: 1, kind: 'row', row: 1 })
  })

  it('a block lead drops its space before after an explicit break or by the plan, never on a section start', () => {
    const blocks = [
      para(0, 400),
      para(400, 100, { breakBefore: true }),
      para(500, 100, { dropAtPageTop: true }),
      para(600, 100, { dropAtPageTop: false }),
      para(700, 100, { section: 1 }),
      para(800, 100, { section: 1 }),
    ]
    const slices = [
      slice(0, 400),
      slice(400, 500),
      slice(500, 600),
      slice(600, 700),
      slice(700, 800, 1),
      slice(800, 900, 1),
    ]
    const cuts = buildPageCuts(
      blocks,
      slices,
      [geom(400), geom(400)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    expect(cuts.map((c) => c.leadDropsSpaceBefore)).toEqual([false, true, true, false, false, true])
  })

  it('a folded legacy pageBreakBefore lead keeps its trimmed margin', () => {
    const blocks = [para(0, 400), para(390, 110, { breakBefore: true, leadMtPx: 10 })]
    const cuts = buildPageCuts(
      blocks,
      [slice(0, 390), slice(390, 500)],
      [geom(400)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    expect(cuts[1].leadDropsSpaceBefore).toBe(false)
    expect(cuts[1].leadMtPx).toBe(10)
  })

  it('pad and pull-up follow the physical height of a regioned page and the body height of a vertical one', () => {
    const blocks = [para(0, 300), para(300, 300)]
    const cuts = buildPageCuts(
      blocks,
      [{ start: 0, end: 300, section: 0, regions: [], physHeight: 200 }, slice(300, 600, 1)],
      [geom(500), geom(900, { vertical: 'tbRl', contentWidth: 400 })],
      [{ headerPx: 0, footerPx: 0 }],
    )
    expect(cuts[0]).toMatchObject({ usedPx: 300, physUsedPx: 200, padPx: 300 })
    expect(cuts[1]).toMatchObject({ capacityPx: 900, padPx: 100 })
  })

  it('vertical pad follows the header variant of the page (titlePg, even)', () => {
    const blocks = [para(0, 300), para(300, 300), para(600, 300)]
    const cuts = buildPageCuts(
      blocks,
      [slice(0, 300), slice(300, 600), slice(600, 900)],
      [
        geom(900, {
          vertical: 'tbRl',
          contentWidth: 400,
          firstBodyHeight: 350,
          evenBodyHeight: 380,
        }),
      ],
      [{ headerPx: 0, footerPx: 0, firstHeaderPx: 50, evenHeaderPx: 20 }],
      { evenOddHf: true, titlePgOf: () => true },
    )
    expect(cuts.map((c) => c.hf.variant)).toEqual(['first', 'even', 'default'])
    expect(cuts.map((c) => c.padPx)).toEqual([50, 80, 100])
  })

  it('trailing blank page has a tail lead and is blank', () => {
    const blocks = [para(0, 100, { breakAfter: true })]
    const cuts = buildPageCuts(
      blocks,
      [slice(0, 100), slice(100, 100)],
      [geom(400)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    expect(cuts[1].lead).toEqual({ block: -1, kind: 'tail' })
    expect(cuts[1].blank).toBe(true)
    expect(cuts[0].blank).toBe(false)
  })

  it('capacity, header/footer variant and page numbers follow the section geometry', () => {
    const geoms = [
      geom(500, { firstContentHeight: 450 }),
      geom(500, { evenContentHeight: 480, pageNumberStart: 4 }),
    ]
    const hf = [
      { headerPx: 20, footerPx: 20, firstHeaderPx: 70, firstFooterPx: 20 },
      { headerPx: 20, footerPx: 20, evenHeaderPx: 40, evenFooterPx: 20 },
    ]
    const blocks = [para(0, 450), para(450, 500), para(950, 500), para(1450, 100)]
    const slices = [slice(0, 450), slice(450, 950, 0), slice(950, 1430, 1), slice(1430, 1550, 1)]
    const cuts = buildPageCuts(blocks, slices, geoms, hf, {
      evenOddHf: true,
      titlePgOf: (s) => s === 0,
      pageNumberFmtOf: (s) => (s === 1 ? 'lowerRoman' : undefined),
    })
    expect(cuts.map((c) => c.pageNo)).toEqual([1, 2, 4, 5])
    expect(cuts.map((c) => c.display)).toEqual(['1', '2', 'iv', 'v'])
    expect(cuts.map((c) => c.hf.variant)).toEqual(['first', 'even', 'even', 'default'])
    expect(cuts.map((c) => c.hf.headerPx)).toEqual([70, 20, 40, 20])
    expect(cuts.map((c) => c.capacityPx)).toEqual([450, 500, 480, 500])
  })

  it('names the block each block-kind gap widget opens a page in front of', () => {
    const blocks = [para(0, 400), para(400, 300), table(700, [100, 100]), para(900, 100)]
    const cuts = buildPageCuts(
      blocks,
      [slice(0, 400), slice(400, 800), slice(800, 1000)],
      [geom(400)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    const gap = { heightPx: 0, marginBottomPx: 0, marginTopPx: 0, pullUpPx: 0, repeatHeaderPx: 0 }
    applyGapLedger(cuts, [
      { page: 1, kind: 'block', ...gap },
      { page: 2, kind: 'table', ...gap },
    ] as GapLedgerEntry[])
    expect(cuts.map((c) => c.lead.kind)).toEqual(['block', 'block', 'row'])
    // the row lead resets the wrap obstacles in front of the block after the table
    expect([...pageLeadsOf(cuts, blocks)]).toEqual([
      [blocks[1].el, 1],
      [blocks[3].el, 2],
    ])
    applyGapLedger(cuts, [])
    expect(pageLeadsOf(cuts, blocks).size).toBe(0)
  })

  it('in-block leads (row / cell / line) reset in front of the next block; the last turn wins (B4)', () => {
    const gap = { heightPx: 0, marginBottomPx: 0, marginTopPx: 0, pullUpPx: 0, repeatHeaderPx: 0 }
    // a paragraph split by a line lead, then a table straddling two turns
    const blocks = [para(0, 500), para(500, 100), table(600, [100, 100, 100, 100]), para(1000, 100)]
    const cuts = buildPageCuts(
      blocks,
      [slice(0, 400), slice(400, 700), slice(700, 900), slice(900, 1100)],
      [geom(400)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    expect(cuts.map((c) => c.lead.kind)).toEqual(['block', 'line', 'row', 'row'])
    applyGapLedger(cuts, [
      { page: 1, kind: 'inline', ...gap },
      { page: 2, kind: 'table', ...gap },
      { page: 3, kind: 'table', ...gap },
    ] as GapLedgerEntry[])
    expect([...pageLeadsOf(cuts, blocks)]).toEqual([
      [blocks[1].el, 1],
      [blocks[3].el, 3],
    ])
    // a split floating table's in-table turns move only the float: no reset
    applyGapLedger(cuts, [
      { page: 2, kind: 'table', floatGap: true, ...gap },
      { page: 3, kind: 'table', floatGap: true, ...gap },
    ] as GapLedgerEntry[])
    expect(pageLeadsOf(cuts, blocks).size).toBe(0)
  })

  it('pins floated blocks to the page their anchor lands on', () => {
    const blocks = [para(0, 400), para(400, 400), para(420, 50, { floated: true })]
    const cuts = buildPageCuts(
      blocks,
      [slice(0, 400), slice(400, 800)],
      [geom(400)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    expect(cuts.map((c) => c.floats)).toEqual([[], [2]])
  })

  it('names the page of every anchor-shifted block and floating box', () => {
    const blocks = [para(0, 400, { anchorShifted: true }), para(400, 400), para(420, 50)]
    const box = (anchorTop: number) => ({
      el: document.createElement('div'),
      top: anchorTop + 30,
      height: 20,
      anchorTop,
      pinned: false,
      pageRelV: false,
    })
    const cuts = buildPageCuts(
      blocks,
      [slice(0, 400), slice(400, 800)],
      [geom(400)],
      [{ headerPx: 0, footerPx: 0 }],
      { floats: [box(410), box(10), box(-5), box(900)] },
    )
    expect(cuts.map((c) => c.anchored)).toEqual([[0], []])
    expect(cuts.map((c) => c.boxes)).toEqual([
      [1, 2],
      [0, 3],
    ])
    const owners = pageOwners(cuts)
    expect([...owners.blocks]).toEqual([[0, 0]])
    expect([...owners.boxes].sort()).toEqual([
      [0, 1],
      [1, 0],
      [2, 0],
      [3, 1],
    ])
    expect(pageCutDiffs(cuts, cuts)).toEqual([])
    const other = buildPageCuts(
      blocks,
      [slice(0, 400), slice(400, 800)],
      [geom(400)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    expect(pageCutDiffs(cuts, other)).toEqual(['p1 boxes 1,2 vs ', 'p2 boxes 0,3 vs '])
  })
})

describe('page windows and signatures', () => {
  it('carries the seam flags, lifted tops and lead spills the sheets apply', () => {
    const blocks = [
      para(0, 300, { lineLeadPx: 6 }),
      table(300, [100, 100, 100, 100]),
      para(700, 200, { liftPx: 40 }),
      para(900, 100),
    ]
    const slices = [slice(0, 500), slice(500, 720), slice(720, 1000)]
    const cuts = buildPageCuts(blocks, slices, [geom(500)], [{ headerPx: 0, footerPx: 0 }])
    expect(cuts[1].cutTable).toBe(true)
    expect(cuts[2].liftTop).toBe(20)
    expect(cuts.map((c) => c.leadSpill)).toEqual([undefined, undefined, undefined])
    expect(buildPageCuts(blocks, slices, [geom(500)], [{ headerPx: 0, footerPx: 0 }])).toEqual(cuts)
  })

  it('signs every page fact a sheet depends on', () => {
    const blocks = [para(0, 300), para(300, 300)]
    const a = buildPageCuts(
      blocks,
      [slice(0, 300), slice(300, 600)],
      [geom(300)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    const b = buildPageCuts(
      blocks,
      [slice(0, 300), slice(300, 600)],
      [geom(300)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    expect(pageCutSignature(a)).toBe(pageCutSignature(b))
    expect(pageCutSignature(a).split('|')).toHaveLength(2)
    const moved = buildPageCuts(
      blocks,
      [slice(0, 300), slice(300, 601)],
      [geom(301)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    expect(pageCutSignature(moved)).not.toBe(pageCutSignature(a))
    expect(pageCutSignature([])).toBe('')
  })

  it('reports where two cut lists disagree, within half a pixel', () => {
    const blocks = [para(0, 300), para(300, 300), para(600, 300)]
    const hf = [{ headerPx: 0, footerPx: 0 }]
    const a = buildPageCuts(
      blocks,
      [slice(0, 300), slice(300, 600), slice(600, 900)],
      [geom(300)],
      hf,
    )
    const near = buildPageCuts(
      blocks,
      [slice(0, 300.2), slice(300.2, 600), slice(600, 900)],
      [geom(300)],
      hf,
    )
    expect(pageCutDiffs(a, near)).toEqual([])
    const b = buildPageCuts(blocks, [slice(0, 300), slice(300, 900)], [geom(600)], hf)
    const diffs = pageCutDiffs(a, b)
    expect(diffs[0]).toBe('pages 3 vs 2')
    expect(diffs).toContain('p2 end 600 vs 900')
    expect(diffs.some((d) => d.startsWith('p2 capacity'))).toBe(true)
  })
})

describe('headingPagesOf', () => {
  it('maps a restamped heading node to its page by data-idx', () => {
    const h1 = para(0, 100, { docxIndex: 0 })
    const h2 = para(600, 100, { docxIndex: 7 })
    h1.el!.dataset.idx = '0'
    h2.el!.dataset.idx = '7'
    const blocks = [h1, para(100, 500, { docxIndex: 3 }), h2, para(700, 300, { docxIndex: 9 })]
    const cuts = buildPageCuts(
      blocks,
      [slice(0, 500), slice(500, 1000)],
      [geom(500)],
      [{ headerPx: 0, footerPx: 0 }],
    )
    const restamped = document.createElement('p')
    restamped.dataset.idx = '7'
    const stray = document.createElement('p')
    expect(headingPagesOf(blocks, cuts, [h1.el!, restamped])).toEqual({ pages: [1, 2], misses: 0 })
    expect(headingPagesOf(blocks, cuts, [stray, null])).toEqual({ pages: [1, 1], misses: 2 })
  })
})

describe('gap ledger on the cuts', () => {
  const entry = (page: number, extra: Partial<GapLedgerEntry> = {}): GapLedgerEntry => ({
    page,
    kind: 'block',
    heightPx: 220,
    marginBottomPx: 96,
    marginTopPx: 96,
    marginLeftPx: 90,
    topPushPx: 0,
    pullUpPx: 0,
    repeatHeaderPx: 0,
    ...extra,
  })
  const cutsOf = () =>
    buildPageCuts(
      [para(0, 400), para(400, 400), para(800, 400), para(1200, 400), para(1600, 400)],
      [slice(0, 400), slice(400, 800), slice(800, 1200), slice(1200, 1600), slice(1600, 2000)],
      [geom(400)],
      [{ headerPx: 0, footerPx: 0 }],
    )

  it('accumulates page-turn gaps for the geometry and flow gaps for the displacement', () => {
    const cuts = cutsOf()
    applyGapLedger(cuts, [
      entry(1, { topPushPx: 26 }),
      entry(2, { kind: 'table', floatGap: true, repeatHeaderPx: 30 }),
      entry(3, { kind: 'table', repeatHeaderPx: 30, pullUpPx: 20 }),
      entry(4, {
        kind: 'cut',
        heightPx: 0,
        marginBottomPx: 0,
        marginTopPx: 0,
        topPushPx: undefined,
      }),
    ])
    expect(cuts.map((c) => c.gapOffsetPx)).toEqual([0, 220, 440, 640, 640])
    expect(cuts.map((c) => c.flowOffsetPx)).toEqual([0, 220, 220, 450, 450])
    expect(cuts[0].gap).toBeUndefined()
    expect(topPushAt(cuts, 4)).toBe(0)
    expect(topPushAt(cuts, 1)).toBe(26)
    // a carry spacer counts for flow y at or below it only
    const carries = [{ y: 900, px: 220 }]
    expect(flowOffsetAt(cuts, carries, 850)).toBe(220)
    expect(flowOffsetAt(cuts, carries, 900)).toBe(440)
    expect(flowOffsetAt(cuts, carries, 1300)).toBe(670)
  })

  it('paper edges follow the gap boxes; pages without a gap inherit the offsets', () => {
    const cuts = cutsOf()
    applyGapLedger(cuts, [entry(1), entry(3, { pullUpPx: 20 })])
    expect(cuts.map((c) => c.gapOffsetPx)).toEqual([0, 220, 220, 420, 420])
    const edges = pagePaperEdges(cuts, 96)
    expect(edges.map((e) => e.page)).toEqual([1, 3])
    // gap bottom = first margin + start + offset; paper top is a top margin above it
    expect(edges[0]).toMatchObject({
      gapBottom: 96 + 400 + 220,
      top: 400 + 220,
      prevBottom: 96 + 400 + 96,
    })
    expect(edges[1]).toMatchObject({ gapBottom: 96 + 1200 + 420, top: 1200 + 420 })
  })

  it('learned widget tops move the paper edges and the flow displacement off the slice boundary', () => {
    const cuts = cutsOf()
    applyGapLedger(cuts, [entry(1), entry(3, { pullUpPx: 20 })])
    // page 2's widget sits 10 px under the boundary (the previous block's space after)
    expect(learnGapFlowTops(cuts, [410, 1200 + 220 - 20])).toBe(true)
    expect(cuts[1].gap?.flowTopPx).toBe(410)
    expect(cuts[3].gap?.flowTopPx).toBe(1180)
    const edges = pagePaperEdges(cuts, 96)
    expect(edges[0]).toMatchObject({ gapTop: 96 + 410, gapBottom: 96 + 410 + 220 })
    expect(edges[1]).toMatchObject({ gapBottom: 96 + 1200 + 420, top: 1200 + 420 })
    expect(flowOffsetAt(cuts, [], 405)).toBe(0)
    expect(flowOffsetAt(cuts, [], 410)).toBe(220)
    expect(learnGapFlowTops(cuts, [410])).toBe(false)
  })
})
