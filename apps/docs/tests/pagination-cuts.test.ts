// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import {
  buildPageCuts,
  headingPagesOf,
  pageCutDiffs,
  pageCutSignature,
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
