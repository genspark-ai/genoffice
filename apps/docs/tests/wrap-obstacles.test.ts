import { describe, expect, it } from 'vitest'
import { bandGapRows, placeLine, type ObstacleRect } from '../src/renderer/editor/wrap-obstacles'
import { computeSectionedSlicesF2, type BlockBox } from '../src/renderer/pagination'

// column 0..624 px (468pt), 200pt picture = 266.7 px, 9pt clearance = 12 px
const COL = 624
const LINE = 17.9
const box = (left: number, width: number, top: number, height: number): ObstacleRect => ({
  left: left - 12,
  right: left + width + 12,
  top,
  bottom: top + height,
})

describe('placeLine (Word probe 2026-10-13)', () => {
  it('keeps a line that crosses no rectangle at full width', () => {
    expect(placeLine([box(0, 266.7, 100, 133)], 0, LINE, 0, COL, 120)).toEqual({
      top: 0,
      left: 0,
      right: COL,
    })
  })

  it('takes the widest gap that holds the first word; any overlap counts', () => {
    // picture at the column left, line overlapping it by 2 px: text goes right of it
    const p = placeLine([box(0, 266.7, 50, 133)], 50 - LINE + 2, LINE, 0, COL, 120)
    expect(p.top).toBeCloseTo(50 - LINE + 2)
    expect(p.left).toBeCloseTo(266.7 + 12)
    expect(p.right).toBe(COL)
  })

  it('a 40pt gap holds a 25pt word but not a 90pt one', () => {
    const rects = [box(COL - 53.3 - 266.7 - 12, 266.7, 20, 133)]
    const short = placeLine(rects, 20, LINE, 0, COL, 33)
    expect(short.top).toBe(20)
    expect(short.left).toBe(0)
    const long = placeLine(rects, 20, LINE, 0, COL, 120)
    expect(long.top).toBe(20)
    expect(long.left).toBe(0)
    expect(long.right).toBeCloseTo(COL - 53.3 - 266.7 - 24)
  })

  it('a blocked line lands at the blocking rectangle bottom, not on the line grid', () => {
    const wide = { left: -12, right: COL + 12, top: 272, bottom: 332 }
    const p = placeLine([wide], 260, LINE, 0, COL, 60)
    expect(p).toEqual({ top: 332, left: 0, right: COL })
  })

  it('steps past a second row when the first row bottom is still blocked', () => {
    const rects = [
      box(0, 200, 0, 300),
      box(230, 200, 0, 300),
      box(460, 200, 0, 280),
      box(0, 200, 310, 300),
      box(230, 200, 330, 300),
    ]
    // first row leaves 6 px slivers: the heading drops to the end of the
    // shortest first-row box (280) and finds the 150 px gap right of the second row
    const p = placeLine(rects, 10, LINE, 0, COL, 140)
    expect(p.top).toBe(280)
    expect(p.left).toBeCloseTo(430 + 12)
    expect(p.right).toBe(COL)
  })

  it('an empty line fits any gap wider than a pixel', () => {
    const rects = [box(0, 300, 0, 100), box(360, 264, 0, 100)]
    const p = placeLine(rects, 0, LINE, 0, COL, 0)
    expect(p.top).toBe(0)
    expect(p.left).toBeCloseTo(312)
    expect(p.right).toBeCloseTo(348)
  })
})

describe('anchor paragraph page turn (Word probe 2026-10-13)', () => {
  const cap = 900
  const geoms = [{ contentHeight: cap, forceBreak: false }]
  const anchor = (top: number, bottom: number, maxH: number): BlockBox => ({
    top,
    height: 20,
    floatKeepWhole: true,
    floatBottom: top + bottom,
    floatMaxH: maxH,
  })

  it('turns the page when the tallest drawing cannot end at the body bottom', () => {
    const blocks: BlockBox[] = [
      { top: 0, height: 800 },
      anchor(800, 400, 250),
      { top: 820, height: 20 },
    ]
    const slices = computeSectionedSlicesF2(blocks, geoms, 840)
    expect(slices.map((s) => s.start)).toEqual([0, 800])
  })

  it('stays when the drawing fits once pulled up to the body bottom', () => {
    const blocks: BlockBox[] = [
      { top: 0, height: 800 },
      anchor(800, 150, 90),
      { top: 820, height: 20 },
    ]
    expect(computeSectionedSlicesF2(blocks, geoms, 840)).toHaveLength(1)
  })

  it('stays when the drawings end above the body bottom', () => {
    const blocks: BlockBox[] = [
      { top: 0, height: 500 },
      anchor(500, 300, 300),
      { top: 520, height: 20 },
    ]
    expect(computeSectionedSlicesF2(blocks, geoms, 540)).toHaveLength(1)
  })

  it('never turns a page it opens', () => {
    const blocks: BlockBox[] = [anchor(0, 1200, 1200), { top: 20, height: 20 }]
    expect(computeSectionedSlicesF2(blocks, geoms, 40)[0].start).toBe(0)
  })
})

describe('bandGapRows (prod corpus 20261020 sample 026, stat tiles above a title bar)', () => {
  // anchor line 6 px; hairline at 8, five tiles 18..84, title bar 94..127
  const bands: Array<[number, number]> = [
    [8, 9],
    [18, 84],
    [18, 84],
    [18, 84],
    [18, 84],
    [18, 84],
    [94, 127],
  ]

  it('merges overlapping tiles and keeps the rows between the bands', () => {
    expect(bandGapRows(bands, 6)).toEqual([
      [8, 9],
      [18, 84],
      [94, 127],
    ])
  })

  it('a single band hugging the anchor line keeps the union-band route', () => {
    expect(bandGapRows([[0, 60]], 20)).toBeNull()
    expect(bandGapRows([[22, 60]], 20)).toBeNull()
  })

  it('a band starting well below the anchor line leaves a usable row', () => {
    expect(bandGapRows([[40, 60]], 20)).toEqual([[40, 60]])
  })

  it('ignores degenerate intervals', () => {
    expect(
      bandGapRows(
        [
          [10, 10],
          [Number.NaN, 5],
        ],
        6,
      ),
    ).toBeNull()
  })
})
