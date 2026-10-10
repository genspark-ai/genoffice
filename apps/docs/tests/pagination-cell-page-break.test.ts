import { describe, expect, it } from 'vitest'
import { cellCutYs, forcedCutYs } from '../src/renderer/pagination-lines'

/**
 * An explicit page break inside a table cell has to become a cut point.
 *
 * The break is the author's instruction, not a heuristic: Word turns the page at
 * that line, and so must this. Without it the break renders its marker (the
 * `br.doc-page-br` line) but the row stays whole, so the page never turns and
 * the new page is missing from print and PDF export alike.
 */
describe('cellCutYs — forced page-break cuts', () => {
  it('emits a cut at an explicit break even where no line boundary falls', () => {
    // one 300px text line: there is no line-to-line midpoint, so the heuristic
    // alone can only produce "no cuts"
    const bands: Array<Array<[number, number]>> = [[[0, 300]]]
    expect(cellCutYs(bands, 300)).toEqual([])

    // the break sits mid-line at y=150
    expect(cellCutYs(bands, 300, undefined, [150])).toEqual([150])
  })

  it('keeps the authored breaks even when they sit inside a text line band', () => {
    // two bands whose midpoint heuristic would cut at 100; the author asked at 60
    const bands: Array<Array<[number, number]>> = [
      [
        [0, 60],
        [60, 200],
      ],
    ]
    expect(cellCutYs(bands, 300, undefined, [60])).toEqual([60])
  })

  it('adds the breaks to the ones the line boundaries already allow', () => {
    // three lines, so the heuristic already cuts at 100 and 200; a break at 40
    // lands inside the first line and is only reachable because it is forced
    const bands: Array<Array<[number, number]>> = [
      [
        [0, 100],
        [100, 200],
        [200, 300],
      ],
    ]
    expect(cellCutYs(bands, 300)).toEqual([100, 200])
    expect(cellCutYs(bands, 300, undefined, [200, 40, 200])).toEqual([40, 100, 200])
  })

  it('ignores breaks that fall on the row edge, which are not cuts', () => {
    const bands: Array<Array<[number, number]>> = [[[0, 300]]]
    expect(cellCutYs(bands, 300, undefined, [0, 300])).toEqual([])
  })

  it('does not invent cuts when there is no break to honour', () => {
    const bands: Array<Array<[number, number]>> = [
      [
        [0, 100],
        [100, 300],
      ],
    ]
    expect(cellCutYs(bands, 300, undefined, [])).toEqual([100])
  })
})

/** the DOM half, as a pure function over rects so jsdom's lack of layout is moot */
describe('forcedCutYs — reads the break positions', () => {
  const EL_TOP = 1000
  const ROW_TOP = 0
  const ZOOM = 1
  const noGaps = () => 0
  const rect = (top: number, height: number) =>
    ({
      top,
      bottom: top + height,
      left: 0,
      right: 100,
      width: 100,
      height,
      x: 0,
      y: top,
    }) as DOMRect

  it('reports a break at its top edge, collapsed box or not', () => {
    // a <br> has no height of its own, so the collapsed case is the normal one
    expect(forcedCutYs([{ rect: rect(EL_TOP + 150, 0) }], EL_TOP, ROW_TOP, noGaps, ZOOM)).toEqual([
      150,
    ])
    expect(forcedCutYs([{ rect: rect(EL_TOP + 150, 12) }], EL_TOP, ROW_TOP, noGaps, ZOOM)).toEqual([
      150,
    ])
  })

  it('subtracts the page gaps above it and rescales by the zoom', () => {
    // one 40px gap sits above the break
    const gaps = (top: number) => (top >= EL_TOP + 100 ? 40 : 0)
    expect(forcedCutYs([{ rect: rect(EL_TOP + 140, 10) }], EL_TOP, ROW_TOP, gaps, ZOOM)).toEqual([
      100,
    ])
    expect(forcedCutYs([{ rect: rect(EL_TOP + 200, 10) }], EL_TOP, ROW_TOP, noGaps, 2)).toEqual([
      100,
    ])
  })

  it('keeps every break, in document order', () => {
    const breaks = [{ rect: rect(EL_TOP + 200, 8) }, { rect: rect(EL_TOP + 80, 8) }]
    const ys = forcedCutYs(breaks, EL_TOP, ROW_TOP, noGaps, ZOOM)
    expect([...ys].sort((a, b) => a - b)).toEqual([80, 200])
  })
})
