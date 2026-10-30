import { describe, expect, it } from 'vitest'
import {
  floatKeepWholeOf,
  layoutHeightOf,
  obstacleBounds,
  obstacleFactsOf,
  turnedBounds,
} from '../src/renderer/editor/obstacle-plan'

const EMU = 9525
const box = (o: Record<string, unknown>) => ({
  paras: [],
  floating: true,
  wrapSides: true,
  widthPx: 200,
  offsetXEmu: 50 * EMU,
  offsetYEmu: 10 * EMU,
  wrapDistPx: { l: 12, r: 12, t: 0, b: 0 },
  ...o,
})
const COL = { w: 624, marginLeft: 96, marginRight: 96 }

describe('obstacleFactsOf', () => {
  it('is undefined for inline boxes and page-pinned cover art', () => {
    expect(obstacleFactsOf({ textboxes: [box({ floating: false })] })).toBeUndefined()
    expect(obstacleFactsOf({ textboxes: [box({ pagePinned: true })] })).toBeUndefined()
    expect(obstacleFactsOf({ textboxes: null })).toBeUndefined()
  })

  it('places a fixed-height side box at its offset plus the 2px box margin', () => {
    const f = obstacleFactsOf({ textboxes: [box({ heightPx: 100 })] }, COL)!
    expect(f.sides).toEqual([
      { x: 50, y: 12, w: 200, h: 100, dist: [12, 12, 0, 0], pageRelV: false },
    ])
    expect(f.allSides).toBe(true)
    expect(f.bandKeep).toBe(false)
    expect(f.bands).toEqual([])
  })

  it('an auto-fit box takes the live height; a collapsed one has none', () => {
    const attrs = { textboxes: [box({}), box({ offsetXEmu: 0 })] }
    const f = obstacleFactsOf(attrs, COL, [80, undefined])!
    expect(f.sides.map((s) => s.h)).toEqual([80, undefined])
  })

  it('clamps a column X onto the paper like the renderer', () => {
    const f = obstacleFactsOf(
      { textboxes: [box({ offsetXEmu: -300 * EMU }), box({ offsetXEmu: 900 * EMU })] },
      COL,
    )!
    expect(f.sides.map((s) => s.x)).toEqual([-96, 624 - 200 + 96])
    const raw = obstacleFactsOf({ textboxes: [box({ offsetXEmu: -300 * EMU })] })!
    expect(raw.sides[0].x).toBe(-300)
  })

  it('a wrapNone sibling makes the anchor not all-sides; bandOverflow is band-keep', () => {
    const f = obstacleFactsOf({
      textboxes: [box({ heightPx: 50 }), box({ wrapSides: false, bandOverflow: true })],
    })!
    expect(f.sides).toHaveLength(1)
    expect(f.allSides).toBe(false)
    expect(f.bandKeep).toBe(true)
  })

  it('wrapTopAndBottom boxes yield the data-bands intervals', () => {
    const f = obstacleFactsOf({
      textboxes: [box({ wrapSides: false, bandTopPx: 10.4, heightPx: 30 })],
    })!
    expect(f.bands).toEqual([[10, 40]])
  })

  it('a page-positioned box is marked and left out of the keep-whole rule', () => {
    const f = obstacleFactsOf({
      textboxes: [box({ heightPx: 300, pageRelV: true }), box({ heightPx: 40 })],
    })!
    expect(f.sides.map((s) => s.pageRelV)).toEqual([true, false])
    expect(floatKeepWholeOf(f)).toEqual({ bottom: 52, maxH: 40 })
  })
})

describe('floatKeepWholeOf', () => {
  it('reports the lowest bottom and the tallest box, rounded like the data attrs', () => {
    const f = obstacleFactsOf({
      textboxes: [box({ heightPx: 100.26 }), box({ offsetYEmu: 30 * EMU, heightPx: 60 })],
    })!
    expect(floatKeepWholeOf(f)).toEqual({ bottom: 112.3, maxH: 100.3 })
  })

  it('does not apply to a band-keep or mixed anchor, nor to collapsed boxes', () => {
    expect(
      floatKeepWholeOf(
        obstacleFactsOf({ textboxes: [box({ heightPx: 10, bandOverflow: true })] })!,
      ),
    ).toBeUndefined()
    expect(
      floatKeepWholeOf(
        obstacleFactsOf({ textboxes: [box({ heightPx: 10 }), box({ wrapSides: false })] })!,
      ),
    ).toBeUndefined()
    expect(floatKeepWholeOf(obstacleFactsOf({ textboxes: [box({})] }, COL, [0])!)).toBeUndefined()
  })
})

describe('rotation', () => {
  it('turned bounds keep the centre and swap the sides at a quarter turn', () => {
    expect(turnedBounds({ x: 0, y: 0, w: 200, h: 100 }, 90)).toEqual({
      x: 50,
      y: -50,
      w: 100,
      h: 200,
    })
    expect(turnedBounds({ x: 0, y: 0, w: 200, h: 100 }, 0)).toEqual({ x: 0, y: 0, w: 200, h: 100 })
  })

  it('layoutHeightOf inverts the rendered bounds of a turned box', () => {
    for (const rot of [0, 30, 90, 135, 270]) {
      const b = turnedBounds({ x: 0, y: 0, w: 200, h: 80 }, rot)
      expect(layoutHeightOf(b.w, b.h, 200, rot)).toBeCloseTo(80, 6)
    }
  })

  it('a rotated fixed box becomes its turned bounds as an obstacle', () => {
    const f = obstacleFactsOf({ textboxes: [box({ heightPx: 100, rotDeg: 90 })] })!
    expect(obstacleBounds(f.sides[0], 100)).toEqual({ x: 100, y: -38, w: 100, h: 200 })
    expect(floatKeepWholeOf(f)).toEqual({ bottom: 162, maxH: 200 })
  })
})
