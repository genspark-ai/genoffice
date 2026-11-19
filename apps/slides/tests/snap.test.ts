import { describe, it, expect } from 'vitest'
import { computeSnap, computeSpacingSnap, type SnapTarget } from '../src/renderer/snap'

const pageCenter: SnapTarget = { v: [640], h: [360] } // 1280x720 center

describe('alignment snapping computeSnap', () => {
  it('left edge snaps to page-left when within threshold', () => {
    const r = computeSnap({ x: 3, y: 100, w: 200, h: 100 }, [{ v: [0], h: [] }], 6)
    expect(r.x).toBe(0)
    expect(r.guides).toContainEqual({ axis: 'v', pos: 0 })
  })

  it('center snaps to page center line', () => {
    // box width 200, center should reach 640 -> x should be 540
    const r = computeSnap({ x: 538, y: 0, w: 200, h: 100 }, [pageCenter], 6)
    expect(r.x).toBe(540)
    expect(r.guides).toContainEqual({ axis: 'v', pos: 640 })
  })

  it('vertical middle snaps to page middle', () => {
    const r = computeSnap({ x: 0, y: 308, w: 100, h: 100 }, [pageCenter], 6)
    // middle = y+50 should reach 360 -> y=310
    expect(r.y).toBe(310)
    expect(r.guides).toContainEqual({ axis: 'h', pos: 360 })
  })

  it('no snap when outside threshold', () => {
    const r = computeSnap({ x: 100, y: 100, w: 50, h: 50 }, [{ v: [0], h: [0] }], 6)
    expect(r.x).toBeNull()
    expect(r.y).toBeNull()
    expect(r.guides).toEqual([])
  })

  it('snaps to another element edge (right→left alignment)', () => {
    // Another element's left edge at x=500; dragged box's right edge (x+w=200+298=498) is 2 from 500 -> snap right to 500 -> x=300
    const r = computeSnap({ x: 298, y: 0, w: 200, h: 100 }, [{ v: [500], h: [] }], 6)
    expect(r.x).toBe(300)
  })

  it('picks nearest line among multiple', () => {
    const r = computeSnap({ x: 100, y: 0, w: 100, h: 50 }, [{ v: [98, 250], h: [] }], 6)
    expect(r.x).toBe(98) // left(100)→98 dist 2
  })
})

describe('equal-spacing snapping computeSpacingSnap', () => {
  // Three elements in a row: A[0..100] B[200..300], dragged box width 100
  const A = { x: 0, y: 100, w: 100, h: 50 }
  const B = { x: 200, y: 100, w: 100, h: 50 }

  it('centering: snaps when landing between two neighbors with equal gaps on both sides', () => {
    // A's right edge 100, B's left edge 200, box width 60 -> equal-gap position x = (100+200-60)/2 = 120
    const r = computeSpacingSnap({ x: 117, y: 110, w: 60, h: 30 }, [A, B], 6)
    expect(r.x).toBe(120)
    expect(r.indicators.filter((i) => i.axis === 'x')).toHaveLength(2)
    // The two gaps are equal (100->120 and 180->200, 20 each)
    const [g1, g2] = r.indicators.filter((i) => i.axis === 'x')
    expect(g1!.to - g1!.from).toBeCloseTo(g2!.to - g2!.from)
  })

  it('continuation: two elements on the left with gap g=100, dragged box snaps to g right of the nearest neighbor', () => {
    // A[0..100] B[200..300] gap 100 -> dragged box should snap to x=400
    const r = computeSpacingSnap({ x: 396, y: 110, w: 80, h: 30 }, [A, B], 6)
    expect(r.x).toBe(400)
  })

  it('continuation: two elements on the right with gap g, dragged box snaps to their left', () => {
    // B[200..300] C[400..500] gap 100 -> box width 50 snaps to x=150? gap=100 -> target = 200-100-50 = 50
    const C = { x: 400, y: 100, w: 100, h: 50 }
    const r = computeSpacingSnap({ x: 54, y: 110, w: 50, h: 30 }, [B, C], 6)
    expect(r.x).toBe(50)
  })

  it('elements in a different band (no projection overlap) do not participate', () => {
    const far = { x: 200, y: 500, w: 100, h: 50 } // y band doesn't overlap at all
    const r = computeSpacingSnap({ x: 117, y: 110, w: 60, h: 30 }, [A, far], 6)
    expect(r.x).toBeNull()
  })

  it('does not snap beyond the threshold', () => {
    const r = computeSpacingSnap({ x: 108, y: 110, w: 60, h: 30 }, [A, B], 6)
    expect(r.x).toBeNull()
  })

  it('vertical equal spacing (y axis)', () => {
    const T = { x: 100, y: 0, w: 100, h: 50 }
    const Bt = { x: 100, y: 200, w: 100, h: 50 }
    // Top edge 50, bottom edge 200, box height 50 -> y = (50+200-50)/2 = 100
    const r = computeSpacingSnap({ x: 110, y: 97, w: 80, h: 50 }, [T, Bt], 6)
    expect(r.y).toBe(100)
    expect(r.indicators.filter((i) => i.axis === 'y')).toHaveLength(2)
  })

  it('overlapping elements (negative gap) produce no snapping', () => {
    // Neighbor overlaps the dragged box: A's right edge 100 > box left 90
    const r = computeSpacingSnap({ x: 90, y: 110, w: 60, h: 30 }, [A, B], 6)
    // Centering candidate target=(100+200-60)/2=120 is beyond the threshold from 90 -> null
    expect(r.x).toBeNull()
  })
})
