/**
 * Unit tests for pure geometry align/distribute functions (alignRects / distributeRects).
 * No IO, no model dependency; verifies correctness of 6 alignments + 2 distributions.
 */
import { describe, it, expect } from 'vitest'
import { alignRects, distributeRects } from '../src/align'

// ── Helpers ───────────────────────────────────────────────────────────────────

const R = (x: number, y: number, w: number, h: number) => ({ x, y, w, h })

// ── Alignment tests ───────────────────────────────────────────────────────────

describe('alignRects', () => {
  describe('multi-select (no container — aligns to bounding box)', () => {
    const rects = [R(100, 200, 50, 30), R(300, 100, 80, 60)]

    it('left: all elements x = bounding box left edge', () => {
      const result = alignRects(rects, 'left', null)
      expect(result[0]!.x).toBe(100)
      expect(result[1]!.x).toBe(100)
      // y unchanged
      expect(result[0]!.y).toBe(200)
      expect(result[1]!.y).toBe(100)
    })

    it('right: all elements right edge = bounding box right edge', () => {
      const result = alignRects(rects, 'right', null)
      // bbox right = 300+80 = 380; el0 right = x+50 → x = 380-50 = 330
      expect(result[0]!.x).toBe(330)
      // el1 right = x+80 → x = 380-80 = 300 (unchanged)
      expect(result[1]!.x).toBe(300)
    })

    it('center-h: all elements centered horizontally on bounding box', () => {
      const result = alignRects(rects, 'center-h', null)
      // bbox x=100, w=280 → center x=100+280/2=240; el0: 240-50/2=215
      expect(result[0]!.x).toBe(215)
      // el1: 240-80/2=200
      expect(result[1]!.x).toBe(200)
    })

    it('top: all elements y = bounding box top edge', () => {
      const result = alignRects(rects, 'top', null)
      expect(result[0]!.y).toBe(100)
      expect(result[1]!.y).toBe(100)
    })

    it('bottom: all elements bottom edge = bounding box bottom edge', () => {
      const result = alignRects(rects, 'bottom', null)
      // bbox bottom = max(200+30, 100+60) = 230; el0: 230-30=200; el1: 230-60=170
      expect(result[0]!.y).toBe(200)
      expect(result[1]!.y).toBe(170)
    })

    it('center-v: all elements vertically centered on bounding box', () => {
      const result = alignRects(rects, 'center-v', null)
      // bbox y=100, h=130 → center y=100+130/2=165; el0: 165-30/2=150; el1: 165-60/2=135
      expect(result[0]!.y).toBe(150)
      expect(result[1]!.y).toBe(135)
    })
  })

  describe('single-select (container = slide page rect)', () => {
    const slide = R(0, 0, 1280, 720)
    const rects = [R(100, 100, 200, 100)]

    it('left aligns to slide left', () => {
      const result = alignRects(rects, 'left', slide)
      expect(result[0]!.x).toBe(0)
    })

    it('right aligns to slide right edge', () => {
      const result = alignRects(rects, 'right', slide)
      expect(result[0]!.x).toBe(1280 - 200)
    })

    it('center-h centers on slide', () => {
      const result = alignRects(rects, 'center-h', slide)
      expect(result[0]!.x).toBe((1280 - 200) / 2)
    })

    it('top aligns to slide top', () => {
      const result = alignRects(rects, 'top', slide)
      expect(result[0]!.y).toBe(0)
    })

    it('bottom aligns to slide bottom edge', () => {
      const result = alignRects(rects, 'bottom', slide)
      expect(result[0]!.y).toBe(720 - 100)
    })

    it('center-v centers on slide vertically', () => {
      const result = alignRects(rects, 'center-v', slide)
      expect(result[0]!.y).toBe((720 - 100) / 2)
    })
  })
})

// ── Distribution tests ────────────────────────────────────────────────────────

describe('distributeRects', () => {
  it('horizontal: 3 equal-width rects distribute evenly', () => {
    // a: x=0,w=50  b: x=200,w=50  c: x=300,w=50 → span=350, gap=(350-150)/2=100
    // sorted: a(0), b(200), c(300)
    // first=a at 0, cursor: a→50, gap=100, b→150, c→300
    const rects = [R(0, 0, 50, 20), R(200, 0, 50, 20), R(300, 0, 50, 20)]
    const result = distributeRects(rects, 'horizontal')
    const xs = result.map((r) => r.x).sort((a, b) => a - b)
    expect(xs[0]).toBe(0)
    expect(xs[1]).toBe(150)
    expect(xs[2]).toBe(300)
  })

  it('vertical: 3 equal-height rects distribute evenly', () => {
    const rects = [R(0, 0, 20, 50), R(0, 200, 20, 50), R(0, 300, 20, 50)]
    const result = distributeRects(rects, 'vertical')
    const ys = result.map((r) => r.y).sort((a, b) => a - b)
    expect(ys[0]).toBe(0)
    expect(ys[1]).toBe(150)
    expect(ys[2]).toBe(300)
  })

  it('with < 3 elements returns original positions unchanged', () => {
    const rects = [R(10, 20, 50, 30), R(100, 200, 50, 30)]
    const result = distributeRects(rects, 'horizontal')
    expect(result[0]).toEqual({ x: 10, y: 20 })
    expect(result[1]).toEqual({ x: 100, y: 200 })
  })

  it('horizontal: 4 variable-width rects distribute evenly', () => {
    // a: x=0,w=30  b: x=100,w=40  c: x=200,w=20  d: x=300,w=50
    // sorted by x: a,b,c,d
    // span = 350; totalWidth = 140; gap = (350-140)/3 = 70
    // positions: a=0, b=0+30+70=100, c=100+40+70=210, d=210+20+70=300
    const rects = [R(0, 0, 30, 20), R(100, 0, 40, 20), R(200, 0, 20, 20), R(300, 0, 50, 20)]
    const result = distributeRects(rects, 'horizontal')
    const xs = result.map((r) => r.x)
    expect(xs[0]).toBe(0)   // a stays
    expect(xs[1]).toBeCloseTo(100)  // b
    expect(xs[2]).toBeCloseTo(210)  // c
    expect(xs[3]).toBe(300) // d stays
  })
})
