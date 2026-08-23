/** Shape normalization unit tests (pure geometry, no wasm). */
import { describe, expect, it } from 'vitest'
import {
  extractBackgroundPanels,
  extractPageBackground,
  extractTextBackdrops,
  isNearWhite,
  normalizeShapes,
  rectOfSubpath,
} from '../src/analyze'
import type { RawPath, RawSubpath } from '../src/ir'

const sub = (points: Array<[number, number]>, closed = true, hasCurves = false): RawSubpath => ({
  points: points.map(([x, y]) => ({ x, y })),
  closed,
  hasCurves,
})

const path = (subpaths: RawSubpath[], over: Partial<RawPath> = {}): RawPath => ({
  subpaths,
  filled: false,
  stroked: false,
  fillColor: '000000',
  strokeColor: '000000',
  strokeWidth: 1,
  ...over,
})

const rect = (x0: number, y0: number, x1: number, y1: number): RawSubpath =>
  sub([
    [x0, y0],
    [x1, y0],
    [x1, y1],
    [x0, y1],
  ])

describe('rectOfSubpath', () => {
  it('detects an axis-aligned rectangle (4 points, either winding)', () => {
    expect(rectOfSubpath(rect(10, 20, 110, 40))).toEqual({ x0: 10, y0: 20, x1: 110, y1: 40 })
  })

  it('accepts a closing point repeating the first', () => {
    const r = sub(
      [
        [10, 20],
        [110, 20],
        [110, 40],
        [10, 40],
        [10, 20],
      ],
      false,
    )
    expect(rectOfSubpath(r)).toEqual({ x0: 10, y0: 20, x1: 110, y1: 40 })
  })

  it('accepts a rectangle whose outline starts mid-edge (P16 A)', () => {
    // drawn from a point on the top edge: 5 corners + closing repeat
    const r = sub(
      [
        [298, 278],
        [54, 278],
        [54, 414],
        [542, 414],
        [542, 278],
        [298, 278],
      ],
      true,
    )
    expect(rectOfSubpath(r)).toEqual({ x0: 54, y0: 278, x1: 542, y1: 414 })
  })

  it('accepts doubled/duplicate points on the outline (P16 A)', () => {
    const r = sub(
      [
        [0, 0],
        [595, 0],
        [595, 880],
        [0, 880],
        [0, 0],
        [0, 0],
      ],
      true,
    )
    expect(rectOfSubpath(r)).toEqual({ x0: 0, y0: 0, x1: 595, y1: 880 })
  })

  it('rejects diagonals and non-quads', () => {
    expect(
      rectOfSubpath(
        sub([
          [0, 0],
          [50, 50],
          [100, 0],
          [50, -50],
        ]),
      ),
    ).toBeNull()
    expect(
      rectOfSubpath(
        sub([
          [0, 0],
          [10, 0],
          [10, 10],
        ]),
      ),
    ).toBeNull()
  })
})

describe('normalizeShapes', () => {
  it('keeps the source alpha on translucent fills, omits it on opaque ones (P11 A)', () => {
    const shapes = normalizeShapes([
      path([rect(0, 0, 400, 540)], { filled: true, fillColor: '0D0D0D', fillAlpha: 178 }),
      path([rect(0, 0, 400, 100)], { filled: true, fillColor: '112233' }),
    ])
    expect(shapes.fills).toHaveLength(2)
    expect(shapes.fills[0]!.alpha).toBe(178)
    expect(shapes.fills[1]!.alpha).toBeUndefined()
  })

  it('turns a thin filled rectangle into a horizontal stroke (table line / underline)', () => {
    const shapes = normalizeShapes([
      path([rect(72, 699, 472, 701)], { filled: true, fillColor: '333333' }),
    ])
    expect(shapes.fills).toHaveLength(0)
    expect(shapes.strokes).toHaveLength(1)
    const s = shapes.strokes[0]!
    expect(s.orientation).toBe('h')
    expect(s.widthPt).toBeCloseTo(2, 5)
    expect(s.color).toBe('333333')
    expect(s.box.x0).toBeCloseTo(72)
    expect(s.box.x1).toBeCloseTo(472)
  })

  it('turns a thin vertical rectangle into a vertical stroke', () => {
    const shapes = normalizeShapes([path([rect(199, 600, 201, 700)], { filled: true })])
    expect(shapes.strokes).toHaveLength(1)
    expect(shapes.strokes[0]!.orientation).toBe('v')
  })

  it('keeps a plain filled rectangle as a Fill (cell shading candidate)', () => {
    const shapes = normalizeShapes([
      path([rect(72, 660, 200, 700)], { filled: true, fillColor: 'FFCC00' }),
    ])
    expect(shapes.strokes).toHaveLength(0)
    expect(shapes.fills).toEqual([{ box: { x0: 72, y0: 660, x1: 200, y1: 700 }, color: 'FFCC00' }])
  })

  it('emits four border strokes for a stroked (outlined) rectangle', () => {
    const shapes = normalizeShapes([path([rect(72, 660, 200, 700)], { stroked: true })])
    expect(shapes.fills).toHaveLength(0)
    expect(shapes.strokes).toHaveLength(4)
    expect(shapes.strokes.filter((s) => s.orientation === 'h')).toHaveLength(2)
    expect(shapes.strokes.filter((s) => s.orientation === 'v')).toHaveLength(2)
  })

  it('emits both the Fill and its border strokes for a filled+stroked rect', () => {
    const shapes = normalizeShapes([
      path([rect(72, 660, 200, 700)], { filled: true, stroked: true }),
    ])
    expect(shapes.fills).toHaveLength(1)
    expect(shapes.strokes).toHaveLength(4)
  })

  it('turns stroked h/v segments into strokes and counts diagonals as ignored', () => {
    const zigzag = sub(
      [
        [0, 100],
        [50, 100], // h
        [50, 40], // v
        [90, 10], // diagonal
      ],
      false,
    )
    const shapes = normalizeShapes([path([zigzag], { stroked: true, strokeWidth: 2 })])
    expect(shapes.strokes).toHaveLength(2)
    expect(shapes.strokes[0]!.widthPt).toBe(2)
    expect(shapes.ignoredPaths).toBe(1)
  })

  it('ignores bezier subpaths and non-rect fills (vector art is P4)', () => {
    const curved = sub(
      [
        [0, 0],
        [10, 20],
        [20, 20],
        [30, 0],
      ],
      false,
      true,
    )
    const triangle = sub([
      [0, 0],
      [40, 0],
      [20, 30],
    ])
    const shapes = normalizeShapes([path([curved, triangle], { filled: true })])
    expect(shapes.strokes).toHaveLength(0)
    expect(shapes.fills).toHaveLength(0)
    expect(shapes.ignoredPaths).toBe(2)
  })

  it('skips invisible (clip-only) paths entirely', () => {
    const shapes = normalizeShapes([path([rect(0, 0, 100, 100)])])
    expect(shapes.strokes).toHaveLength(0)
    expect(shapes.fills).toHaveLength(0)
    expect(shapes.ignoredPaths).toBe(0)
  })

  it('salvages a thin axis-aligned filled bar with redundant points as a stroke (P7)', () => {
    // decorative title rule drawn as a 6-point closed polygon (collinear midpoints)
    const bar = sub([
      [54, 792.9],
      [120, 792.9],
      [189, 792.9],
      [189, 794.4],
      [120, 794.4],
      [54, 794.4],
    ])
    const shapes = normalizeShapes([path([bar], { filled: true, fillColor: '0A3C61' })])
    expect(shapes.ignoredPaths).toBe(0)
    expect(shapes.strokes).toHaveLength(1)
    const stroke = shapes.strokes[0]!
    expect(stroke.orientation).toBe('h')
    expect(stroke.color).toBe('0A3C61')
    expect(stroke.widthPt).toBeCloseTo(1.5, 1)
    expect(stroke.box.x0).toBeCloseTo(54)
    expect(stroke.box.x1).toBeCloseTo(189)
  })

  it('a thick rect with a redundant collinear corner is a fill; diagonals stay ignored', () => {
    // 5-point outline whose extra point sits mid-edge — a rectangle (P16 A)
    const thick = sub([
      [10, 10],
      [60, 10],
      [110, 10],
      [110, 60],
      [10, 60],
    ])
    const diag = sub([
      [10, 10],
      [110, 12.5],
      [110, 14],
      [10, 11.5],
    ])
    const shapes = normalizeShapes([path([thick, diag], { filled: true })])
    expect(shapes.strokes).toHaveLength(0)
    expect(shapes.fills).toHaveLength(1)
    expect(shapes.fills[0]!.box).toEqual({ x0: 10, y0: 10, x1: 110, y1: 60 })
    expect(shapes.ignoredPaths).toBe(1)
  })
})

describe('extractPageBackground', () => {
  const W = 612
  const H = 792

  it('pulls a full-page fill out of the pool and returns its color', () => {
    const shapes = normalizeShapes([
      path([rect(0, 0, W, H)], { filled: true, fillColor: 'F0E8D8' }),
      path([rect(50, 50, 150, 80)], { filled: true, fillColor: 'FFCC00' }),
    ])
    const bg = extractPageBackground(shapes.fills, W, H)
    expect(bg).toBe('F0E8D8')
    expect(shapes.fills).toHaveLength(1)
    expect(shapes.fills[0]!.color).toBe('FFCC00')
  })

  it('keeps the topmost color when backgrounds stack, removing all of them', () => {
    const shapes = normalizeShapes([
      path([rect(0, 0, W, H)], { filled: true, fillColor: '111111' }),
      path([rect(-2, -2, W + 2, H + 2)], { filled: true, fillColor: '2A9D8F' }),
    ])
    const bg = extractPageBackground(shapes.fills, W, H)
    expect(bg).toBe('2A9D8F')
    expect(shapes.fills).toHaveLength(0)
  })

  it('ignores fills that cover most width but not height (banners)', () => {
    const shapes = normalizeShapes([
      path([rect(0, 0, W, 120)], { filled: true, fillColor: '333333' }),
    ])
    const bg = extractPageBackground(shapes.fills, W, H)
    expect(bg).toBeUndefined()
    expect(shapes.fills).toHaveLength(1)
  })
})

describe('extractBackgroundPanels (P10 B)', () => {
  const W = 960
  const H = 540
  // a column of characters sitting on the left-spine panel
  const spineChars = Array.from({ length: 6 }, (_, i) => ({
    x0: 60,
    y0: 100 + i * 40,
    x1: 80,
    y1: 120 + i * 40,
  }))

  it('lifts a full-height edge-flush spine panel with text on it', () => {
    const shapes = normalizeShapes([
      path([rect(0, 0, 220, H)], { filled: true, fillColor: '2A0A0F' }),
      path([rect(400, 200, 560, 280)], { filled: true, fillColor: 'FFCC00' }),
    ])
    const panels = extractBackgroundPanels(shapes.fills, spineChars, W, H)
    expect(panels).toHaveLength(1)
    expect(panels[0]!.color).toBe('2A0A0F')
    expect(shapes.fills).toHaveLength(1) // the mid-page box stays a fill
  })

  it('rejects text-less panels, mid-page slabs, and small shading bars', () => {
    const midPage = normalizeShapes([
      // large but not flush with any edge (a content card, e.g. a checklist panel)
      path([rect(40, 60, 900, 480)], { filled: true, fillColor: 'BDD7EE' }),
    ])
    expect(extractBackgroundPanels(midPage.fills, spineChars, W, H)).toHaveLength(0)
    expect(midPage.fills).toHaveLength(1)

    const noText = normalizeShapes([
      path([rect(0, 0, 220, H)], { filled: true, fillColor: '2A0A0F' }),
    ])
    expect(extractBackgroundPanels(noText.fills, [], W, H)).toHaveLength(0)

    const bar = normalizeShapes([
      // full-width but tiny area (a highlight band)
      path([rect(0, 500, W, 530)], { filled: true, fillColor: 'FFF2CC' }),
    ])
    expect(extractBackgroundPanels(bar.fills, spineChars, W, H)).toHaveLength(0)
  })

  it('only takes panels from the bottom of the z-order', () => {
    const shapes = normalizeShapes([
      path([rect(300, 100, 800, 400)], { filled: true, fillColor: 'FFCC00' }),
      path([rect(0, 0, 220, H)], { filled: true, fillColor: '2A0A0F' }),
    ])
    // first fill is not a panel → the scan stops, the spine behind it stays
    expect(extractBackgroundPanels(shapes.fills, spineChars, W, H)).toHaveLength(0)
    expect(shapes.fills).toHaveLength(2)
  })
})

describe('isNearWhite', () => {
  it('treats pure and near white as white', () => {
    expect(isNearWhite('FFFFFF')).toBe(true)
    expect(isNearWhite('F8F8F8')).toBe(true)
  })
  it('keeps real colors', () => {
    expect(isNearWhite('F0E8D8')).toBe(false)
    expect(isNearWhite('000000')).toBe(false)
  })
})

describe('translucent fills (P10 C)', () => {
  it('drops low-alpha fills — glows/shadows never become shading or highlight', () => {
    const glow = path([rect(369, 189, 591, 411)], { filled: true, fillColor: '000000' })
    ;(glow as { fillAlpha?: number }).fillAlpha = 89
    const shapes = normalizeShapes([glow])
    expect(shapes.fills).toHaveLength(0)
    expect(shapes.ignoredPaths).toBe(1)
  })

  it('keeps opaque fills (explicit alpha 255 and legacy paths without alpha)', () => {
    const opaque = path([rect(100, 100, 200, 140)], { filled: true, fillColor: 'FFCC00' })
    ;(opaque as { fillAlpha?: number }).fillAlpha = 255
    const legacy = path([rect(300, 100, 400, 140)], { filled: true, fillColor: '00CCFF' })
    const shapes = normalizeShapes([opaque, legacy])
    expect(shapes.fills).toHaveLength(2)
  })
})

describe('extractTextBackdrops (P10 B)', () => {
  const W = 960
  const H = 540
  const tileChars = (x0: number, y0: number, color: string, n = 6) =>
    Array.from({ length: n }, (_, i) => ({
      box: { x0: x0 + 20 + i * 14, y0: y0 + 60, x1: x0 + 32 + i * 14, y1: y0 + 74 },
      code: 0x4e2d,
      color,
    }))

  it('lifts a dark card fill under near-white text', () => {
    const shapes = normalizeShapes([
      path([rect(136, 255, 466, 426)], { filled: true, fillColor: '00A0EB' }),
    ])
    const out = extractTextBackdrops(shapes.fills, tileChars(136, 255, 'FFFFFF'), [], W, H)
    expect(out).toHaveLength(1)
    expect(out[0]!.color).toBe('00A0EB')
    expect(shapes.fills).toHaveLength(0)
  })

  it('keeps line-hugging bars, low-contrast fills, and lattice cell shading', () => {
    // a text-height highlight bar with dark text: not a plate (height < 2.5×)
    const bar = normalizeShapes([
      path([rect(136, 310, 306, 340)], { filled: true, fillColor: '00A0EB' }),
    ])
    expect(extractTextBackdrops(bar.fills, tileChars(136, 255, '1F1613'), [], W, H)).toHaveLength(0)
    expect(bar.fills).toHaveLength(1)

    // light fill under light text: no contrast either way
    const lightFill = normalizeShapes([
      path([rect(136, 255, 466, 426)], { filled: true, fillColor: 'EFEFEF' }),
    ])
    expect(
      extractTextBackdrops(lightFill.fills, tileChars(136, 255, 'FFFFFF'), [], W, H),
    ).toHaveLength(0)

    const shaded = normalizeShapes([
      path([rect(136, 255, 466, 426)], { filled: true, fillColor: '00A0EB' }),
    ])
    const grid = [{ x0: 100, y0: 200, x1: 500, y1: 460 }]
    expect(
      extractTextBackdrops(shaded.fills, tileChars(136, 255, 'FFFFFF'), grid, W, H),
    ).toHaveLength(0)
  })

  it('lifts a mid-size label plate far taller than its dark label (P16 I)', () => {
    // a grey caption-style plate: ~4% of the page, text ≪ plate height
    const plate = normalizeShapes([
      path([rect(136, 255, 306, 380)], { filled: true, fillColor: 'D4D4D6' }),
    ])
    const out = extractTextBackdrops(plate.fills, tileChars(136, 255, '404040', 4), [], W, H)
    expect(out).toHaveLength(1)
    expect(out[0]!.color).toBe('D4D4D6')
    expect(plate.fills).toHaveLength(0)
  })

  it('lifts a page-scale divider panel with dark text (contrast rule)', () => {
    const divider = normalizeShapes([
      path([rect(20, 250, 540, 520)], { filled: true, fillColor: 'D9D9D9' }),
    ])
    const out = extractTextBackdrops(divider.fills, tileChars(20, 300, '404040'), [], W, H)
    expect(out).toHaveLength(1)
    expect(out[0]!.color).toBe('D9D9D9')
  })

  describe('small capsule backdrops (P12 C)', () => {
    // ~IPD cover pill: white rounded capsule (646,190)-(873,230) with six
    // dark 15pt chars on it, page carries a photo/wash backdrop
    const capsule = { box: { x0: 646, y0: 190, x1: 873, y1: 230 }, color: 'FFFFFF' }
    const pillChars = Array.from({ length: 6 }, (_, i) => ({
      box: { x0: 712 + i * 16, y0: 202, x1: 727 + i * 16, y1: 218 },
      code: 0x4e3b,
      color: '333333',
    }))

    it('lifts a white pill with contrasting text when the page has a backdrop', () => {
      const out = extractTextBackdrops([], pillChars, [], W, H, [capsule], true)
      expect(out).toHaveLength(1)
      expect(out[0]!.color).toBe('FFFFFF')
    })

    it('skips a white pill on a plain white page (nothing to contrast with)', () => {
      expect(extractTextBackdrops([], pillChars, [], W, H, [capsule], false)).toHaveLength(0)
    })

    it('skips text-less rounded decor and low-contrast pills', () => {
      expect(extractTextBackdrops([], [], [], W, H, [capsule], true)).toHaveLength(0)
      const grayText = pillChars.map((c) => ({ ...c, color: 'E8E8E8' }))
      expect(extractTextBackdrops([], grayText, [], W, H, [capsule], true)).toHaveLength(0)
    })

    it('skips a wide swoosh whose text covers only a sliver of it', () => {
      const swoosh = { box: { x0: 0, y0: 190, x1: 960, y1: 240 }, color: '1F4E79' }
      const strayLabel = pillChars.slice(0, 2)
      expect(extractTextBackdrops([], strayLabel, [], W, H, [swoosh], true)).toHaveLength(0)
    })

    it('leaves capsules inside a table grid to cell shading', () => {
      const grid = [{ x0: 600, y0: 150, x1: 900, y1: 260 }]
      expect(extractTextBackdrops([], pillChars, grid, W, H, [capsule], true)).toHaveLength(0)
    })
  })
})

describe('curved-fill coverage gate (P29 C)', () => {
  const curvedPath = (points: Array<[number, number]>): RawPath => ({
    subpaths: [sub(points, true, true)],
    filled: true,
    stroked: false,
    fillColor: '10c4aa',
    strokeColor: '000000',
    strokeWidth: 0,
  })

  it('keeps a rounded-rect-like curved fill (polygon fills its bbox)', () => {
    const shapes = normalizeShapes([
      curvedPath([
        [0, 0],
        [200, 0],
        [200, 80],
        [0, 80],
      ]),
    ])
    expect(shapes.curvedFills).toHaveLength(1)
  })

  it('drops logo art whose polygon covers little of its bbox', () => {
    // thin L-shaped swoosh: area 1900 over a 100×100 bbox
    const shapes = normalizeShapes([
      curvedPath([
        [0, 0],
        [100, 0],
        [100, 10],
        [10, 10],
        [10, 100],
        [0, 100],
      ]),
    ])
    expect(shapes.curvedFills ?? []).toHaveLength(0)
  })
})

describe('curved-art z group (P29 C)', () => {
  const W = 960
  const H = 540
  const chars = (x0: number, y0: number) =>
    Array.from({ length: 8 }, (_, i) => ({
      box: { x0: x0 + 20 + i * 14, y0: y0 + 60, x1: x0 + 32 + i * 14, y1: y0 + 74 },
      code: 0x4e2d,
      color: 'FFFFFF',
    }))
  const curved = (x0: number, y0: number, x1: number, y1: number, z: number) => ({
    box: { x0, y0, x1, y1 },
    color: '10C4AA',
    z,
  })

  it('rejects mutually-overlapping curved fills at one z (watermark art)', () => {
    const art = [
      curved(100, 100, 400, 400, 0),
      curved(150, 150, 450, 450, 0),
      curved(120, 200, 380, 500, 0),
    ]
    const out = extractTextBackdrops([], chars(100, 100), [], W, H, art)
    expect(out).toHaveLength(0)
  })

  it('keeps disjoint rounded cards that merely share a form z', () => {
    const cards = [
      curved(40, 100, 300, 300, 0),
      curved(340, 100, 600, 300, 0),
      curved(640, 100, 900, 300, 0),
    ]
    const out = extractTextBackdrops([], chars(40, 100), [], W, H, cards)
    expect(out.length).toBeGreaterThan(0)
  })
})
