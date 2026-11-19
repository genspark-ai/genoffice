/** Full-bleed tile group extraction (P22 B). */
import { describe, expect, it } from 'vitest'
import { extractFullBleedTiles } from '../src/analyze/shapes'
import type { Fill, PageShapes, Stroke } from '../src/ir'

const W = 612
const H = 792

const fill = (
  x0: number,
  y0: number,
  x1: number,
  y1: number,
  color: string,
  alpha?: number,
): Fill => ({ box: { x0, y0, x1, y1 }, color, ...(alpha !== undefined ? { alpha } : {}) }) as Fill

const edge = (x0: number, y0: number, x1: number, y1: number): Stroke => ({
  box: { x0, y0, x1, y1 },
  orientation: x1 - x0 >= y1 - y0 ? 'h' : 'v',
  widthPt: 0.2,
  color: '000000',
})

const shapesOf = (fills: Fill[], strokes: Stroke[] = []): PageShapes => ({
  fills,
  strokes,
  ignoredPaths: 0,
})

/** red top half, green bottom-left, blue bottom-right — edge to edge */
const mosaic = (): Fill[] => [
  fill(0, 425, W, H, 'FF0000'),
  fill(0, 0, 306, 428, '00FF00'),
  fill(306, 0, W, 425, '0000FF'),
]

describe('extractFullBleedTiles', () => {
  it('claims an edge-to-edge mosaic on a text-free page', () => {
    const shapes = shapesOf(mosaic(), [edge(0, 425, W, 425.3), edge(305.8, 0, 306.1, 428)])
    const tiles = extractFullBleedTiles(shapes, 0, W, H)
    expect(tiles).toHaveLength(3)
    expect(shapes.fills).toHaveLength(0)
    // tile edge strokes leave the pool — no ghost lattice grid
    expect(shapes.strokes).toHaveLength(0)
    // boxes clamp to the page
    for (const t of tiles) {
      expect(t.box.x0).toBeGreaterThanOrEqual(0)
      expect(t.box.y1).toBeLessThanOrEqual(H)
    }
  })

  it('does not fire on a text-bearing page (banner + table shading)', () => {
    const shapes = shapesOf(mosaic())
    expect(extractFullBleedTiles(shapes, 500, W, H)).toHaveLength(0)
    expect(shapes.fills).toHaveLength(3)
  })

  it('does not fire below the coverage floor (half-page banner)', () => {
    const shapes = shapesOf([
      fill(0, H / 2, W, H, 'FF0000'),
      fill(0, H * 0.4, W, H * 0.5, '00FF00'),
    ])
    expect(extractFullBleedTiles(shapes, 0, W, H)).toHaveLength(0)
  })

  it('ignores translucent washes and single fills', () => {
    const translucent = mosaic().map((f) => ({ ...f, alpha: 120 }))
    expect(extractFullBleedTiles(shapesOf(translucent), 0, W, H)).toHaveLength(0)
    expect(extractFullBleedTiles(shapesOf([fill(0, 0, W, H, 'FF0000')]), 0, W, H)).toHaveLength(0)
  })

  it('keeps a stroke spanning multiple tiles (not a single tile edge)', () => {
    // crosses from the green tile into the red one — fits no single tile box
    const spanning = edge(100, 100, 100.4, 700)
    const shapes = shapesOf(mosaic(), [spanning])
    extractFullBleedTiles(shapes, 0, W, H)
    expect(shapes.strokes).toHaveLength(1)
  })
})
