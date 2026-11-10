/** Empty stroke-frame extraction (P16 K): hand-built strokes, no wasm. */
import { describe, expect, it } from 'vitest'
import { extractEmptyFrames, frameImage } from '../src/analyze'
import type { Stroke } from '../src/ir'

const h = (x0: number, x1: number, y: number, w = 1): Stroke => ({
  box: { x0, x1, y0: y - w / 2, y1: y + w / 2 },
  orientation: 'h',
  widthPt: w,
  color: '000000',
})
const v = (y0: number, y1: number, x: number, w = 1): Stroke => ({
  box: { x0: x - w / 2, x1: x + w / 2, y0, y1 },
  orientation: 'v',
  widthPt: w,
  color: '000000',
})

/** the europa quote box: (71,329)-(522,367), no text inside */
const frameStrokes = (): Stroke[] => [
  h(71, 522, 367),
  h(71, 522, 329),
  v(329, 367, 71),
  v(329, 367, 522),
]

const charAt = (x: number, y: number) => ({
  box: { x0: x, y0: y, x1: x + 10, y1: y + 12 },
  code: 65,
})

describe('extractEmptyFrames (P16 K)', () => {
  it('lifts an empty 4-stroke rectangle out of the pool as a frame', () => {
    const strokes = [...frameStrokes(), h(72, 522, 800)] // plus an unrelated rule
    const frames = extractEmptyFrames(strokes, [], [])
    expect(frames).toHaveLength(1)
    expect(frames[0]!.box.x0).toBeCloseTo(70.5, 0)
    expect(frames[0]!.box.y1).toBeCloseTo(367.5, 0)
    // the frame's strokes leave the pool; the rule stays
    expect(strokes).toHaveLength(1)
    expect(strokes[0]!.box.y0).toBeCloseTo(799.5)
  })

  it('a frame with text inside stays in the pool (lattice/decor territory)', () => {
    const strokes = frameStrokes()
    const frames = extractEmptyFrames(strokes, [charAt(200, 340)], [])
    expect(frames).toHaveLength(0)
    expect(strokes).toHaveLength(4)
  })

  it('a frame inside a solved grid stays (its lines belong to the table)', () => {
    const strokes = frameStrokes()
    const frames = extractEmptyFrames(strokes, [], [{ x0: 0, y0: 300, x1: 600, y1: 400 }])
    expect(frames).toHaveLength(0)
    expect(strokes).toHaveLength(4)
  })

  it('incomplete rectangles (three sides) never become frames', () => {
    const strokes = [h(71, 522, 367), h(71, 522, 329), v(329, 367, 71)]
    expect(extractEmptyFrames(strokes, [], [])).toHaveLength(0)
    expect(strokes).toHaveLength(3)
  })
})

describe('frameImage bitmap budget', () => {
  it('renders normal frames at full 2x scale', () => {
    const img = frameImage({ x0: 71, y0: 329, x1: 522, y1: 367 }, '000000', 1)
    expect(img.pixelWidth).toBe(Math.round((522 - 71) * 2))
    expect(img.pixelHeight).toBe(Math.round((367 - 329) * 2))
  })

  it('caps giant frame bitmaps instead of allocating hundreds of MB', () => {
    const img = frameImage({ x0: 0, y0: 0, x1: 5000, y1: 5000 }, '000000', 1)
    expect(img.pixelWidth).toBeLessThanOrEqual(2048)
    expect(img.pixelHeight).toBeLessThanOrEqual(2048)
    expect(img.pixelWidth).toBe(img.pixelHeight)
  })

  it('degrades non-finite boxes to the minimum bitmap instead of throwing', () => {
    const img = frameImage({ x0: 0, y0: 0, x1: NaN, y1: 100 }, '000000', 1)
    expect(img.pixelWidth).toBe(2)
    expect(img.pixelHeight).toBe(200)
  })
})
