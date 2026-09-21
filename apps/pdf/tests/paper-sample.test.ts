import { describe, expect, it } from 'vitest'
import { ringMedianColor } from '../src/renderer/paper-sample'

/** w×h RGBA bitmap filled with `bg`, with `ink` painted inside `rect` */
function bitmap(
  w: number,
  h: number,
  bg: [number, number, number],
  rect: { x: number; y: number; w: number; h: number },
  ink: [number, number, number],
): Uint8ClampedArray {
  const data = new Uint8ClampedArray(w * h * 4)
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      const inside = x >= rect.x && x < rect.x + rect.w && y >= rect.y && y < rect.y + rect.h
      const c = inside ? ink : bg
      data.set([c[0], c[1], c[2], 255], (y * w + x) * 4)
    }
  }
  return data
}

describe('ringMedianColor', () => {
  const rect = { x: 10, y: 10, w: 40, h: 12 }

  it('reads the page color around the run, not the run ink', () => {
    const data = bitmap(80, 40, [11, 22, 60], rect, [255, 255, 255])
    expect(ringMedianColor(data, 80, 40, rect, 3)).toEqual([11, 22, 60])
  })

  it('survives glyph ink poking a pixel past the box (median, not mean)', () => {
    const data = bitmap(
      80,
      40,
      [11, 22, 60],
      { ...rect, y: rect.y - 1, h: rect.h + 2 },
      [255, 255, 255],
    )
    expect(ringMedianColor(data, 80, 40, rect, 3)).toEqual([11, 22, 60])
  })

  it('clamps the ring to the bitmap and returns null when nothing is left', () => {
    const data = bitmap(40, 20, [0, 0, 0], { x: 0, y: 0, w: 40, h: 20 }, [9, 9, 9])
    expect(ringMedianColor(data, 40, 20, { x: 0, y: 0, w: 40, h: 20 }, 3)).toBeNull()
    const partial = bitmap(40, 20, [200, 100, 50], { x: 0, y: 0, w: 30, h: 20 }, [1, 1, 1])
    expect(ringMedianColor(partial, 40, 20, { x: 0, y: 0, w: 30, h: 20 }, 3)).toEqual([
      200, 100, 50,
    ])
  })
})
