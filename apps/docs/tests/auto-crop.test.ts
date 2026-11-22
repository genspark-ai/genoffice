import { describe, expect, it } from 'vitest'
import { autoCropBox } from '../src/renderer/editor/auto-crop'

type Rgba = [number, number, number, number]

/** a picture of `bg` with one `fg` rectangle on it */
function picture(
  w: number,
  h: number,
  bg: Rgba,
  fg: Rgba,
  box: { x: number; y: number; w: number; h: number } | null,
) {
  const data = new Uint8ClampedArray(w * h * 4)
  for (let y = 0; y < h; y++)
    for (let x = 0; x < w; x++) {
      const inside = box && x >= box.x && x < box.x + box.w && y >= box.y && y < box.y + box.h
      data.set(inside ? fg : bg, (y * w + x) * 4)
    }
  return data
}

const WHITE: Rgba = [255, 255, 255, 255]
const RED: Rgba = [200, 0, 0, 255]
const CLEAR: Rgba = [0, 0, 0, 0]

describe('autoCropBox', () => {
  it('trims a border of the corner colour', () => {
    const box = { x: 3, y: 2, w: 5, h: 4 }
    expect(autoCropBox(picture(12, 9, WHITE, RED, box), 12, 9)).toEqual(box)
  })

  it('trims a transparent border, whatever colour hides in it', () => {
    const box = { x: 0, y: 4, w: 6, h: 2 }
    expect(autoCropBox(picture(6, 8, CLEAR, RED, box), 6, 8)).toEqual(box)
  })

  it('takes a border that is nearly the corner colour as border', () => {
    const data = picture(8, 8, WHITE, RED, { x: 2, y: 2, w: 4, h: 4 })
    data.set([250, 252, 249, 255], (1 * 8 + 1) * 4)
    expect(autoCropBox(data, 8, 8)).toEqual({ x: 2, y: 2, w: 4, h: 4 })
  })

  it('leaves a picture with no border, or nothing but border, alone', () => {
    expect(autoCropBox(picture(4, 4, WHITE, RED, { x: 0, y: 0, w: 4, h: 4 }), 4, 4)).toBeNull()
    expect(autoCropBox(picture(4, 4, WHITE, RED, null), 4, 4)).toBeNull()
    expect(autoCropBox(new Uint8ClampedArray(0), 0, 0)).toBeNull()
  })
})
