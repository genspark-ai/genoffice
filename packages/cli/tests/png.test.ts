import { PNG } from 'pngjs'
import { describe, expect, it } from 'vitest'
import { contactSheet, cropPng } from '../src/formats/png'

function solid(width: number, height: number, rgb: [number, number, number]): Buffer {
  const png = new PNG({ width, height })
  for (let i = 0; i < width * height; i++) {
    png.data[i * 4] = rgb[0]
    png.data[i * 4 + 1] = rgb[1]
    png.data[i * 4 + 2] = rgb[2]
    png.data[i * 4 + 3] = 255
  }
  return PNG.sync.write(png)
}

describe('png helpers', () => {
  it('crops with padding and clamps to the image', () => {
    const { png, rect } = cropPng(solid(100, 80, [10, 20, 30]), { x: 90, y: 5, w: 20, h: 10 }, 8)
    expect(rect).toEqual({ x: 82, y: 0, w: 18, h: 23 })
    const out = PNG.sync.read(png)
    expect([out.width, out.height]).toEqual([18, 23])
    expect(Array.from(out.data.slice(0, 4))).toEqual([10, 20, 30, 255])
  })

  it('lays pages out row-major on a white sheet', () => {
    const pages = [1, 2, 3].map((page) => ({ page, png: solid(200, 100, [0, 0, 0]) }))
    const sheet = contactSheet(pages, 2, 50)
    expect([sheet.cols, sheet.rows]).toEqual([2, 2])
    expect([sheet.width, sheet.height]).toEqual([2 * 50 + 3 * 8, 2 * 25 + 3 * 8])
    expect(sheet.tiles.map((t) => [t.page, t.x, t.y])).toEqual([
      [1, 8, 8],
      [2, 66, 8],
      [3, 8, 41],
    ])
    const out = PNG.sync.read(sheet.png)
    const at = (x: number, y: number) => out.data[(y * out.width + x) * 4]
    expect(at(0, 0)).toBe(255)
    expect(at(10, 10)).toBe(0)
    expect(at(70, 45)).toBe(255)
  })
})
