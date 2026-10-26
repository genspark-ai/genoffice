/**
 * The native "+" new-file menu and the tab-overload menu load their icons from
 * `src/main/assets/menu-*.png`: a 16x16 1x image plus a hand-attached 32x32
 * @2x representation (electron-vite's hashed asset names hide the `@2x`
 * sibling nativeImage would otherwise find on its own).
 *
 * A 1x PNG whose every pixel is fully transparent renders as NO icon at 100%
 * DPI — the menu shows the 2x representation only above 100% — which is how
 * the "+" menu's AI PDF entry lost its icon at 100% scaling (genoffice#2133):
 * menu-pdf.png had shipped as an all-zero-alpha image while menu-pdf@2x.png
 * carried the real pixels (menu-home.png had the same defect).
 *
 * Pinned here:
 * - every 1x icon is visible at 100% DPI (not an empty image);
 * - every 1x icon carries the same artwork as its @2x sibling, judged on the
 *   premultiplied mean color and the alpha coverage (tolerant of resampling
 *   and hand-earthed edges, so the pair can never drift into "no icon" again).
 */
import { readFileSync, readdirSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { PNG } from 'pngjs'
import { describe, expect, it } from 'vitest'

const assets = join(dirname(fileURLToPath(import.meta.url)), '../src/main/assets')

interface RepresentationStats {
  /** premultiplied mean color per channel, 0-255 */
  color: [number, number, number]
  /** mean alpha, 0-255 */
  alpha: number
}

function stats(png: PNG): RepresentationStats {
  let r = 0
  let g = 0
  let b = 0
  let a = 0
  const count = (png.width * png.height) | 0
  for (let i = 0; i < count; i++) {
    const at = i * 4
    const alpha = png.data[at + 3]! / 255
    r += png.data[at]! * alpha
    g += png.data[at + 1]! * alpha
    b += png.data[at + 2]! * alpha
    a += png.data[at + 3]!
  }
  const color: [number, number, number] = [
    Math.round(r / count),
    Math.round(g / count),
    Math.round(b / count),
  ]
  return { color, alpha: Math.round(a / count) }
}

describe('native menu icons', () => {
  const icons = readdirSync(assets).filter(
    (name) => /^menu-.+\.png$/.test(name) && !name.includes('@2x'),
  )

  it('every icon has a 2x sibling', () => {
    for (const name of icons)
      expect(readFileSync(join(assets, name.replace('.png', '@2x.png'))).length).toBeGreaterThan(0)
  })

  it('the 1x representation is visible at 100% DPI (not fully transparent)', () => {
    for (const name of icons) {
      const png = PNG.sync.read(readFileSync(join(assets, name)))
      expect(png.width, name).toBe(16)
      expect(png.height, name).toBe(16)
      let visible = false
      for (let i = 3; i < png.data.length; i += 4) if (png.data[i]! > 0) visible = true
      expect(visible, `${name} renders an icon at 100% DPI`).toBe(true)
    }
  })

  it('the 1x representation carries the same artwork as the @2x sibling', () => {
    for (const name of icons) {
      const big = PNG.sync.read(readFileSync(join(assets, name.replace('.png', '@2x.png'))))
      expect(big.width, name).toBe(32)
      expect(big.height, name).toBe(32)
      const one = stats(PNG.sync.read(readFileSync(join(assets, name))))
      const two = stats(big)
      // a broken 1x (e.g. fully transparent) sits hundreds of counts away from
      // the real artwork; a faithful resample stays within a few
      expect(one.alpha, name).toBeGreaterThan(2)
      expect(Math.abs(one.alpha - two.alpha)!, name).toBeLessThanOrEqual(12)
      for (let c = 0; c < 3; c++)
        expect(
          Math.abs(one.color[c]! - two.color[c]!)!,
          `${name} channel ${c} stays with the @2x artwork`,
        ).toBeLessThanOrEqual(16)
    }
  })
})
