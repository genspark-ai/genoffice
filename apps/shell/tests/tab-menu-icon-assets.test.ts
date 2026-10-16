import { describe, expect, it } from 'vitest'
import { existsSync, readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { PNG } from 'pngjs'
import type { TabMenuIcon } from '../src/main/tab-menu-icons'

const ALL_18_ACTIONS: TabMenuIcon[] = [
  'save',
  'saveAs',
  'print',
  'rename',
  'duplicate',
  'splitBeside',
  'openInNewWindow',
  'copyPath',
  'copyName',
  'reveal',
  'openDefault',
  'moveLeft',
  'moveRight',
  'close',
  'closeOthers',
  'closeRight',
  'closeAll',
  'reopen',
]

const pngDir = join(__dirname, '../src/main/assets/tab-actions/png')
const svgDir = join(__dirname, '../src/main/assets/tab-actions/svg')

describe('tab menu icon assets integrity', () => {
  it('all 18 source SVG files exist and are valid markup', () => {
    for (const action of ALL_18_ACTIONS) {
      const svgPath = join(svgDir, `${action}.svg`)
      expect(existsSync(svgPath), `SVG for ${action} must exist`).toBe(true)
      const content = readFileSync(svgPath, 'utf8')
      expect(content).toContain('<svg')
      expect(content).toContain('</svg>')
      expect(content).toContain('viewBox="0 0 24 24"')
      expect(content).toContain('stroke="currentColor"')
    }
  })

  it('all 72 required PNG variants exist (1x/2x for light and dark across 18 actions)', () => {
    for (const action of ALL_18_ACTIONS) {
      const light1x = join(pngDir, `${action}.png`)
      const light2x = join(pngDir, `${action}@2x.png`)
      const dark1x = join(pngDir, `${action}-dark.png`)
      const dark2x = join(pngDir, `${action}-dark@2x.png`)

      expect(existsSync(light1x), `Missing ${action}.png`).toBe(true)
      expect(existsSync(light2x), `Missing ${action}@2x.png`).toBe(true)
      expect(existsSync(dark1x), `Missing ${action}-dark.png`).toBe(true)
      expect(existsSync(dark2x), `Missing ${action}-dark@2x.png`).toBe(true)
    }

    const allPngs = readdirSync(pngDir).filter((f) => f.endsWith('.png'))
    expect(allPngs.length).toBe(72)
  })

  it('PNG dimensions and headers are strictly valid (16x16 for 1x, 32x32 for 2x)', () => {
    for (const action of ALL_18_ACTIONS) {
      for (const variant of ['', '-dark']) {
        const p1 = PNG.sync.read(readFileSync(join(pngDir, `${action}${variant}.png`)))
        expect(p1.width, `${action}${variant}.png width`).toBe(16)
        expect(p1.height, `${action}${variant}.png height`).toBe(16)

        const p2 = PNG.sync.read(readFileSync(join(pngDir, `${action}${variant}@2x.png`)))
        expect(p2.width, `${action}${variant}@2x.png width`).toBe(32)
        expect(p2.height, `${action}${variant}@2x.png height`).toBe(32)
      }
    }
  })

  it('background transparency is preserved in all four corners across all 72 PNGs', () => {
    const allPngs = readdirSync(pngDir).filter((f) => f.endsWith('.png'))
    for (const file of allPngs) {
      const png = PNG.sync.read(readFileSync(join(pngDir, file)))
      const w = png.width
      const h = png.height
      const cornerAlphas = [
        png.data[3], // top-left (0, 0)
        png.data[(w - 1) * 4 + 3], // top-right (w-1, 0)
        png.data[(h - 1) * w * 4 + 3], // bottom-left (0, h-1)
        png.data[((h - 1) * w + w - 1) * 4 + 3], // bottom-right (w-1, h-1)
      ]
      for (let i = 0; i < 4; i++) {
        expect(cornerAlphas[i], `${file} corner ${i} must be fully transparent (alpha 0)`).toBe(0)
      }
    }
  })

  it('every icon has visible non-zero glyph pixels and is never an empty or solid rectangle', () => {
    const allPngs = readdirSync(pngDir).filter((f) => f.endsWith('.png'))
    for (const file of allPngs) {
      const png = PNG.sync.read(readFileSync(join(pngDir, file)))
      const totalPixels = png.width * png.height
      let nonZeroAlpha = 0
      let fullyTransparent = 0

      for (let i = 0; i < png.data.length; i += 4) {
        const a = png.data[i + 3]
        if (a > 0) nonZeroAlpha++
        if (a === 0) fullyTransparent++
      }

      // Must have visible glyph lines (at least 5% coverage)
      expect(nonZeroAlpha, `${file} glyph must have non-zero pixels`).toBeGreaterThan(
        totalPixels * 0.05,
      )
      // Must NOT be a solid square (must have at least 20% transparent background)
      expect(fullyTransparent, `${file} must preserve background transparency`).toBeGreaterThan(
        totalPixels * 0.2,
      )
    }
  })

  it('light and dark variants have distinct, theme-appropriate contrast', () => {
    for (const action of ALL_18_ACTIONS) {
      const light = PNG.sync.read(readFileSync(join(pngDir, `${action}.png`)))
      const dark = PNG.sync.read(readFileSync(join(pngDir, `${action}-dark.png`)))

      // Find average luminance of glyph pixels (alpha > 128)
      let lightLumaSum = 0
      let lightCount = 0
      for (let i = 0; i < light.data.length; i += 4) {
        if (light.data[i + 3] > 128) {
          lightLumaSum += (light.data[i] + light.data[i + 1] + light.data[i + 2]) / 3
          lightCount++
        }
      }

      let darkLumaSum = 0
      let darkCount = 0
      for (let i = 0; i < dark.data.length; i += 4) {
        if (dark.data[i + 3] > 128) {
          darkLumaSum += (dark.data[i] + dark.data[i + 1] + dark.data[i + 2]) / 3
          darkCount++
        }
      }

      const avgLightLuma = lightCount > 0 ? lightLumaSum / lightCount : 0
      const avgDarkLuma = darkCount > 0 ? darkLumaSum / darkCount : 255

      // Light icon glyphs must be dark (#1f2328, luma < 80)
      expect(avgLightLuma, `${action} light glyph average luminance`).toBeLessThan(80)
      // Dark icon glyphs must be light (#f0f0f0, luma > 200)
      expect(avgDarkLuma, `${action} dark glyph average luminance`).toBeGreaterThan(200)
    }
  })
})
