import { existsSync, readFileSync } from 'node:fs'
import { resolve } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import * as opentype from 'opentype.js'

vi.mock('@genoffice/ui/fonts/Carlito-Regular.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Regular.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-Bold.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Bold.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-Italic.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-Italic.ttf'),
}))
vi.mock('@genoffice/ui/fonts/Carlito-BoldItalic.ttf?asset', () => ({
  default: resolve(process.cwd(), '../../packages/ui/src/fonts/Carlito-BoldItalic.ttf'),
}))
vi.mock('../src/main/shaped-metrics', () => ({
  initShapedMetrics: () => {},
  shapedMeasure: () => null,
  shapedFamily: () => null,
  gtMeasure: () => null,
  complexScriptOf: () => null,
}))

import { wrapSafeAdvance } from '../src/main/fonts'

const TNR = '/System/Library/Fonts/Supplemental/Times New Roman.ttf'
const ARIAL = '/System/Library/Fonts/Supplemental/Arial.ttf'
const CONSTANTIA = '/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts/Constan.ttf'

const load = (path: string) => {
  const buf = readFileSync(path)
  return opentype.parse(buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength))
}
const widths = (path: string, text: string) => {
  const font = wrapSafeAdvance(load(path))
  return {
    kerned: font.getAdvanceWidth(text, 100, { kerning: true }),
    plain: font.getAdvanceWidth(text, 100, { kerning: false }),
  }
}

describe.runIf(process.platform === 'darwin' && existsSync(TNR) && existsSync(ARIAL))(
  'safe-advance kerning',
  () => {
    it('kerns a font whose pairs live only in the kern table (GPOS without a kern feature)', () => {
      const { kerned, plain } = widths(TNR, 'Te Va')
      expect(kerned).toBeLessThan(plain)
    })

    it('kerns a font with a plain GPOS pair-adjustment lookup', () => {
      const { kerned, plain } = widths(ARIAL, 'Te Va')
      expect(kerned).toBeLessThan(plain)
    })

    it.runIf(existsSync(CONSTANTIA))(
      'kerns a font whose GPOS kern lookups are Extension-wrapped',
      () => {
        const { kerned, plain } = widths(CONSTANTIA, 'Te Va')
        expect(kerned).toBeLessThan(plain)
      },
    )

    it('leaves pairs the font does not kern alone', () => {
      const { kerned, plain } = widths(TNR, 'nn mm')
      expect(kerned).toBe(plain)
    })
  },
)
