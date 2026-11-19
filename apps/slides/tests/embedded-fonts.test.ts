import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'
import type { RunStyle } from '@genoffice/pptx-render'

vi.mock('../src/main/shaped-metrics', () => ({
  initShapedMetrics: () => {},
  shapedMeasure: () => null,
  shapedFamily: () => null,
  gtMeasure: () => null,
  complexScriptOf: () => null,
}))

import {
  createSystemFontMetrics,
  listPrivateFontFaces,
  registerEmbeddedFonts,
} from '../src/main/fonts'

const style = (fontFamily: string, over: Partial<RunStyle> = {}): RunStyle => ({
  fontFamily,
  fontSizePx: 100,
  bold: false,
  italic: false,
  ...over,
})

describe('document-embedded fonts', () => {
  it('a bold-only embed serves plain runs too (PowerPoint draws the only embedded face)', () => {
    const sfnt = new Uint8Array(
      readFileSync(join(__dirname, '../../docs/src/renderer/fonts/Caladea-Bold.ttf')),
    )
    expect(registerEmbeddedFonts([{ typeface: 'Zz Embed Bold Only', style: 'bold', sfnt }])).toBe(
      true,
    )
    const m = createSystemFontMetrics()
    expect(m.displayFamily!(style('Zz Embed Bold Only'))).toMatch(/Caladea/)
    expect(m.measure('MMMM', style('Zz Embed Bold Only'))).toBe(
      m.measure('MMMM', style('Zz Embed Bold Only', { bold: true })),
    )
  })
})

describe('alias families split by weight across files', () => {
  const office = existsSync('/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts')
  it.runIf(process.platform === 'darwin' && office)(
    'Yu Gothic UI bold resolves to the Bold collection, registering a distinct bold face',
    () => {
      const m = createSystemFontMetrics()
      m.displayFamily!(style('Yu Gothic UI'))
      m.displayFamily!(style('Yu Gothic UI', { bold: true }))
      const ids = listPrivateFontFaces().map((f) => f.id)
      expect(ids).toContain('Yu Gothic UI|00')
      expect(ids).toContain('Yu Gothic UI|10')
    },
  )
})
