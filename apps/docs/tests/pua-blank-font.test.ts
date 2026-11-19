/**
 * Contract of the bundled PUA blanker (tools/build-pua-blank-font.py): every
 * BMP Private Use codepoint maps to a blank 1em glyph, so AI-residue PUA
 * tokens stay invisible (Chromium never system-falls-back for PUA and would
 * otherwise draw the primary font's .notdef box).
 */
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { advanceEm, readWoff2 } from './helpers/woff2-metrics'

const font = readWoff2(join(__dirname, '../src/renderer/fonts/GenOfficePUABlank.woff2'))

describe('GenOffice PUA Blank', () => {
  it('covers the whole BMP Private Use Area at a 1em advance', () => {
    for (const cp of [0xe000, 0xe200, 0xe202, 0xf0b7, 0xf8ff]) {
      expect(font.cmap.get(cp)).toBeDefined()
      expect(advanceEm(font, cp)).toBe(1)
    }
    expect(font.cmap.get(0xdfff)).toBeUndefined()
    expect(font.cmap.get(0xf900)).toBeUndefined()
  })

  it('maps everything to one glyph besides .notdef (blank by construction)', () => {
    expect(new Set(font.cmap.values()).size).toBe(1)
  })
})
