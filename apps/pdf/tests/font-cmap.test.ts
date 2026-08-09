import { existsSync, readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { fontCoversText } from '../src/main/font-cmap'

/** Same candidates as EDIT_FONT_PATHS.arial; the test runs wherever one exists */
const SANS_PATHS = [
  '/System/Library/Fonts/Supplemental/Arial.ttf',
  'C:\\Windows\\Fonts\\arial.ttf',
  '/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf',
]

describe('fontCoversText', () => {
  it('maps Latin text incl. currency/trademark symbols, but not CJK', () => {
    const path = SANS_PATHS.find((p) => existsSync(p))
    if (!path) return
    const font = readFileSync(path)
    // € (U+20AC) and ™ (U+2122) once fell outside a Latin-block allowlist; the cmap says yes
    expect(fontCoversText(font, 'Total €42 ™')).toBe(true)
    expect(fontCoversText(font, '应付金额')).toBe(false)
    expect(fontCoversText(font, 'mixed 金额')).toBe(false)
  })

  it('checks Unicode spaces as drawn glyphs instead of exempting them', () => {
    const path = SANS_PATHS.find((p) => existsSync(p))
    if (!path) return
    const font = readFileSync(path)
    expect(fontCoversText(font, 'a b')).toBe(true)
    // U+3000 (ideographic space) is drawn but unmapped in all three curated faces;
    // a \s-based strip used to exempt it and let the face emit .notdef
    expect(fontCoversText(font, 'a　b')).toBe(false)
  })

  it('treats line-break-only text as covered and unparseable bytes as not', () => {
    expect(fontCoversText(Buffer.alloc(0), '\r\n')).toBe(true)
    expect(fontCoversText(Buffer.alloc(64), 'x')).toBe(false)
  })
})
