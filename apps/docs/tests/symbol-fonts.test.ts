import { describe, expect, it } from 'vitest'

import { isSymbolFontFamily } from '@genoffice/ui'

describe('isSymbolFontFamily (font pickers show symbol-font names in the UI font)', () => {
  it('flags the classic symbol-encoded families', () => {
    for (const family of [
      'Webdings',
      'Wingdings',
      'Wingdings 2',
      'Wingdings 3',
      'Zapf Dingbats',
      'ITC Zapf Dingbats',
      'Symbol',
      'Marlett',
      'MS Outlook',
      'MS Reference Specialty',
      'MT Extra',
      'Bookshelf Symbol 7',
      'Segoe MDL2 Assets',
      'Segoe Fluent Icons',
      'HoloLens MDL2 Assets',
      'Bodoni Ornaments',
    ]) {
      expect(isSymbolFontFamily(family), family).toBe(true)
    }
  })

  it('is case- and whitespace-insensitive', () => {
    expect(isSymbolFontFamily('wingdings')).toBe(true)
    expect(isSymbolFontFamily(' WEBDINGS ')).toBe(true)
  })

  it('keeps text fonts on the preview path, including Unicode symbol fonts', () => {
    for (const family of [
      'Arial',
      'Calibri',
      '宋体',
      'PingFang SC',
      'Zapfino',
      // Unicode-encoded: letters render normally, so the preview stays useful
      'Segoe UI Symbol',
      'Apple Symbols',
      'Noto Sans Symbols',
      'Symbols Nerd Font',
    ]) {
      expect(isSymbolFontFamily(family), family).toBe(false)
    }
  })
})
