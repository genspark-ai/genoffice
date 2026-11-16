import { afterEach, describe, expect, it } from 'vitest'

import { fixFormattedValue, setNumberDisplayLocale } from '../src/renderer/numfmt-fix'
import { numfmtPreview } from '../src/renderer/numfmt-dialog'

// numfmt writes NBSP for both the group separator and padded literals; compare on plain spaces
const plain = (text: string | null): string | null => text && text.replace(/\u00a0/g, ' ')

describe('number display locale (Hungarian UI)', () => {
  afterEach(() => setNumberDisplayLocale(undefined))

  it('keeps en-US separators when no display locale is set', () => {
    expect(plain(fixFormattedValue('#,##0 "Ft"', 3600, '3,600 Ft')) ?? '3,600 Ft').toBe('3,600 Ft')
    expect(fixFormattedValue('#,##0.00', 3600.5, '3,600.50')).toBeNull()
  })

  it('renders Hungarian grouping and decimal comma under hu', () => {
    setNumberDisplayLocale('hu')
    expect(plain(fixFormattedValue('#,##0 "Ft"', 3600, '3,600 Ft'))).toBe('3 600 Ft')
    expect(plain(fixFormattedValue('#,##0.00', 3600.5, '3,600.50'))).toBe('3 600,50')
    expect(plain(fixFormattedValue('[$Ft-40E] #,##0', 3600, 'Ft 3,600'))).toBe('Ft 3 600')
  })

  it('previews the Format Cells dialog in the display locale', () => {
    setNumberDisplayLocale('hu')
    expect(plain(numfmtPreview('#,##0 "Ft"', 3600))).toBe('3 600 Ft')
  })
})
