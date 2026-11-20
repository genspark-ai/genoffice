import { afterEach, describe, expect, it } from 'vitest'

import { CellValueType } from '@univerjs/core'
import {
  fixFormattedValue,
  formatForDisplay,
  formatGeneral,
  localizeNumberText,
  setNumberDisplayLocale,
} from '../src/renderer/numfmt-fix'
import { numfmtPreview } from '../src/renderer/numfmt-dialog'
import { copyEmfCellDisplay } from '../src/renderer/copy-emf'
import { formatAxisValue } from '../src/renderer/WorkbookVisuals'

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

  it('shows General numbers with a decimal comma under hu, en-US otherwise', () => {
    expect(formatGeneral(3.5, 11)).toBe('3.5')
    setNumberDisplayLocale('hu')
    expect(formatGeneral(3.5, 11)).toBe('3,5')
    expect(formatGeneral(-0.125, 11)).toBe('-0,125')
    expect(formatGeneral(123456789012345, 11)).toBe('1,23457E+14')
    expect(plain(localizeNumberText('1,234.5'))).toBe('1 234,5')
  })

  it('formats chart labels and copied-image cells like the grid', () => {
    setNumberDisplayLocale('hu')
    expect(plain(formatForDisplay('#,##0.00', 3600.5))).toBe('3 600,50')
    expect(plain(formatAxisValue(1234567, undefined))).toBe('1 234 567')
    expect(plain(formatAxisValue(3600.5, '#,##0.00'))).toBe('3 600,50')
    const pattern = { s: { n: { pattern: '#,##0.00' } } }
    expect(plain(copyEmfCellDisplay({ v: 3600.5, t: CellValueType.NUMBER, ...pattern }).v)).toBe('3 600,50')
    expect(copyEmfCellDisplay({ v: 3.5, t: CellValueType.NUMBER }).v).toBe('3,5')
  })
})
