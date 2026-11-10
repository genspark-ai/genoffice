import { describe, expect, it } from 'vitest'
import {
  dateFormatFromActions,
  formatDate,
  isDateOnlyFormat,
  isValidDateInput,
  normalizeDateInput,
  readDateParts,
  toIsoDate,
} from '../src/renderer/date-format'

describe('dateFormatFromActions', () => {
  it('reads the AFDate_*Ex format string from keystroke or format scripts', () => {
    expect(dateFormatFromActions({ Keystroke: ['AFDate_KeystrokeEx("yyyy-mm-dd");'] })).toBe(
      'yyyy-mm-dd',
    )
    expect(dateFormatFromActions({ Format: ["AFDate_FormatEx('dd/mm/yyyy');"] })).toBe('dd/mm/yyyy')
  })
  it('maps the legacy AFDate_Keystroke(n) index', () => {
    expect(dateFormatFromActions({ Keystroke: ['AFDate_Keystroke(2);'] })).toBe('mm/dd/yy')
  })
  it('ignores fields without date scripts', () => {
    expect(dateFormatFromActions(undefined)).toBeNull()
    expect(
      dateFormatFromActions({ Keystroke: ['AFNumber_Keystroke(2, 0, 0, 0, "", true);'] }),
    ).toBeNull()
  })
})

describe('isValidDateInput', () => {
  it('accepts real dates in the given format and the empty value', () => {
    expect(isValidDateInput('2026-10-10', 'yyyy-mm-dd')).toBe(true)
    expect(isValidDateInput('', 'yyyy-mm-dd')).toBe(true)
    expect(isValidDateInput('02/29/2024', 'mm/dd/yyyy')).toBe(true)
    expect(isValidDateInput('3/7/26', 'm/d/yy')).toBe(true)
    expect(isValidDateInput('Mar 7, 2026', 'mmm d, yyyy')).toBe(true)
    expect(isValidDateInput('3/7/26 4:05 pm', 'm/d/yy h:MM tt')).toBe(true)
  })
  it('rejects free text, wrong separators and impossible dates', () => {
    expect(isValidDateInput('hello', 'yyyy-mm-dd')).toBe(false)
    expect(isValidDateInput('10/10/2026', 'yyyy-mm-dd')).toBe(false)
    expect(isValidDateInput('2026-13-01', 'yyyy-mm-dd')).toBe(false)
    expect(isValidDateInput('2026-02-30', 'yyyy-mm-dd')).toBe(false)
    expect(isValidDateInput('02/29/2023', 'mm/dd/yyyy')).toBe(false)
    expect(isValidDateInput('Foo 7, 2026', 'mmm d, yyyy')).toBe(false)
    expect(isValidDateInput('3/7/26 13:05 pm', 'm/d/yy h:MM tt')).toBe(false)
  })
})

describe('normalizeDateInput', () => {
  it('keeps an exact match as typed', () => {
    expect(normalizeDateInput(' 2026-10-10 ', 'yyyy-mm-dd')).toBe('2026-10-10')
    expect(normalizeDateInput('', 'yyyy-mm-dd')).toBe('')
  })
  it('rewrites other common layouts into the field format', () => {
    expect(normalizeDateInput('10/10/2026', 'yyyy-mm-dd')).toBe('2026-10-10')
    expect(normalizeDateInput('2026/10/10', 'yyyy-mm-dd')).toBe('2026-10-10')
    expect(normalizeDateInput('2026.1.5', 'yyyy-mm-dd')).toBe('2026-01-05')
    expect(normalizeDateInput('20261010', 'yyyy-mm-dd')).toBe('2026-10-10')
    expect(normalizeDateInput('05102026', 'dd/mm/yyyy')).toBe('05/10/2026')
    expect(normalizeDateInput('2026-10-05', 'dd/mm/yyyy')).toBe('05/10/2026')
    expect(normalizeDateInput('5 Oct 2026', 'dd/mm/yyyy')).toBe('05/10/2026')
    expect(normalizeDateInput('October 5, 2026', 'mm/dd/yy')).toBe('10/05/26')
    expect(normalizeDateInput('2026-10-05', 'mmm d, yyyy')).toBe('Oct 5, 2026')
    expect(normalizeDateInput('10-5-26', 'mm/dd/yyyy')).toBe('10/05/2026')
  })
  it('uses the other day/month order only when the format order is impossible', () => {
    expect(normalizeDateInput('25/10/2026', 'mm/dd/yyyy')).toBe('10/25/2026')
    expect(normalizeDateInput('2026-25-10', 'yyyy-mm-dd')).toBe('2026-10-25')
    expect(normalizeDateInput('5/10/2026', 'mm/dd/yyyy')).toBe('05/10/2026')
  })
  it('rejects what is no date', () => {
    expect(normalizeDateInput('abc', 'yyyy-mm-dd')).toBeNull()
    expect(normalizeDateInput('2026-13-40', 'yyyy-mm-dd')).toBeNull()
    expect(normalizeDateInput('2026-10', 'yyyy-mm-dd')).toBeNull()
    expect(normalizeDateInput('10 Foo 2026', 'yyyy-mm-dd')).toBeNull()
  })
  it('stays strict for formats with a time or a missing part', () => {
    expect(normalizeDateInput('2026-10-10', 'm/d/yy HH:MM')).toBeNull()
    expect(normalizeDateInput('2026-10', 'mm/yy')).toBeNull()
    expect(normalizeDateInput('10/26', 'mm/yy')).toBe('10/26')
  })
})

describe('readDateParts / formatDate / toIsoDate', () => {
  it('round-trips through ISO for the picker', () => {
    const parts = readDateParts('5/10/26', 'm/d/yy')
    expect(parts).toEqual({ year: 2026, month: 5, day: 10 })
    expect(toIsoDate(parts!)).toBe('2026-05-10')
    expect(formatDate({ year: 2026, month: 5, day: 10 }, 'm/d/yy')).toBe('5/10/26')
    expect(formatDate({ year: 1999, month: 12, day: 1 }, 'mmmm d, yyyy')).toBe('December 1, 1999')
    expect(readDateParts('10/26', 'mm/yy')).toBeNull()
  })
})

describe('isDateOnlyFormat', () => {
  it('is true only for year+month+day formats without a time', () => {
    expect(isDateOnlyFormat('yyyy-mm-dd')).toBe(true)
    expect(isDateOnlyFormat('mmm d, yyyy')).toBe(true)
    expect(isDateOnlyFormat('m/d/yy HH:MM')).toBe(false)
    expect(isDateOnlyFormat('mm/yy')).toBe(false)
  })
})
