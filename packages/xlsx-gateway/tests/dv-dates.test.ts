import { describe, expect, it } from 'vitest'
import { applyDvRules, DvEditError } from '../src/gateway/xlsx-dv'

const SHEET = '<worksheet><sheetData/></worksheet>'

function formula1For(type: string, value: string): string | undefined {
  const xml = applyDvRules(SHEET, [
    {
      ranges: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
      rule: { type, formula1: value },
    },
  ])
  return /<formula1>(.*?)<\/formula1>/.exec(xml)?.[1]
}

describe('xlsx-dv date and time guards', () => {
  it('accepts leap-day 2024-02-29', () => {
    expect(formula1For('date', '2024-02-29')).toBe('45351')
  })

  it('rejects non-leap-day 2023-02-29 with passthrough', () => {
    expect(formula1For('date', '2023-02-29')).toBe('2023-02-29')
  })

  it('rejects month 13 with passthrough', () => {
    expect(formula1For('date', '2024-13-01')).toBe('2024-13-01')
  })

  it('rejects 24:00 with passthrough', () => {
    expect(formula1For('time', '24:00')).toBe('24:00')
  })

  it('converts a valid datetime to a serial spot-check', () => {
    expect(formula1For('date', '2024-01-01 12:00:00')).toBe('45292.5')
  })

  it('converts a valid time to a fraction', () => {
    expect(formula1For('time', '12:00')).toBe('0.5')
  })

  it('rejects a pre-1900 date instead of storing a 19xx serial', () => {
    // Date.UTC reads 0099 as 1999, which stored 36161 — a valid-looking
    // serial for the wrong date. Excel's epoch starts at 1900-01-01, so
    // there is no serial to write and the save must say so.
    expect(() => formula1For('date', '0099-01-01')).toThrow(DvEditError)
    expect(() => formula1For('date', '0099-01-01')).toThrow(/before 1900/)
    expect(() => formula1For('date', '1899-12-31')).toThrow(/before 1900/)
  })

  it('still converts a 1900 date', () => {
    // 1900-03-01 is the first day where "days since 1899-12-30" and Excel's
    // serial agree (61); earlier 1900 dates are off by Excel's phantom
    // 1900-02-29, which is pre-existing and not what this guards.
    expect(formula1For('date', '1900-03-01')).toBe('61')
  })
})
