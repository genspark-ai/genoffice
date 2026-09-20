import { describe, expect, it } from 'vitest'
import { applyDvRules } from '../src/gateway/xlsx-dv'

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
})
