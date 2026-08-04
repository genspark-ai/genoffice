import { describe, expect, it } from 'vitest'

import { NUMBER_FORMAT_CATEGORIES } from '../src/renderer/number-format'
import {
  DATE_PATTERNS,
  DEFAULT_NUMFMT_OPTIONS,
  FRACTION_PATTERNS,
  NEGATIVE_STYLES,
  numfmtOptionsOf,
  numfmtPattern,
  numfmtPreview,
  sampleValue,
  TIME_PATTERNS,
  todaySerial,
  type NumfmtOptions,
} from '../src/renderer/numfmt-dialog'

const options = (overrides: Partial<NumfmtOptions>): NumfmtOptions => ({
  ...DEFAULT_NUMFMT_OPTIONS,
  ...overrides,
})

describe('numfmtPattern', () => {
  it.each([
    [options({ category: 'general' }), 'General'],
    [options({ category: 'number' }), '0.00'],
    [options({ category: 'number', decimals: 0, thousands: true }), '#,##0'],
    [options({ category: 'number', negative: 'red' }), '0.00;[Red]0.00'],
    [
      options({ category: 'number', thousands: true, negative: 'red-parens' }),
      '#,##0.00_);[Red](#,##0.00)',
    ],
    [options({ category: 'currency', symbol: '$' }), '$#,##0.00'],
    [options({ category: 'currency', symbol: '$', negative: 'parens' }), '$#,##0.00_);($#,##0.00)'],
    [options({ category: 'currency', symbol: 'US$' }), '"US$"#,##0.00'],
    [
      options({ category: 'accounting', symbol: '$' }),
      '_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
    ],
    [
      options({ category: 'accounting', symbol: '', decimals: 0 }),
      '_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
    ],
    [options({ category: 'percentage', decimals: 1 }), '0.0%'],
    [options({ category: 'scientific', decimals: 3 }), '0.000E+00'],
    [options({ category: 'text' }), '@'],
    [options({ category: 'custom', custom: '0.0"kg"' }), '0.0"kg"'],
  ])('%o → %s', (input, expected) => {
    expect(numfmtPattern(input)).toBe(expected)
  })
})

describe('round-trip', () => {
  it('number: every decimals × thousands × negative combination', () => {
    for (const decimals of [0, 2, 5, 30]) {
      for (const thousands of [false, true]) {
        for (const negative of NEGATIVE_STYLES) {
          const input = options({ category: 'number', decimals, thousands, negative })
          const parsed = numfmtOptionsOf(numfmtPattern(input))
          expect(parsed).toMatchObject({ category: 'number', decimals, thousands, negative })
        }
      }
    }
  })

  it('currency: symbols (incl. multi-char) × negatives', () => {
    for (const symbol of ['$', '¥', '€', 'US$', 'HK$']) {
      for (const negative of NEGATIVE_STYLES) {
        const input = options({ category: 'currency', symbol, negative, decimals: 2 })
        const parsed = numfmtOptionsOf(numfmtPattern(input))
        expect(parsed).toMatchObject({ category: 'currency', symbol, negative, decimals: 2 })
      }
    }
  })

  it('accounting: symbols × decimals', () => {
    for (const symbol of ['$', '¥', 'US$', '']) {
      for (const decimals of [0, 2, 4]) {
        const input = options({ category: 'accounting', symbol, decimals })
        const parsed = numfmtOptionsOf(numfmtPattern(input))
        expect(parsed).toMatchObject({ category: 'accounting', symbol, decimals })
      }
    }
  })

  it('date/time/fraction lists', () => {
    for (const datePattern of DATE_PATTERNS) {
      expect(numfmtOptionsOf(datePattern)).toMatchObject({ category: 'date', datePattern })
    }
    for (const timePattern of TIME_PATTERNS) {
      expect(numfmtOptionsOf(timePattern)).toMatchObject({ category: 'time', timePattern })
    }
    for (const fractionPattern of FRACTION_PATTERNS) {
      expect(numfmtOptionsOf(fractionPattern)).toMatchObject({
        category: 'fraction',
        fractionPattern,
      })
    }
  })

  it.each([
    ['General', 'general'],
    ['', 'general'],
    ['@', 'text'],
    ['0%', 'percentage'],
    ['0.00E+00', 'scientific'],
  ])('%s → %s', (pattern, category) => {
    expect(numfmtOptionsOf(pattern).category).toBe(category)
  })
})

describe('unparseable patterns land in custom', () => {
  it.each([
    '0.0"kg"',
    '0.00%;[Red]0.00%',
    'yyyy mm 0.00',
    '#,##0.00;',
    '[$USD-409] #,##0.00',
    'd-mmm-yyyy',
  ])('%s', (pattern) => {
    expect(numfmtOptionsOf(pattern)).toMatchObject({ category: 'custom', custom: pattern })
  })
})

describe('ribbon convergence', () => {
  it('dropdown patterns parse back into the matching dialog category', () => {
    const expected: Record<string, string> = {
      General: 'general',
      Number: 'number',
      Currency: 'currency',
      Accounting: 'accounting',
      'Short Date': 'date',
      'Long Date': 'date',
      Time: 'time',
      Percentage: 'percentage',
      Fraction: 'fraction',
      Scientific: 'scientific',
      Text: 'text',
    }
    for (const { label, pattern } of NUMBER_FORMAT_CATEGORIES) {
      expect(numfmtOptionsOf(pattern).category).toBe(expected[label])
    }
  })
})

describe('preview', () => {
  it('formats numbers with the pattern', () => {
    expect(numfmtPreview('#,##0.00', 1234.56)).toBe('1,234.56')
    expect(numfmtPreview('0.00%', 0.5)).toBe('50.00%')
    expect(numfmtPreview('', 1234.56)).toBe('1234.56')
  })

  it('renders negatives per style', () => {
    expect(numfmtPreview('#,##0.00_);(#,##0.00)', -1234.56)).toContain('(1,234.56)')
  })

  it('sampleValue prefers the anchor number, then category stand-ins', () => {
    expect(sampleValue('number', 42, 100)).toBe(42)
    expect(sampleValue('date', null, 45000)).toBe(45000)
    expect(sampleValue('text', 'abc', 100)).toBe('abc')
    expect(sampleValue('number', null, 100)).toBe(1234.56)
  })

  it('todaySerial is a plausible 1900-system serial', () => {
    const serial = todaySerial(new Date(2026, 0, 1, 12, 0, 0))
    expect(serial).toBeCloseTo(46023.5, 3)
  })
})
