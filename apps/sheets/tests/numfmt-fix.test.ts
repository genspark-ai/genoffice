import { describe, expect, it } from 'vitest'

import {
  decimalRoundForPattern,
  fixFormattedValue,
  formatGeneral,
  generalCharBudget,
} from '../src/renderer/numfmt-fix'

const NBSP = ' '

describe('fixFormattedValue — empty sections', () => {
  it('renders 0 as empty when the zero section is empty', () => {
    expect(fixFormattedValue('#,##0_);(#,##0);', 0, 0)).toBe('')
  })

  it('hides every value under ;;;', () => {
    expect(fixFormattedValue(';;;', 123, 123)).toBe('')
    expect(fixFormattedValue(';;;', 'abc', 'abc')).toBe('')
  })

  it('leaves non-zero values on the normal sections', () => {
    expect(fixFormattedValue('#,##0_);(#,##0);', 5, `5${NBSP}`.replace(NBSP, ' '))).toBe(`5${NBSP}`)
    expect(fixFormattedValue('#,##0_);(#,##0);', -5, '(5)')).toBeNull()
  })
})

describe('fixFormattedValue — _x padding and text section', () => {
  it('upgrades trailing padding spaces to NBSP so layout keeps them', () => {
    expect(fixFormattedValue('#,##0.0_);(#,##0.0)', 765.89, '765.9 ')).toBe(`765.9${NBSP}`)
  })

  it('pads the positive right edge by exactly one char (the width of `)`)', () => {
    const pos = fixFormattedValue('#,##0.0_);(#,##0.0)', 765.89, '765.9 ') as string
    expect(pos.endsWith(NBSP)).toBe(true)
    expect(pos).toHaveLength('765.9'.length + 1)
  })

  it('applies the 4th (text) section to string cells', () => {
    expect(fixFormattedValue('#,##0.0_);(#,##0.0);0.0_);@_)', 'abc', 'abc')).toBe(`abc${NBSP}`)
    expect(
      fixFormattedValue(
        '#,##0.0_);(#,##0.0);0.0_);@_)',
        'Training The Street',
        'Training The Street',
      ),
    ).toBe(`Training The Street${NBSP}`)
  })

  it('leaves strings alone when the pattern has no text section', () => {
    expect(fixFormattedValue('#,##0.0_);(#,##0.0)', 'abc', 'abc')).toBeNull()
  })

  it('does not touch values Univer formatted with a different rendering', () => {
    expect(fixFormattedValue('#,##0.0_);(#,##0.0)', 765.89, '765,9 ')).toBeNull()
  })
})

describe('formatGeneral', () => {
  it('shows 0.326883 at the default 8.43-char column width', () => {
    // characterWidthToPixels(8.43) === 64
    expect(generalCharBudget(64)).toBe(8)
    expect(formatGeneral(0.326882822311, 8)).toBe('0.326883')
  })

  it('shows more digits as the column widens, capped at 11 significant', () => {
    expect(formatGeneral(0.326882822311, 10)).toBe('0.32688282')
    expect(formatGeneral(0.326882822311, 20)).toBe('0.326882822')
  })

  it('keeps short values untouched', () => {
    expect(formatGeneral(123, 8)).toBe('123')
    expect(formatGeneral(44562.5, 8)).toBe('44562.5')
    expect(formatGeneral(0.5, 8)).toBe('0.5')
  })

  it('handles negatives (sign counts against the budget)', () => {
    expect(formatGeneral(-0.326882822311, 8)).toBe('-0.32688')
  })

  it('switches to scientific when the integer part cannot fit', () => {
    expect(formatGeneral(123456789012, 8)).toBe('1.23E+11')
    expect(formatGeneral(1234567890123456, 20)).toBe('1.23457E+15')
  })

  it('uses scientific for tiny fractions instead of rounding to 0', () => {
    expect(formatGeneral(0.0000001234, 9)).toBe('1.234E-07')
    expect(formatGeneral(0.0000001234, 8)).toBe('1.23E-07')
  })

  it('strips trailing zeros after shrinking', () => {
    expect(formatGeneral(0.100000004, 8)).toBe('0.1')
  })

  it('never crashes on degenerate budgets', () => {
    expect(generalCharBudget(0)).toBe(1)
    expect(typeof formatGeneral(0.326882822311, 1)).toBe('string')
  })
})

describe('decimal half-way rounding (Excel rounds the decimal literal)', () => {
  it('rounds 1.005 up under 0.00 like Excel', () => {
    expect(fixFormattedValue('0.00_ ', 1.005, `1.00${NBSP}`)).toBe(`1.01${NBSP}`)
  })

  it('leaves non-half-way values alone', () => {
    expect(fixFormattedValue('0.00_ ', 1.004, `1.00${NBSP}`)).toBeNull()
    expect(decimalRoundForPattern('0.00', 1.23)).toBe(1.23)
  })

  it('rounds negatives away from zero', () => {
    expect(decimalRoundForPattern('0', -2.5)).toBe(-3)
    expect(decimalRoundForPattern('0.00', -1.005)).toBe(-1.01)
  })

  it('pre-scales percent patterns', () => {
    expect(decimalRoundForPattern('0.00%', 0.12345)).toBeCloseTo(0.1235, 10)
  })

  it('skips multi-section, date, fraction and scientific patterns', () => {
    expect(decimalRoundForPattern('#,##0.0;(#,##0.00)', 1.005)).toBeNull()
    expect(decimalRoundForPattern('yyyy-mm-dd', 1.005)).toBeNull()
    expect(decimalRoundForPattern('# ?/?', 1.005)).toBeNull()
    expect(decimalRoundForPattern('0.00E+00', 1.005)).toBeNull()
  })
})
