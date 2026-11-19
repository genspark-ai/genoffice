import { describe, expect, it } from 'vitest'

import { isA1Reference, resolveGoToRef } from '../src/renderer/goto'

const NAMES = [
  { name: 'SalesData', ref: 'Sheet1!$A$1:$B$2' },
  { name: 'Spilled', ref: '=Sheet2!$C$3' },
  { name: 'Moving', ref: 'OFFSET(Sheet1!$A$1,1,1)' },
]

describe('isA1Reference', () => {
  it('accepts cells, ranges, and spans', () => {
    expect(isA1Reference('C5')).toBe(true)
    expect(isA1Reference('B2:D5')).toBe(true)
    expect(isA1Reference('$A$1:$B$2')).toBe(true)
    expect(isA1Reference('A:C')).toBe(true)
    expect(isA1Reference('3:8')).toBe(true)
    expect(isA1Reference('Sheet1!A1')).toBe(true)
    expect(isA1Reference("'My Sheet'!A1:B2")).toBe(true)
  })

  it('rejects names, formulas, and malformed refs', () => {
    expect(isA1Reference('SalesData')).toBe(false)
    expect(isA1Reference('OFFSET(A1,1,1)')).toBe(false)
    expect(isA1Reference('A1:B')).toBe(false)
    expect(isA1Reference('!A1')).toBe(false)
    expect(isA1Reference('A1,B2')).toBe(false)
    expect(isA1Reference('')).toBe(false)
  })

  it('rejects shape-valid refs outside the grid', () => {
    // Row 0, over-long columns, and over-max rows/cols used to pass the
    // shape regexes and leak into getRange as NaN/out-of-bounds.
    expect(isA1Reference('A0')).toBe(false)
    expect(isA1Reference('0:5')).toBe(false)
    expect(isA1Reference('ZZZ1')).toBe(false)
    expect(isA1Reference('A1048577')).toBe(false)
    expect(isA1Reference('A1:ZZZ2')).toBe(false)
    expect(isA1Reference('Sheet1!A0')).toBe(false)
  })

  it('accepts grid-boundary refs', () => {
    expect(isA1Reference('XFD1048576')).toBe(true)
    expect(isA1Reference('$XFD$1048576')).toBe(true)
    expect(isA1Reference('A:A')).toBe(true)
    expect(isA1Reference('1048576:1048576')).toBe(true)
  })
})

describe('resolveGoToRef', () => {
  it('returns a plain A1 address as typed', () => {
    expect(resolveGoToRef('C5', NAMES)).toBe('C5')
    expect(resolveGoToRef(' B2:D5 ', NAMES)).toBe('B2:D5')
  })

  it('resolves a defined name to its reference, case-insensitively', () => {
    expect(resolveGoToRef('SalesData', NAMES)).toBe('Sheet1!$A$1:$B$2')
    expect(resolveGoToRef('salesdata', NAMES)).toBe('Sheet1!$A$1:$B$2')
  })

  it('strips a leading = from a defined-name reference', () => {
    expect(resolveGoToRef('Spilled', NAMES)).toBe('Sheet2!$C$3')
  })

  it('returns null for unknown names, formula names, and empty input', () => {
    expect(resolveGoToRef('NoSuchName', NAMES)).toBeNull()
    expect(resolveGoToRef('Moving', NAMES)).toBeNull()
    expect(resolveGoToRef('   ', NAMES)).toBeNull()
    expect(resolveGoToRef('C5', [])).toBe('C5')
  })
})
