import { describe, expect, it } from 'vitest'
import {
  columnLabel,
  formatAddress,
  parseAddress,
  parseRange,
  rangeAddresses,
} from '../src/domain/cell-address'

describe('cell-address hardening', () => {
  it('stays case-strict like the sheets consumer contract', () => {
    // apps/sheets/tests/cell-address.test.ts pins lowercase rejection;
    // this package must not accept what the consumer refuses.
    expect(() => parseAddress('a1')).toThrow('Invalid cell address')
  })

  it('parses ranges with surrounding whitespace', () => {
    expect(parseRange(' A1 : B2 ')).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 1,
      endColumn: 1,
    })
  })

  it('parses absolute $-anchored addresses', () => {
    expect(parseAddress('$C$33')).toEqual({ row: 32, column: 2 })
  })

  it('rejects negative column labels', () => {
    expect(() => columnLabel(-1)).toThrow(RangeError)
  })

  it('rejects negative formatAddress coordinates', () => {
    expect(() => formatAddress(-1, -1)).toThrow(RangeError)
  })

  it('still expands small ranges', () => {
    expect(rangeAddresses({ startRow: 0, startColumn: 0, endRow: 1, endColumn: 1 })).toEqual([
      'A1',
      'B1',
      'A2',
      'B2',
    ])
  })
})
