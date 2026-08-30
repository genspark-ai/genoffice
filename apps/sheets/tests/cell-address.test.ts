import { describe, expect, it } from 'vitest'

import { formatAddress, parseAddress, parseRange } from '../src/domain/cell-address'

describe('parseAddress', () => {
  it('parses plain and $-anchored A1 notation alike', () => {
    expect(parseAddress('C33')).toEqual({ row: 32, column: 2 })
    // pivot location refs arrive as $C$33 from some producers; refreshing
    // such a pivot used to throw "Invalid cell address"
    expect(parseAddress('$C$33')).toEqual({ row: 32, column: 2 })
    expect(parseAddress('C$33')).toEqual({ row: 32, column: 2 })
    expect(parseAddress('$C33')).toEqual({ row: 32, column: 2 })
  })

  it('still rejects malformed addresses', () => {
    for (const bad of ['33', 'C', 'C0', '$$C$33', 'c33', 'C33:D34']) {
      expect(() => parseAddress(bad)).toThrow('Invalid cell address')
    }
  })

  it('round-trips through formatAddress', () => {
    expect(formatAddress(32, 2)).toBe('C33')
  })
})

describe('parseRange', () => {
  it('accepts $-anchored range refs', () => {
    expect(parseRange('$C$33:$H$40')).toEqual({
      startRow: 32,
      startColumn: 2,
      endRow: 39,
      endColumn: 7,
    })
  })
})
