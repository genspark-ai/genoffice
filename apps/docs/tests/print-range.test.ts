import { describe, expect, it } from 'vitest'
import { parsePrintRange } from '../src/renderer/print-range'

describe('parsePrintRange', () => {
  it('parses single pages and ranges into 0-based indices', () => {
    expect(parsePrintRange('1,3,5-8', 10)).toEqual([0, 2, 4, 5, 6, 7])
  })

  it('dedupes overlaps and sorts', () => {
    expect(parsePrintRange('5-6, 2, 5', 10)).toEqual([1, 4, 5])
  })

  it('accepts CJK separators and en-dash', () => {
    expect(parsePrintRange('1，3、5–6', 6)).toEqual([0, 2, 4, 5])
  })

  it('rejects empty, malformed, reversed, and out-of-bounds input', () => {
    expect(parsePrintRange('', 5)).toBeNull()
    expect(parsePrintRange('a-b', 5)).toBeNull()
    expect(parsePrintRange('4-2', 5)).toBeNull()
    expect(parsePrintRange('0', 5)).toBeNull()
    expect(parsePrintRange('6', 5)).toBeNull()
    expect(parsePrintRange('1-9', 5)).toBeNull()
  })
})
