import { describe, expect, it } from 'vitest'

import { shrinkToFitFontSize } from '../src/renderer/univer-sync'

// 10px per character at the base size; scales linearly with font size.
const measure = (line: string): number => line.length * 10

describe('shrinkToFitFontSize', () => {
  it('returns null when the text already fits', () => {
    expect(shrinkToFitFontSize('1980', 12, 50, measure)).toBeNull()
  })

  it('scales the font down proportionally and floors to an integer', () => {
    // 5 chars * 10px = 50px into 42px: 12 * 42/50 = 10.08 -> 10
    expect(shrinkToFitFontSize('1980X', 12, 42, measure)).toBe(10)
  })

  it('uses the widest line of a multiline cell', () => {
    expect(shrinkToFitFontSize('ab\r\nabcdef', 12, 30, measure)).toBe(6)
  })

  it('clamps at 1pt and rejects a zero-width column', () => {
    expect(shrinkToFitFontSize('abcdefghij', 12, 1, measure)).toBe(1)
    expect(shrinkToFitFontSize('abc', 12, 0, measure)).toBeNull()
  })
})
