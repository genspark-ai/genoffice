import { describe, expect, it } from 'vitest'
import {
  colorRunsEqual,
  colorSegments,
  colorsToRuns,
  decodeStyle,
  encodeStyle,
  mapCharColors,
  patchStyle,
  runsToColors,
  spliceCharColors,
} from '../src/renderer/color-runs'

const R = '#d32f2f'
const B = '#1a73e8'

describe('spliceCharColors', () => {
  it('keeps colors across an insertion, inherited from the left neighbor', () => {
    // 'abc' with b red → type 'X' after b: 'abXc'
    const out = spliceCharColors('abc', ['', R, ''], 'abXc')
    expect(out).toEqual(['', R, R, ''])
  })

  it('keeps colors across a deletion', () => {
    const out = spliceCharColors('abXc', ['', R, R, ''], 'abc')
    expect(out).toEqual(['', R, ''])
  })

  it('handles a full select-all replacement (colors reset to base)', () => {
    expect(spliceCharColors('abc', [R, R, R], 'xyzw')).toEqual(['', '', '', ''])
  })

  it('handles a mid-word replacement across a color boundary', () => {
    // 'aabb' (aa red, bb blue), replace 'ab' (idx 1-2) with 'ZZZ'
    const out = spliceCharColors('aabb', [R, R, B, B], 'aZZZb')
    expect(out).toEqual([R, R, R, R, B])
  })

  it('insertion at the very start has no left neighbor and stays base', () => {
    expect(spliceCharColors('ab', [R, R], 'Xab')).toEqual(['', R, R])
  })
})

describe('mapCharColors', () => {
  it('maps colors through a re-wrap that only rearranges whitespace', () => {
    // 'hello world' with 'world' red → wrapped as 'hello\nworld'
    const src = 'hello world'
    const colors = runsToColors(src.length, [{ start: 6, end: 11, color: R }])
    const out = mapCharColors(src, colors, 'hello\nworld')
    expect(colorsToRuns(out)).toEqual([{ start: 6, end: 11, color: R }])
  })

  it('maps back from wrapped text to the logical paragraph', () => {
    const wrapped = 'hello\nworld'
    const colors = runsToColors(wrapped.length, [{ start: 6, end: 11, color: R }])
    const out = mapCharColors(wrapped, colors, 'hello world')
    expect(colorsToRuns(out)).toEqual([{ start: 6, end: 11, color: R }])
  })

  it('whitespace inherits the preceding char color (colored word keeps its space)', () => {
    const src = 'ab cd'
    const colors = runsToColors(src.length, [{ start: 0, end: 2, color: R }])
    const out = mapCharColors(src, colors, 'ab cd')
    expect(out).toEqual([R, R, R, '', ''])
  })

  it('CJK re-wrap with no spaces maps one-to-one', () => {
    const src = '你好世界再见'
    const colors = runsToColors(src.length, [{ start: 2, end: 4, color: B }])
    const out = mapCharColors(src, colors, '你好世\n界再见')
    expect(out).toEqual(['', '', B, B, B, '', ''])
  })
})

describe('runs round-trip', () => {
  it('colorsToRuns merges adjacent equal colors and skips base', () => {
    expect(colorsToRuns(['', R, R, B, '', B])).toEqual([
      { start: 1, end: 3, color: R },
      { start: 3, end: 4, color: B },
      { start: 5, end: 6, color: B },
    ])
  })

  it('runsToColors ↔ colorsToRuns round-trips', () => {
    const runs = [
      { start: 1, end: 3, color: R },
      { start: 5, end: 6, color: B },
    ]
    expect(colorsToRuns(runsToColors(8, runs))).toEqual(runs)
  })

  it('colorRunsEqual compares structurally', () => {
    expect(colorRunsEqual([{ start: 1, end: 2, color: R }], [{ start: 1, end: 2, color: R }])).toBe(
      true,
    )
    expect(colorRunsEqual([{ start: 1, end: 2, color: R }], [{ start: 1, end: 2, color: B }])).toBe(
      false,
    )
    expect(colorRunsEqual(undefined, [])).toBe(true)
  })
})

describe('colorSegments', () => {
  it('splits text into consecutive same-color spans including base spans', () => {
    expect(colorSegments('aabbc', [R, R, '', '', B])).toEqual([
      { text: 'aa', color: R },
      { text: 'bb', color: '' },
      { text: 'c', color: B },
    ])
  })
})

describe('style keys', () => {
  it("round-trips every field and canonicalizes the empty style to ''", () => {
    expect(encodeStyle({})).toBe('')
    expect(decodeStyle('')).toEqual({})
    const full = { color: R, font: 'arial', size: 12.5, bold: true, italic: false }
    expect(decodeStyle(encodeStyle(full))).toEqual(full)
    // Explicit off is distinct from inherit
    expect(encodeStyle({ bold: false })).not.toBe(encodeStyle({}))
    expect(decodeStyle(encodeStyle({ bold: false }))).toEqual({ bold: false })
  })

  it('distinct styles never collide and identical styles always do', () => {
    expect(encodeStyle({ color: R })).not.toBe(encodeStyle({ font: R }))
    expect(encodeStyle({ color: R, bold: true })).toBe(encodeStyle({ bold: true, color: R }))
  })

  it('patchStyle merges fields and clears with null', () => {
    const key = encodeStyle({ color: R, bold: true })
    expect(decodeStyle(patchStyle(key, { italic: true }))).toEqual({
      color: R,
      bold: true,
      italic: true,
    })
    expect(decodeStyle(patchStyle(key, { bold: null }))).toEqual({ color: R })
    expect(patchStyle(encodeStyle({ bold: true }), { bold: null })).toBe('')
    // Untouched patch keeps the key intact
    expect(patchStyle(key, {})).toBe(key)
  })

  it('style keys flow through the generic run helpers unchanged', () => {
    const bold = encodeStyle({ bold: true })
    const out = spliceCharColors('abc', ['', bold, ''], 'abXc')
    expect(out).toEqual(['', bold, bold, ''])
    expect(colorsToRuns(out)).toEqual([{ start: 1, end: 3, color: bold }])
  })
})
