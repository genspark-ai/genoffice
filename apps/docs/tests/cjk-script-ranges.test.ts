import { describe, expect, it } from 'vitest'
import { cjkScriptRanges, hangulSpaceWideningFor } from '../src/renderer/line-metrics'

describe('cjkScriptRanges', () => {
  const H = '한글'
  it('ends a lifted stretch at a space so the space keeps the paragraph line box', () => {
    expect(cjkScriptRanges(`${H} ${H}`)).toEqual([
      { from: 0, to: 2 },
      { from: 3, to: 5 },
    ])
    expect(cjkScriptRanges(H + '\u00a0' + H + ' A')).toEqual([
      { from: 0, to: 2 },
      { from: 3, to: 5 },
    ])
  })
  it('keeps a stretch across CJK punctuation-free syllables and drops Latin', () => {
    expect(cjkScriptRanges(`A ${H}${H}1`)).toEqual([{ from: 2, to: 6 }])
    expect(cjkScriptRanges('abc 123')).toEqual([])
  })
})

describe('hangulSpaceWideningFor', () => {
  it('widens hangul-context spaces only under both compat flags', () => {
    expect(hangulSpaceWideningFor({ balanceDbcsSpacing: true, useFELayout: true })).toBe(true)
    expect(hangulSpaceWideningFor({ balanceDbcsSpacing: true })).toBe(false)
    expect(hangulSpaceWideningFor({ useFELayout: true })).toBe(false)
    expect(hangulSpaceWideningFor({})).toBe(false)
  })
})
