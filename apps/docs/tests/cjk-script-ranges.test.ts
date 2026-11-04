import { describe, expect, it } from 'vitest'
import { cjkScriptRanges, balancedSpaceWideningFor } from '../src/renderer/line-metrics'

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

describe('balancedSpaceWideningFor', () => {
  it('needs the balance flag plus useFELayout or a CJK document East Asian face', () => {
    expect(balancedSpaceWideningFor({ balanceDbcsSpacing: true, useFELayout: true })).toBe(true)
    expect(balancedSpaceWideningFor({ balanceDbcsSpacing: true }, '\u5b8b\u4f53')).toBe(true)
    expect(balancedSpaceWideningFor({ balanceDbcsSpacing: true }, 'Malgun Gothic')).toBe(true)
    expect(balancedSpaceWideningFor({ balanceDbcsSpacing: true }, 'DengXian')).toBe(true)
    expect(balancedSpaceWideningFor({ balanceDbcsSpacing: true }, 'Calibri')).toBe(false)
    expect(balancedSpaceWideningFor({ balanceDbcsSpacing: true })).toBe(false)
    expect(balancedSpaceWideningFor({ useFELayout: true }, '\u5b8b\u4f53')).toBe(false)
    expect(balancedSpaceWideningFor({})).toBe(false)
  })
})
