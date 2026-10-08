import { describe, expect, it } from 'vitest'
import { formatAutoNum } from '../src/auto-num'

describe('auto-num bounds', () => {
  it('degrades an astronomical roman startAt instead of looping forever', () => {
    // toRoman(1e20) built an unbounded string for 36 s and climbed past 4 GB
    // of RSS on main; the roman system tops out at 3999 by construction.
    const t0 = performance.now()
    const label = formatAutoNum(1e20, 'romanLcPeriod')
    expect(performance.now() - t0).toBeLessThan(50)
    expect(label).toBe('100000000000000000000.')
  })

  it('keeps sane roman numerals unchanged', () => {
    expect(formatAutoNum(1, 'romanLcPeriod')).toBe('i.')
    expect(formatAutoNum(1994, 'romanUcPeriod')).toBe('MCMXCIV.')
  })

  it('clamps a startAt below 1 instead of indexing past the CJK digit table', () => {
    expect(formatAutoNum(-1, 'ea1ChsPeriod')).toBe('\u4e00.')
    expect(formatAutoNum(0, 'ea1ChtPlain')).toBe('\u4e00')
    expect(formatAutoNum(-5, 'arabicDbPeriod')).toBe('\uff11.')
    expect(formatAutoNum(0, 'arabicPeriod')).toBe('1.')
    expect(formatAutoNum(Number.NaN, 'alphaLcPeriod')).toBe('a.')
  })

  it('guards alpha the same way', () => {
    expect(formatAutoNum(1e20, 'alphaLcPeriod')).toBe('100000000000000000000.')
    expect(formatAutoNum(27, 'alphaLcPeriod')).toBe('aa.')
  })
})
