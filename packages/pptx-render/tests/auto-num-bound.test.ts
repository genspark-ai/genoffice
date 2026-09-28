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

  it('guards alpha the same way', () => {
    expect(formatAutoNum(1e20, 'alphaLcPeriod')).toBe('100000000000000000000.')
    expect(formatAutoNum(27, 'alphaLcPeriod')).toBe('aa.')
  })
})
