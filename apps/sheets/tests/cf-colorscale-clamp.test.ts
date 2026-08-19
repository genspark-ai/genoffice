import { describe, expect, it } from 'vitest'

import {
  clampColorScaleStops,
  evaluateArithmetic,
  hasRelativeReference,
  substituteRelativeReferences,
} from '../src/renderer/univer-sync'

const num = (value: string) => ({ kind: 'num', value })

describe('hasRelativeReference', () => {
  it('flags relative and mixed cell references', () => {
    expect(hasRelativeReference('2*A1+3')).toBe(true)
    expect(hasRelativeReference('=$A1+1')).toBe(true)
    expect(hasRelativeReference('A$1')).toBe(true)
  })

  it('accepts absolute references, names, and functions', () => {
    expect(hasRelativeReference('$A$1*2')).toBe(false)
    expect(hasRelativeReference("'My Sheet'!$B$2")).toBe(false)
    expect(hasRelativeReference('MyName+1')).toBe(false)
    expect(hasRelativeReference('LOG10(5)')).toBe(false)
    expect(hasRelativeReference('"A1 in a string"')).toBe(false)
  })
})

describe('substituteRelativeReferences + evaluateArithmetic', () => {
  it('zeroes relative refs and evaluates the rest like Excel', () => {
    // colorscale.xlsx sheet1 F3:F6 max threshold: Excel renders 2*A1+2 as 2.
    expect(evaluateArithmetic(substituteRelativeReferences('2*A1+2'))).toBe(2)
    expect(evaluateArithmetic(substituteRelativeReferences('=2*A1+3'))).toBe(3)
    expect(evaluateArithmetic(substituteRelativeReferences('(1+B2)*4/2'))).toBe(2)
  })

  it('keeps absolute refs and string literals intact', () => {
    expect(substituteRelativeReferences('$A$1+B2')).toBe('$A$1+0')
    expect(substituteRelativeReferences('"A1"&C3')).toBe('"A1"&0')
  })

  it('returns null for non-arithmetic leftovers', () => {
    expect(evaluateArithmetic('$A$1+0')).toBeNull()
    expect(evaluateArithmetic('SUM(0)')).toBeNull()
    expect(evaluateArithmetic('1/0')).toBeNull()
    expect(evaluateArithmetic('')).toBeNull()
  })
})

describe('clampColorScaleStops', () => {
  it('lifts a later stop up to an earlier one and eps-steps the tie', () => {
    // colorscale.xlsx sheet2 F2:F7: num 0 / num 4 / rejected formula -> 0.
    const clamped = clampColorScaleStops([num('0'), num('4'), num('0')])
    expect(Number(clamped[0]!.value)).toBe(0)
    const mid = Number(clamped[1]!.value)
    const max = Number(clamped[2]!.value)
    expect(max).toBe(4)
    expect(mid).toBeLessThan(4)
    expect(mid).toBeGreaterThan(3.999)
  })

  it('returns the same array when stops are already ascending', () => {
    const stops = [num('0'), num('5'), num('10')]
    expect(clampColorScaleStops(stops)).toBe(stops)
  })

  it('skips non-numeric stops without clamping across them', () => {
    const stops = [{ kind: 'min' }, { kind: 'percent', value: '50' }, num('2')]
    expect(clampColorScaleStops(stops)).toBe(stops)
  })
})
