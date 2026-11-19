import { describe, expect, it } from 'vitest'

import {
  clampColorScaleStops,
  evaluateArithmetic,
  evaluateThresholdFormula,
  type ThresholdReader,
} from '../src/renderer/cf-thresholds'

const num = (value: string) => ({ kind: 'num', value })

/// A1 = 1 on the rule's sheet; nothing else populated.
const reader: ThresholdReader = {
  readValues: async (sheetName, range) =>
    sheetName === null && range.startRow === 0 && range.startColumn === 0 ? [1] : [],
  definedName: () => null,
  tableColumn: () => null,
}

describe('relative references in scale thresholds', () => {
  it('zeroes relative and mixed refs and evaluates the rest like Excel', async () => {
    // colorscale.xlsx sheet1 F3:F6 max threshold: Excel renders 2*A1+2 as 2
    // although A1 holds 1.
    expect(await evaluateThresholdFormula('2*A1+2', reader)).toBe(2)
    expect(await evaluateThresholdFormula('=2*A1+3', reader)).toBe(3)
    expect(await evaluateThresholdFormula('(1+B2)*4/2', reader)).toBe(2)
    expect(await evaluateThresholdFormula('$A1+A$1', reader)).toBe(0)
  })

  it('reads absolute refs', async () => {
    expect(await evaluateThresholdFormula('$A$1*2', reader)).toBe(2)
  })
})

describe('evaluateArithmetic', () => {
  it('folds parenthesised substitutions and exponent literals', () => {
    expect(evaluateArithmetic('2*(0)+3')).toBe(3)
    expect(evaluateArithmetic('3-(-5)')).toBe(8)
    expect(evaluateArithmetic('2*1e3')).toBe(2000)
    expect(evaluateArithmetic('(1.5)/(3)')).toBe(0.5)
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
