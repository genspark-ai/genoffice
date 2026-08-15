import { describe, expect, it } from 'vitest'
import { rowOfVisIdx, spreadRows, stepPage } from '../src/renderer/spread'

describe('spreadRows', () => {
  it('gives every page its own row in single mode', () => {
    expect(spreadRows([10, 11, 12], 1)).toEqual([[10], [11], [12]])
  })

  it('pairs 1+2, 3+4, … in two-page mode (mainstream default, no cover row)', () => {
    expect(spreadRows([10, 11, 12, 13], 2)).toEqual([
      [10, 11],
      [12, 13],
    ])
  })

  it('leaves a trailing odd page alone in two-page mode', () => {
    expect(spreadRows([10, 11, 12], 2)).toEqual([[10, 11], [12]])
  })

  it('handles empty and single-page lists', () => {
    expect(spreadRows([], 2)).toEqual([])
    expect(spreadRows([7], 2)).toEqual([[7]])
  })
})

describe('stepPage', () => {
  const five = [0, 1, 2, 3, 4]

  it('steps one page at a time in single mode, clamped at the ends', () => {
    expect(stepPage(five, 1, 2, 1)).toBe(3)
    expect(stepPage(five, 1, 2, -1)).toBe(1)
    expect(stepPage(five, 1, 1, -1)).toBe(1)
    expect(stepPage(five, 1, 5, 1)).toBe(5)
  })

  it('steps a whole row in two-page mode so paired pages are never a dead end', () => {
    expect(stepPage(five, 2, 1, 1)).toBe(3) // row (1,2) -> row (3,4)
    expect(stepPage(five, 2, 2, 1)).toBe(3) // from the right page of a pair too
    expect(stepPage(five, 2, 3, -1)).toBe(1)
    expect(stepPage(five, 2, 4, 1)).toBe(5) // -> trailing odd row
    expect(stepPage(five, 2, 1, -1)).toBe(1) // clamped at the first row
    expect(stepPage(five, 2, 5, 1)).toBe(5) // clamped at the last row
  })
})

describe('rowOfVisIdx', () => {
  it('is the identity in single mode', () => {
    expect(rowOfVisIdx(0, 1)).toBe(0)
    expect(rowOfVisIdx(5, 1)).toBe(5)
  })

  it('maps pairs to the same row in two-page mode', () => {
    expect(rowOfVisIdx(0, 2)).toBe(0)
    expect(rowOfVisIdx(1, 2)).toBe(0)
    expect(rowOfVisIdx(2, 2)).toBe(1)
    expect(rowOfVisIdx(3, 2)).toBe(1)
    expect(rowOfVisIdx(4, 2)).toBe(2)
  })
})
