import { describe, expect, it } from 'vitest'

import { cellIsBlankDiverges } from '../src/renderer/univer-sync'

describe('cellIsBlankDiverges', () => {
  it('flags rules Excel matches on blanks (blank = 0) but Univer skips', () => {
    expect(cellIsBlankDiverges('equal', 0, NaN)).toBe(true)
    expect(cellIsBlankDiverges('lessThan', 5, NaN)).toBe(true)
    expect(cellIsBlankDiverges('lessThanOrEqual', 0, NaN)).toBe(true)
    expect(cellIsBlankDiverges('greaterThan', -1, NaN)).toBe(true)
    expect(cellIsBlankDiverges('greaterThanOrEqual', 0, NaN)).toBe(true)
    expect(cellIsBlankDiverges('between', -1, 1)).toBe(true)
  })

  it('flags rules Univer matches on blanks but Excel does not', () => {
    expect(cellIsBlankDiverges('notEqual', 0, NaN)).toBe(true)
    expect(cellIsBlankDiverges('notBetween', -1, 1)).toBe(true)
  })

  it('leaves agreeing rules on the native condition', () => {
    expect(cellIsBlankDiverges('equal', 1, NaN)).toBe(false)
    expect(cellIsBlankDiverges('greaterThan', 0, NaN)).toBe(false)
    expect(cellIsBlankDiverges('lessThan', 0, NaN)).toBe(false)
    expect(cellIsBlankDiverges('notEqual', 5, NaN)).toBe(false)
    expect(cellIsBlankDiverges('between', 1, 9)).toBe(false)
    expect(cellIsBlankDiverges('notBetween', 1, 9)).toBe(false)
    expect(cellIsBlankDiverges('unknown', 0, NaN)).toBe(false)
  })

  it('normalizes reversed between bounds like Excel', () => {
    expect(cellIsBlankDiverges('between', 1, -1)).toBe(true)
    expect(cellIsBlankDiverges('notBetween', 1, -1)).toBe(true)
  })
})
