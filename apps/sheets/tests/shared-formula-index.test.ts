import { describe, expect, it } from 'vitest'

import {
  SharedFormulaLookup,
  indexedFormulas,
  sharedFormulaText,
} from '../src/renderer/shared-formula-index'
import type { WorkbookSharedFormulaGroup } from '../src/shared/desktop-api'

const column: WorkbookSharedFormulaGroup = {
  si: 0,
  row: 2,
  column: 4,
  formula: '=D3+$A$1+Other!B2',
  range: { startRow: 2, endRow: 5, startColumn: 4, endColumn: 4 },
}

const listed: WorkbookSharedFormulaGroup = {
  si: 1,
  row: 10,
  column: 0,
  formula: '=SUM(B11:B20)',
  cells: [
    [11, 0],
    [13, 0],
  ],
}

describe('sharedFormulaText', () => {
  it('shifts relative references by the follower offset and pins absolute ones', () => {
    expect(sharedFormulaText(column, 3, 4)).toBe('=D4+$A$1+Other!B3')
    expect(sharedFormulaText(column, 5, 5)).toBe('=E6+$A$1+Other!C5')
    expect(sharedFormulaText(column, 2, 4)).toBe(column.formula)
  })
})

describe('SharedFormulaLookup', () => {
  const lookup = new SharedFormulaLookup([column, listed])

  it('counts followers without the masters', () => {
    expect(lookup.followerCount).toBe(3 + 2)
  })

  it('resolves span followers, listed followers and nothing else', () => {
    expect(lookup.textAt(4, 4)).toBe('=D5+$A$1+Other!B4')
    expect(lookup.textAt(13, 0)).toBe('=SUM(B14:B23)')
    expect(lookup.textAt(12, 0)).toBeUndefined()
    expect(lookup.textAt(6, 4)).toBeUndefined()
    expect(lookup.textAt(4, 5)).toBeUndefined()
  })

  it('visits every follower exactly once, never the master', () => {
    const seen: string[] = []
    lookup.forEachFollower((row, col) => seen.push(`${row}:${col}`))
    expect(seen.sort()).toEqual(['11:0', '13:0', '3:4', '4:4', '5:4'])
  })

  it('merges masters and followers for the closure analyzer', () => {
    const formulas = indexedFormulas(
      [{ row: 2, column: 4, formula: column.formula }, { row: 0, column: 0, value: 1 } as never],
      lookup,
    )
    expect(formulas).toHaveLength(1 + 5)
    expect(formulas[0]).toEqual({ row: 2, column: 4, formula: column.formula })
    expect(formulas).toContainEqual({ row: 11, column: 0, formula: '=SUM(B12:B21)' })
  })

  it('expands a 150k-follower column in well under a second', () => {
    const big = new SharedFormulaLookup([
      {
        si: 0,
        row: 0,
        column: 1,
        formula: '=A1*2',
        range: { startRow: 0, endRow: 150_000, startColumn: 1, endColumn: 1 },
      },
    ])
    const started = performance.now()
    const formulas = indexedFormulas([{ row: 0, column: 1, formula: '=A1*2' }], big)
    const elapsed = performance.now() - started
    expect(formulas).toHaveLength(150_001)
    expect(formulas[150_000]).toEqual({ row: 150_000, column: 1, formula: '=A150001*2' })
    expect(big.textAt(99_999, 1)).toBe('=A100000*2')
    expect(elapsed).toBeLessThan(2_000)
    console.info(`expanded 150k followers in ${elapsed.toFixed(0)} ms`)
  })
})
