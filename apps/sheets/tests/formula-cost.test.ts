import { describe, expect, it } from 'vitest'
import { SetRangeValuesCommand, SetRangeValuesMutation } from '@univerjs/sheets'
import { SET_RANGE_VALUES_COMMAND, SET_RANGE_VALUES_MUTATION } from '../src/renderer/app-constants'
import {
  collectCellFormulaTexts,
  degradeQuadraticFormulaCells,
  estimateQuadraticCost,
  MAX_QUADRATIC_COST,
  quadraticFormulaError,
} from '../src/renderer/formula-cost'

const SHEETS = [
  { name: '明细', rows: 88_588, columns: 98 },
  { name: 'Summary', rows: 40, columns: 10 },
]

describe('estimateQuadraticCost', () => {
  it('flags the distinct-count COUNTIF idiom that froze the app', () => {
    const cost = estimateQuadraticCost(
      '=SUMPRODUCT(1/COUNTIF(明细!D2:D88588,明细!D2:D88588))',
      'Summary',
      SHEETS,
    )
    expect(cost).toBe(88_587 * 88_587)
    expect(cost).toBeGreaterThan(MAX_QUADRATIC_COST)
  })

  it('flags the blank-tolerant distinct-count variant (criteria with &"")', () => {
    const cost = estimateQuadraticCost(
      '=SUMPRODUCT((D2:D88588<>"")/COUNTIF(D2:D88588,D2:D88588&""))',
      '明细',
      SHEETS,
    )
    expect(cost).toBeGreaterThan(MAX_QUADRATIC_COST)
  })

  it('flags whole-column criteria clamped to the sheet extent', () => {
    const cost = estimateQuadraticCost('=COUNTIF(D:D,D:D)', '明细', SHEETS)
    expect(cost).toBe(88_588 * 88_588)
  })

  it('flags SUMIFS with a range-valued criteria argument', () => {
    const cost = estimateQuadraticCost(
      '=SUMIFS(明细!E2:E88588,明细!D2:D88588,明细!D2:D88588)',
      'Summary',
      SHEETS,
    )
    expect(cost).toBeGreaterThan(MAX_QUADRATIC_COST)
  })

  it('flags array-context MATCH (lookup value is a large range)', () => {
    const cost = estimateQuadraticCost(
      '=SUM(--(MATCH(明细!D2:D88588,明细!D2:D88588,0)=ROW(明细!D2:D88588)-1))',
      'Summary',
      SHEETS,
    )
    expect(cost).toBeGreaterThan(MAX_QUADRATIC_COST)
  })

  it('allows plain linear aggregation over large ranges', () => {
    expect(estimateQuadraticCost('=SUM(明细!D2:D88588)', 'Summary', SHEETS)).toBe(0)
    expect(estimateQuadraticCost('=COUNTA(明细!D2:D88588)', 'Summary', SHEETS)).toBe(0)
    expect(
      estimateQuadraticCost('=SUMPRODUCT(明细!E2:E88588,明细!F2:F88588)', 'Summary', SHEETS),
    ).toBe(0)
  })

  it('allows scalar-criteria COUNTIF/SUMIFS over large ranges', () => {
    expect(estimateQuadraticCost('=COUNTIF(明细!D2:D88588,"ACME")', 'Summary', SHEETS)).toBe(0)
    expect(estimateQuadraticCost('=COUNTIF(明细!D:D,A2)', 'Summary', SHEETS)).toBe(0)
    expect(
      estimateQuadraticCost('=SUMIFS(明细!E2:E88588,明细!D2:D88588,A2)', 'Summary', SHEETS),
    ).toBe(0)
  })

  it('allows small-range criteria pairs (cost below the limit)', () => {
    const cost = estimateQuadraticCost('=COUNTIF(B2:B40,B2:B40)', 'Summary', SHEETS)
    expect(cost).toBe(39 * 39)
    expect(cost).toBeLessThan(MAX_QUADRATIC_COST)
  })

  it('ignores function names inside string literals', () => {
    expect(estimateQuadraticCost('=CONCAT("use COUNTIF(D:D,D:D)",A1)', '明细', SHEETS)).toBe(0)
  })

  it('handles nested calls and commas inside array literals', () => {
    const cost = estimateQuadraticCost(
      '=SUMPRODUCT(1/COUNTIF(D2:D88588,IF({1,2},D2:D88588)))',
      '明细',
      SHEETS,
    )
    expect(cost).toBeGreaterThan(MAX_QUADRATIC_COST)
  })
})

describe('quadraticFormulaError', () => {
  it('returns a model-facing error pointing at aggregate_range', () => {
    const error = quadraticFormulaError(
      '=SUMPRODUCT(1/COUNTIF(明细!D2:D88588,明细!D2:D88588))',
      'Summary',
      SHEETS,
    )
    expect(error).toContain('freeze')
    expect(error).toContain('aggregate_range')
  })

  it('returns null for safe formulas', () => {
    expect(quadraticFormulaError('=SUM(明细!D2:D88588)', 'Summary', SHEETS)).toBeNull()
  })
})

describe('collectCellFormulaTexts', () => {
  it('finds formulas in single cells, arrays, and matrix payloads', () => {
    expect(collectCellFormulaTexts({ f: '=SUM(A1:A3)' })).toEqual(['=SUM(A1:A3)'])
    expect(collectCellFormulaTexts([[{ v: 1 }, { f: '=A1+1' }], [{ f: '=A1*2' }]])).toEqual([
      '=A1+1',
      '=A1*2',
    ])
    expect(
      collectCellFormulaTexts({ 3: { 5: { f: '=COUNTIF(D:D,D:D)' }, 6: { v: 'x' } } }),
    ).toEqual(['=COUNTIF(D:D,D:D)'])
  })

  it('ignores payloads without formulas', () => {
    expect(collectCellFormulaTexts({ v: 42 })).toEqual([])
    expect(collectCellFormulaTexts(null)).toEqual([])
    expect(collectCellFormulaTexts('text')).toEqual([])
  })

  it('finds formulas in a paste-shaped SetRangeValuesMutation cellValue', () => {
    // Univer's clipboard paste applies the mutation directly (no
    // SetRangeValuesCommand) with this exact params shape.
    const params = {
      unitId: 'file-abc',
      subUnitId: 'sheet-1',
      cellValue: {
        5: { 3: { f: '=SUMPRODUCT(1/COUNTIF(D2:D88588,D2:D88588))', v: null, p: null } },
      },
    }
    expect(collectCellFormulaTexts(params.cellValue)).toEqual([
      '=SUMPRODUCT(1/COUNTIF(D2:D88588,D2:D88588))',
    ])
  })
})

describe('edit gate command ids', () => {
  it('match the installed Univer command and mutation ids', () => {
    expect(SET_RANGE_VALUES_COMMAND).toBe(SetRangeValuesCommand.id)
    expect(SET_RANGE_VALUES_MUTATION).toBe(SetRangeValuesMutation.id)
  })
})

describe('degradeQuadraticFormulaCells', () => {
  const cells = [
    { row: 0, column: 0, value: 7, formula: '=SUM(D2:D88588)' },
    { row: 1, column: 0, value: 42, formula: '=SUMPRODUCT(1/COUNTIF(D2:D88588,D2:D88588))' },
    { row: 2, column: 0, value: 'plain' },
    {
      row: 3,
      column: 0,
      value: 3,
      formula: '=COUNTIF(D:D,D:D)',
      arrayRef: 'A4:A5',
      styleIndex: 2,
    },
  ]

  it('strips only over-budget formulas, keeping cached values and styles', () => {
    const result = degradeQuadraticFormulaCells(cells, '明细', SHEETS)
    expect(result).not.toBe(cells)
    expect(result[0]).toBe(cells[0])
    expect(result[1]).toEqual({ row: 1, column: 0, value: 42 })
    expect(result[2]).toBe(cells[2])
    expect(result[3]).toEqual({ row: 3, column: 0, value: 3, styleIndex: 2 })
    // Input cells are untouched (save-side data keeps the file formula).
    expect(cells[1]?.formula).toBe('=SUMPRODUCT(1/COUNTIF(D2:D88588,D2:D88588))')
  })

  it('returns the input array untouched when every formula is allowed', () => {
    const allowed = [
      { row: 0, column: 0, value: 7, formula: '=SUM(D2:D88588)' },
      { row: 1, column: 0, value: 1, formula: '=COUNTIF(B2:B40,B2:B40)' },
    ]
    expect(degradeQuadraticFormulaCells(allowed, 'Summary', SHEETS)).toBe(allowed)
  })
})
