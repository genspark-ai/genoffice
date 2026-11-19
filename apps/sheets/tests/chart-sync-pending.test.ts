import { describe, expect, it } from 'vitest'

import {
  buildStreamedChartGrid,
  cacheLessRefs,
  hasPendingFormulaCells,
  unionBounds,
} from '../src/renderer/chart-sync-pending'

describe('hasPendingFormulaCells', () => {
  it('flags a formula cell whose value the engine has not written yet', () => {
    // The E1 shape: L2:L4 rewritten as =B6/=C6/=D6, L5:L6 still cached.
    expect(
      hasPendingFormulaCells([
        { f: '=B6' },
        { f: '=C6', v: null },
        { f: '=D6', v: undefined },
        { f: '=E6', v: 142244.59 },
        { f: '=F6', v: 721658.28 },
      ]),
    ).toBe(true)
  })

  it('treats a shared-formula follower without a value as pending', () => {
    expect(hasPendingFormulaCells([{ si: 'abc' }])).toBe(true)
  })

  it('accepts evaluated formulas, including 0 and empty-string results', () => {
    expect(
      hasPendingFormulaCells([
        { f: '=B6', v: 0 },
        { f: '=C6', v: '' },
        { f: '=D6', v: '#REF!' },
        { si: 'abc', v: 12 },
      ]),
    ).toBe(false)
  })

  it('ignores constants and empty cells', () => {
    expect(hasPendingFormulaCells([{ v: 1 }, { v: null }, {}, null, undefined])).toBe(false)
    expect(hasPendingFormulaCells([])).toBe(false)
  })
})

describe('buildStreamedChartGrid', () => {
  // Summary!L2:L6 of a streamed workbook: the sidecar still holds the file's
  // cached results, the AI rewrote L2:L4 as =B6/=C6/=D6.
  const bounds = { startRow: 1, endRow: 5, startColumn: 11, endColumn: 11 }
  const screen = [
    { row: 1, column: 11, value: 232532.97, formula: 'SUM(B2:B5)' },
    { row: 2, column: 11, value: -250695.47, formula: 'SUM(C2:C5)' },
    { row: 3, column: 11, value: 129913.9, formula: 'SUM(D2:D5)' },
    { row: 4, column: 11, value: 142244.59, formula: 'SUM(E2:E5)' },
    { row: 5, column: 11, value: 721658.28, formula: 'SUM(F2:F5)' },
  ]
  const noFill = () => ({ found: false }) as const

  it('flags journaled formulas, which never receive an engine value', () => {
    const grid = buildStreamedChartGrid(
      bounds,
      screen,
      [
        { row: 1, column: 11, hasValue: true, value: null, formula: '=B6' },
        { row: 2, column: 11, hasValue: true, value: null, formula: '=C6' },
        { row: 3, column: 11, hasValue: true, value: null, formula: '=D6' },
      ],
      noFill,
    )
    expect(grid.pendingFormula).toBe(true)
    expect(grid.values).toEqual([[null], [null], [null], [142244.59], [721658.28]])
  })

  it('flags a sidecar formula cell saved without a cached value', () => {
    const grid = buildStreamedChartGrid(
      bounds,
      [{ row: 1, column: 11, value: null, formula: 'B6' }, ...screen.slice(1)],
      [],
      noFill,
    )
    expect(grid.pendingFormula).toBe(true)
  })

  it('reads cached sidecar values and constant edits as settled', () => {
    const grid = buildStreamedChartGrid(
      bounds,
      screen,
      [
        { row: 1, column: 11, hasValue: true, value: 10 },
        { row: 2, column: 11, hasValue: false, value: null, style: {} } as never,
      ],
      noFill,
    )
    expect(grid.pendingFormula).toBe(false)
    expect(grid.values).toEqual([[10], [-250695.47], [129913.9], [142244.59], [721658.28]])
  })

  it('a constant edit or bulk fill over a pending formula settles the cell', () => {
    const journal = [{ row: 1, column: 11, hasValue: true, value: null, formula: '=B6' }]
    const overwritten = buildStreamedChartGrid(
      bounds,
      screen,
      [...journal, { row: 1, column: 11, hasValue: true, value: 5 }],
      noFill,
    )
    expect(overwritten.pendingFormula).toBe(false)
    const filled = buildStreamedChartGrid(bounds, screen, journal, (row) =>
      row === 1 ? { found: true, value: 7 } : { found: false },
    )
    expect(filled.pendingFormula).toBe(false)
    expect(filled.values[0]).toEqual([7])
  })

  it('a cleared cell is not pending: the chart follows it to empty', () => {
    const grid = buildStreamedChartGrid(
      bounds,
      screen,
      [{ row: 1, column: 11, hasValue: true, value: null }],
      noFill,
    )
    expect(grid.pendingFormula).toBe(false)
    expect(grid.values[0]).toEqual([null])
  })
})

describe('cacheLessRefs', () => {
  const touched = (ref: string | undefined) => ref !== undefined && ref.startsWith('Summary!')

  it('returns the touched refs whose file vector is empty', () => {
    const series = {
      valuesRef: 'Summary!$L$2:$L$6',
      categoriesRef: 'Summary!$A$2:$A$6',
      values: [],
      categories: [],
    }
    expect(cacheLessRefs(series, touched)).toEqual(['Summary!$L$2:$L$6', 'Summary!$A$2:$A$6'])
  })

  it('skips cached vectors: those take the refresh path', () => {
    const series = {
      valuesRef: 'Summary!$L$2:$L$6',
      categoriesRef: 'Summary!$A$2:$A$6',
      values: [1, 2, 3, 4, 5],
      categories: [],
    }
    expect(cacheLessRefs(series, touched)).toEqual(['Summary!$A$2:$A$6'])
  })

  it('skips untouched and missing refs', () => {
    expect(
      cacheLessRefs(
        { valuesRef: 'Other!$L$2:$L$6', categoriesRef: undefined, values: [], categories: [] },
        touched,
      ),
    ).toEqual([])
  })

  it('a cache-less series over pending formulas is held, not baked', () => {
    // Values come straight from the grid, and the grid
    // still holds unevaluated =B6/=C6/=D6, so the sync must park the bounds.
    const series = { valuesRef: 'Summary!$L$2:$L$6', values: [], categories: [] }
    const grid = buildStreamedChartGrid(
      { startRow: 1, endRow: 5, startColumn: 11, endColumn: 11 },
      [],
      [
        { row: 1, column: 11, hasValue: true, value: null, formula: '=B6' },
        { row: 2, column: 11, hasValue: true, value: null, formula: '=C6' },
        { row: 3, column: 11, hasValue: true, value: null, formula: '=D6' },
      ],
      () => ({ found: false }),
    )
    const held = cacheLessRefs(series, touched).filter(() => grid.pendingFormula)
    expect(held).toEqual(['Summary!$L$2:$L$6'])
  })
})

describe('unionBounds', () => {
  it('returns the new bounds when nothing was pending', () => {
    const bounds = { startRow: 1, endRow: 3, startColumn: 11, endColumn: 11 }
    expect(unionBounds(undefined, bounds)).toBe(bounds)
  })

  it('grows to the enclosing rectangle', () => {
    expect(
      unionBounds(
        { startRow: 1, endRow: 3, startColumn: 11, endColumn: 11 },
        { startRow: 5, endRow: 5, startColumn: 1, endColumn: 2 },
      ),
    ).toEqual({ startRow: 1, endRow: 5, startColumn: 1, endColumn: 11 })
  })
})
