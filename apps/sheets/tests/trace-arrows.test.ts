import { describe, expect, it } from 'vitest'

import { collectDependents, collectPrecedents } from '../src/renderer/trace-arrows'
import type { DependentsSource } from '../src/renderer/trace-arrows'

const bounds = { lastRow: 49, lastColumn: 9, maxRows: 1000, maxColumns: 100 }

describe('collectPrecedents', () => {
  it('collects same-sheet cell and range references', () => {
    const result = collectPrecedents('=SUM(B2:B10)+$C$1', 'Sheet1', bounds)
    expect(result.offSheet).toBe(0)
    expect(result.areas).toEqual([
      { startRow: 1, endRow: 9, startColumn: 1, endColumn: 1 },
      { startRow: 0, endRow: 0, startColumn: 2, endColumn: 2 },
    ])
  })

  it('keeps qualified references to the active sheet and counts others', () => {
    const result = collectPrecedents("=Sheet1!A1+'My Data'!B2+Other!C3", 'My Data', bounds)
    expect(result.offSheet).toBe(2)
    expect(result.areas).toEqual([{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 }])
  })

  it('clamps whole-column references to the used range', () => {
    const result = collectPrecedents('=SUM(D:D)', 'Sheet1', bounds)
    expect(result.areas).toEqual([{ startRow: 0, endRow: 49, startColumn: 3, endColumn: 3 }])
  })

  it('dedupes repeated references', () => {
    const result = collectPrecedents('=A1+A1*A1', 'Sheet1', bounds)
    expect(result.areas).toHaveLength(1)
  })

  it('ignores references inside string literals', () => {
    const result = collectPrecedents('=IF(A1>0,"see B2:B9",0)', 'Sheet1', bounds)
    expect(result.areas).toEqual([{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }])
  })
})

describe('collectDependents', () => {
  const source = {
    sheets: {
      s1: {
        name: 'Sheet1',
        cellData: {
          0: { 0: { v: 10 }, 1: { f: '=A1*2' } },
          4: { 2: { f: '=SUM(A1:A3)' }, 3: { f: '=SUM(B1:B3)' } },
          9: { 0: { f: '=Sheet2!A1' } },
        },
      },
      s2: {
        name: 'Sheet2',
        cellData: {
          0: { 0: { f: '=Sheet1!A1+1' }, 1: { f: '=A1' } },
        },
      },
    },
    // Real IWorkbookData cells also carry v/s etc.; the type only declares f — cast to skip excess-property checks
  } as DependentsSource

  it('finds same-sheet dependents, including range hits', () => {
    const result = collectDependents(source, 'Sheet1', { row: 0, column: 0 })
    expect(result.cells).toEqual([
      { row: 0, column: 1 },
      { row: 4, column: 2 },
    ])
    expect(result.offSheet).toBe(1)
  })

  it('resolves unqualified references against their own sheet', () => {
    const result = collectDependents(source, 'Sheet2', { row: 0, column: 0 })
    expect(result.cells).toEqual([{ row: 0, column: 1 }])
    expect(result.offSheet).toBe(1)
  })

  it('excludes the target cell itself', () => {
    const self = {
      sheets: { s1: { name: 'Sheet1', cellData: { 0: { 0: { f: '=A1' } } } } },
    }
    const result = collectDependents(self, 'Sheet1', { row: 0, column: 0 })
    expect(result.cells).toEqual([])
    expect(result.offSheet).toBe(0)
  })
})
