import type { ICellData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { applyPivotStyling } from '../src/renderer/univer-sync'

/// Ground truth from POI sample 54436: <location ref="A8:B11"
/// firstHeaderRow="1" firstDataRow="1"/> where A8 = "Row Labels" (header),
/// A9/A10 = data, A11 = Grand Total. firstDataRow is the offset of the FIRST
/// DATA row inside the ref, so header rows are offsets 0..firstDataRow-1.
function paint(pivot: {
  outputRef: string
  headerFill?: string
  firstDataRow?: number
  rowGrandTotals?: boolean
}): boolean[] {
  const range = { startRow: 0, startColumn: 0, endRow: 19, endColumn: 4 }
  const matrix: ICellData[][] = Array.from({ length: 20 }, () =>
    Array.from({ length: 5 }, () => ({})),
  )
  applyPivotStyling(matrix, range, [
    { path: 'xl/pivotTables/pivotTable1.xml', cachePath: null, ...pivot },
  ])
  return matrix.map((row) => Boolean(row[0]?.s))
}

describe('applyPivotStyling', () => {
  it('paints only the header offset(s) below firstDataRow and the grand-total row', () => {
    // Sample 54436 shape: A8:B11, firstDataRow=1 → header A8, data A9/A10, total A11.
    const painted = paint({
      outputRef: 'A8:B11',
      headerFill: '#DCE6F1',
      firstDataRow: 1,
      rowGrandTotals: true,
    })
    expect(painted[7]).toBe(true) // A8 "Row Labels" header
    expect(painted[8]).toBe(false) // A9 first data row stays unpainted
    expect(painted[9]).toBe(false) // A10 data
    expect(painted[10]).toBe(true) // A11 Grand Total
  })

  it('paints two header rows when firstDataRow is 2', () => {
    const painted = paint({
      outputRef: 'A1:C6',
      headerFill: '#DCE6F1',
      firstDataRow: 2,
      rowGrandTotals: true,
    })
    expect(painted[0]).toBe(true)
    expect(painted[1]).toBe(true)
    expect(painted[2]).toBe(false) // first data row
    expect(painted[5]).toBe(true) // grand total
  })

  it('leaves the last row unpainted when rowGrandTotals is false', () => {
    const painted = paint({
      outputRef: 'A8:B11',
      headerFill: '#DCE6F1',
      firstDataRow: 1,
      rowGrandTotals: false,
    })
    expect(painted[7]).toBe(true)
    expect(painted[10]).toBe(false)
  })

  it('paints nothing without a style fill', () => {
    const painted = paint({ outputRef: 'A8:B11', firstDataRow: 1, rowGrandTotals: true })
    expect(painted.every((row) => !row)).toBe(true)
  })
})
