/**
 * The streamed patch pipeline must keep a real empty-string value (shared
 * string / inlineStr '') distinct from a style-only cell (`<c r="B3" s="170"/>`,
 * value null on the wire). Flattening both to `{ v: null }` silently dropped
 * literal '' cells from copies and saves; flattening both to `{ v: '' }` made
 * formula references return '' where Excel returns 0 (see
 * style-only-cell-ref.test.ts).
 */
import { CellValueType, type ICellData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { patchWorksheetRangeInner } from '../src/renderer/univer-sync'
import type { WorkbookRangeResult } from '../src/shared/desktop-api'

function patchedMatrix(cells: WorkbookRangeResult['cells']): ICellData[][] {
  let captured: ICellData[][] = []
  const worksheet = {
    getRange: () => ({
      setValues: (matrix: ICellData[][]) => {
        captured = matrix
      },
    }),
  }
  patchWorksheetRangeInner(
    worksheet as never,
    undefined,
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
    cells,
    [],
    [],
    [],
    null,
    false,
  )
  return captured
}

describe('patchWorksheetRangeInner empty cells', () => {
  it('keeps a literal empty-string cell as an empty STRING', () => {
    const matrix = patchedMatrix([{ row: 0, column: 1, value: '' }])
    expect(matrix[0]?.[1]).toEqual({ v: '', t: CellValueType.STRING })
  })

  it('keeps style-only cells value-less so references return 0 like Excel', () => {
    const matrix = patchedMatrix([{ row: 0, column: 0, value: null }])
    expect(matrix[0]?.[0]).toEqual({ v: null })
  })

  it('still types non-empty strings', () => {
    const matrix = patchedMatrix([{ row: 0, column: 2, value: 'x' }])
    expect(matrix[0]?.[2]).toEqual({ v: 'x', t: CellValueType.STRING })
  })
})
