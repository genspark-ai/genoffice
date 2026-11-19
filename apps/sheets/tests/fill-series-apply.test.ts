import { describe, expect, it } from 'vitest'

import { applyFillSeries } from '../src/renderer/fill-series-apply'
import type { UniverRuntime } from '../src/renderer/univer-state'

interface Write {
  row: number
  column: number
  rows: number
  columns: number
  matrix: unknown[][]
}

function fakeRuntime(
  cells: unknown[][],
  values: unknown[][],
): { runtime: UniverRuntime; writes: Write[] } {
  const writes: Write[] = []
  const range = {
    getRange: () => ({ startRow: 2, startColumn: 1 }),
    getCellDatas: () => cells,
    getValues: () => values,
    getHeight: () => cells.length,
    getWidth: () => cells[0]?.length ?? 0,
  }
  const sheet = {
    getMaxRows: () => 1000,
    getMaxColumns: () => 26,
    insertRows: () => sheet,
    insertColumns: () => sheet,
    getRange: (row: number, column: number, rows: number, columns: number) => ({
      setValues: (matrix: unknown[][]) => writes.push({ row, column, rows, columns, matrix }),
    }),
  }
  const runtime = {
    univerAPI: {
      getActiveWorkbook: () => ({ getActiveSheet: () => sheet, getActiveRange: () => range }),
    },
  } as unknown as UniverRuntime
  return { runtime, writes }
}

const t = (key: string): string => key

describe('applyFillSeries', () => {
  it('replaces formulas and rich text in the target cells and keeps the seed format', () => {
    const { runtime, writes } = fakeRuntime(
      [[{ v: 1, s: 'date-style' }], [{ f: '=A1*2', v: 2 }], [{ p: { body: {} }, v: 'rich' }]],
      [[1], [2], ['rich']],
    )
    const result = applyFillSeries(
      runtime,
      { direction: 'columns', type: 'linear', dateUnit: 'day', trend: false, step: 5, stop: null },
      t as never,
    )
    expect(result).toBeNull()
    expect(writes).toEqual([
      {
        row: 3,
        column: 1,
        rows: 2,
        columns: 1,
        matrix: [
          [{ v: 6, t: 2, f: null, si: null, p: null, s: 'date-style' }],
          [{ v: 11, t: 2, f: null, si: null, p: null, s: 'date-style' }],
        ],
      },
    ])
  })

  it('reads a formula seed from its computed value', () => {
    const { runtime, writes } = fakeRuntime([[{ f: '=1+9' }, null, null]], [[10, null, null]])
    applyFillSeries(
      runtime,
      { direction: 'rows', type: 'linear', dateUnit: 'day', trend: false, step: 1, stop: null },
      t as never,
    )
    expect(writes[0]?.matrix.map((row) => row.map((cell) => (cell as { v: number }).v))).toEqual([
      [11, 12],
    ])
  })
})
