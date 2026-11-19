import { describe, expect, it } from 'vitest'

import { inferContinuousRegion } from '../src/renderer/table-actions'
import type { UniverWorksheet } from '../src/renderer/univer-state'

type Cell = string | number | null

function worksheet(grid: Cell[][]): UniverWorksheet {
  return {
    getLastRow: () => grid.length - 1,
    getLastColumn: () => Math.max(...grid.map((row) => row.length)) - 1,
    getRange: (row: number, column: number) => ({
      getValue: () => grid[row]?.[column] ?? null,
      getDisplayValue: () => {
        const value = grid[row]?.[column]
        return value == null ? '' : String(value)
      },
    }),
  } as unknown as UniverWorksheet
}

const _ = null

describe('inferContinuousRegion', () => {
  const block = worksheet([
    ['h1', 'h2', 'h3', 'h4'],
    [1, 2, 3, 4],
    [5, _, 7, 8],
    [_, _, _, _],
    ['x', 'y'],
  ])

  it('expands a single cell to the data block bounded by empty rows and columns', () => {
    expect(inferContinuousRegion(block, 1, 1)).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 2,
      endColumn: 3,
    })
  })

  it('reaches the block from an edge cell and across an interior blank', () => {
    expect(inferContinuousRegion(block, 2, 0)).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 2,
      endColumn: 3,
    })
  })

  it('returns null for an empty anchor or a header-only island', () => {
    expect(inferContinuousRegion(block, 2, 1)).toBeNull()
    expect(inferContinuousRegion(block, 4, 0)).toBeNull()
  })

  it('grows columns from a ragged row and keeps zero values, not empty strings', () => {
    const ragged = worksheet([
      ['h1', 'h2'],
      [1, 2, 'note'],
      [3, 4],
    ])
    expect(inferContinuousRegion(ragged, 0, 0)).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: 2,
      endColumn: 2,
    })
    const offset = worksheet([[], [_, _, 'h', 'k'], [_, _, 0, ''], [_, _, 0, 'z']])
    expect(inferContinuousRegion(offset, 3, 3)).toEqual({
      startRow: 1,
      startColumn: 2,
      endRow: 3,
      endColumn: 3,
    })
  })

  it('walks a block taller than ten thousand rows to its real end', () => {
    const rows = 12_000
    const tall = worksheet(Array.from({ length: rows }, (_row, i) => [`a${i}`, i]))
    expect(inferContinuousRegion(tall, 0, 0)).toEqual({
      startRow: 0,
      startColumn: 0,
      endRow: rows - 1,
      endColumn: 1,
    })
  })
})
