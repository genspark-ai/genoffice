import { describe, expect, it } from 'vitest'

import { installCellImages } from '../src/renderer/WorkbookVisuals'
import type { createUniver } from '../src/renderer/create-univer'

describe('in-cell picture placement', () => {
  it('uses each host cell or its merged range, including repeated media', () => {
    const merged = { startRow: 3, endRow: 4, startColumn: 0, endColumn: 2 }
    const placements: unknown[] = []
    const worksheet = {
      getMaxRows: () => 10,
      getMaxColumns: () => 10,
      getCellMergeData: (row: number, column: number) =>
        row === 3 && column === 0 ? merged : undefined,
      getRange: (row: number, column: number, rows: number, columns: number) => ({
        startRow: row,
        endRow: row + rows - 1,
        startColumn: column,
        endColumn: column + columns - 1,
      }),
      addFloatDomToRange: (range: unknown) => {
        placements.push(range)
        return { dispose() {} }
      },
    }
    const runtime = {
      univerAPI: {
        getActiveWorkbook: () => ({ getSheetBySheetId: () => worksheet }),
        registerComponent: () => ({ dispose() {} }),
      },
    } as unknown as ReturnType<typeof createUniver>

    installCellImages(
      runtime,
      'session',
      [
        { id: 'image-1', row: 1, column: 1 },
        { id: 'image-2', row: 1, column: 2 },
        { id: 'image-3', row: 3, column: 0 },
      ],
      'sheet-1',
    )
    expect(placements).toEqual([
      { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1 },
      { startRow: 1, endRow: 1, startColumn: 2, endColumn: 2 },
      merged,
    ])
  })
})
