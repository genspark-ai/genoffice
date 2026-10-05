import { afterEach, describe, expect, it } from 'vitest'

import { precheckStructuralDeletes } from '../src/renderer/op-executor'
import { applyAiTableRowDelete } from '../src/renderer/workbook-ops'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

/// An AI delete_table_row is a whole-sheet row delete underneath, so the
/// batch precheck translates it to the same span the apply path deletes. This
/// drives both paths and compares them, so the precheck cannot drift onto a
/// different row base (which either false-rejects a legal batch or misses the
/// row that ⌘S will then refuse to save).

// 0-based header row of the session-added table, so the data rows the op
// addresses are 3..9 — deliberately not row 0.
const HEADER = 2

function state(): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      sheets: [{ id: 'sh1', name: 'Data', rowCount: 100, columnCount: 26, pivotRanges: [] }],
      visuals: [],
    },
    editJournal: {
      cells: new Map(),
      structuralOps: new Map(),
      sheets: { added: new Set(), removed: new Set() },
      tableAdds: [
        {
          sheetId: 'sh1',
          name: 'Sales',
          area: { startRow: HEADER, startColumn: 0, endRow: 8, endColumn: 3 },
          columnNames: ['A', 'B', 'C', 'D'],
        },
      ],
    },
  } as unknown as LazyWorkbookState
}

const workbook = {
  getSheets: () => [{ getSheetId: () => 'sh1', getSheetName: () => 'Data' }],
} as unknown as Parameters<typeof precheckStructuralDeletes>[1]

function stubFormulaOn(sheetRow: number): void {
  ;(globalThis as { window?: unknown }).window = {
    desktopApi: {
      readWorkbookFormulas: async () => ({
        cells: [{ row: 20, column: 5, formula: `=SUM($A$${sheetRow}:$A$${sheetRow})` }],
        truncated: false,
        indexingComplete: true,
      }),
    },
  }
}

const op = (row: number, count = 1) =>
  ({ op: 'delete_table_row', sheetId: 'sh1', tableName: 'Sales', row, count }) as never

/// The 0-based sheet row applyAiTableRowDelete hands to Univer's deleteRows.
function applyRow(row: number, count = 1): number {
  const calls: number[] = []
  const runtime = {
    univerAPI: {
      getActiveWorkbook: () => ({
        getSheetBySheetId: () => ({
          deleteRows: (rowIndex: number) => calls.push(rowIndex),
        }),
      }),
    },
  } as never
  applyAiTableRowDelete(runtime, state(), op(row, count))
  return calls[0] as number
}

async function rejects(row: number, count: number, targetSheetRow: number): Promise<boolean> {
  stubFormulaOn(targetSheetRow)
  return (await precheckStructuralDeletes(state(), workbook, [op(row, count)])) !== null
}

afterEach(() => {
  delete (globalThis as { window?: unknown }).window
})

describe('delete_table_row precheck row base', () => {
  for (const row of [1, 2, 3]) {
    it(`checks the rows op.row=${row} actually removes`, async () => {
      const deleted = applyRow(row)
      const firstSheetRow = deleted + 1
      // Every row the apply path removes must be rejected, and the rows
      // around it must pass — for op.row beyond 0 too, not just the first.
      expect(await rejects(row, 1, firstSheetRow)).toBe(true)
      expect(await rejects(row, 1, firstSheetRow - 1)).toBe(false)
      expect(await rejects(row, 1, firstSheetRow + 1)).toBe(false)
      // A multi-row delete rejects exactly its own span.
      expect(await rejects(row, 2, firstSheetRow + 1)).toBe(true)
      expect(await rejects(row, 2, firstSheetRow + 2)).toBe(false)
    })
  }
})
