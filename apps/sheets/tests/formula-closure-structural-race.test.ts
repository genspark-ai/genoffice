import { afterEach, describe, expect, it } from 'vitest'

import { activateFormulaClosure } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

interface PatchCall {
  sheetId: string
  startRow: number
  startColumn: number
}

const sheetMeta = (id: string, name: string) => ({
  id,
  name,
  rowCount: 10,
  columnCount: 5,
  tables: [],
  pivotTables: [],
  freeze: null,
  columnWidths: [],
})

function closureState(): LazyWorkbookState {
  return {
    formulaMode: false,
    file: {
      sessionId: 'session-1',
      styles: [],
      sheets: [sheetMeta('sheet-1', 'Alpha'), sheetMeta('sheet-2', 'Beta')],
    },
    editJournal: { structuralOps: new Map(), cells: new Map() },
    closure: { status: 'idle', pinned: new Map() },
    formulaText: new Map(),
    cachedFormulaValues: new Map(),
    rowColStyleKeys: new Map(),
    hiddenFileRows: new Map(),
    hiddenRowsCoveredThrough: new Map(),
  } as unknown as LazyWorkbookState
}

function stubWorksheet(id: string, patches: PatchCall[]): unknown {
  return {
    getSheetId: () => id,
    getRange: (row: number, column: number) => ({
      setValues: () => {
        patches.push({ sheetId: id, startRow: row, startColumn: column })
      },
    }),
  }
}

function stubRuntime(patches: PatchCall[]): unknown {
  const worksheets = new Map([
    ['sheet-1', stubWorksheet('sheet-1', patches)],
    ['sheet-2', stubWorksheet('sheet-2', patches)],
  ])
  return {
    univerAPI: {
      getActiveWorkbook: () => ({
        getSheetBySheetId: (id: string) => worksheets.get(id) ?? null,
      }),
    },
  }
}

/// Each sheet holds one formula at D1 referencing A1, so the closure fetches
/// exactly one band per sheet — two sequential sidecar reads in total.
function stubDesktopApi(onRangeRead: (call: number) => void): void {
  let rangeReads = 0
  ;(globalThis as { window?: unknown }).window = {
    desktopApi: {
      readWorkbookFormulas: async () => ({
        truncated: false,
        indexingComplete: true,
        cells: [{ row: 0, column: 3, formula: 'A1*2' }],
      }),
      readWorkbookRange: async () => {
        rangeReads += 1
        onRangeRead(rangeReads)
        return {
          cells: [
            { row: 0, column: 0, value: 2 },
            { row: 0, column: 3, formula: 'A1*2', value: 4 },
          ],
          rows: [],
          merges: [],
          hyperlinks: [],
          conditionalRules: [],
          dataValidations: [],
          autoFilter: null,
          sheetProtection: null,
          indexedThroughRow: 10,
          indexingComplete: true,
        }
      },
    },
  }
}

describe('activateFormulaClosure structural-edit race', () => {
  afterEach(() => {
    delete (globalThis as { window?: unknown }).window
  })

  it('installs and pins the closure when no structural edit interferes', async () => {
    const state = closureState()
    const patches: PatchCall[] = []
    stubDesktopApi(() => undefined)

    await activateFormulaClosure(stubRuntime(patches) as never, { current: state }, () => undefined)

    expect(state.closure.status).toBe('active')
    expect(patches).toHaveLength(2)
    expect(state.closure.pinned.get('sheet-1')?.get('0:3')).toEqual({ f: 'A1*2', v: 4 })
    expect(state.closure.pinned.get('sheet-2')?.get('0:0')).toEqual({ v: 2 })
  })

  it('gives up and drops all pins when a structural edit lands during an install read', async () => {
    const state = closureState()
    const patches: PatchCall[] = []
    // The first sheet installs cleanly; while the second sheet's read is in
    // flight the user inserts rows. Installing that file-space result (or
    // keeping the already-stored pins for later re-application) would land
    // cells at shifted screen positions.
    stubDesktopApi((call) => {
      if (call === 2) {
        state.editJournal.structuralOps.set('sheet-2', [
          { kind: 'insert-rows', index: 0, count: 3 },
        ])
      }
    })

    await activateFormulaClosure(stubRuntime(patches) as never, { current: state }, () => undefined)

    expect(state.closure.status).toBe('unavailable')
    expect(patches).toHaveLength(1)
    expect(state.closure.pinned.size).toBe(0)
  })
})
