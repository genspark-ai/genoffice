import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { readCopySourceDirect } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

type RangeCall = { sessionId: string; sheetId: string; range: Record<string, number> }

const readWorkbookRange = vi.fn(async (call: RangeCall): Promise<Record<string, unknown>> => ({
  cells: [{ row: 0, column: 0, value: 7, formula: '=B1*2', styleIndex: 0 }],
  rows: [],
  merges: [],
  hyperlinks: [],
  conditionalRules: [],
  dataValidations: [],
  indexedThroughRow: call.range.endRow,
  indexingComplete: true,
}))

function state(): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      name: 'book.xlsx',
      sheets: [{ id: 's1', name: 'Sheet1', rowCount: 100, columnCount: 10 }],
      styles: [{ bold: true }],
    },
    generation: 1,
    loadedRanges: new Map(),
    loadingKeys: new Map(),
    retryTimers: new Map(),
    appliedMerges: new Map(),
    appliedRowKeys: new Map(),
    sheetProtections: new Map(),
    sheetPageBreaks: new Map(),
    sheetProtectedRanges: new Map(),
    uninstalledDefinedNames: new Set(),
    appliedCfSheets: new Set(),
    appliedFilterSheets: new Set(),
    appliedDvSheets: new Set(),
    decorationsPendingSheets: new Set(),
    hyperlinkTargets: new Map(),
    frozenStripKeys: new Map(),
    filterOrigins: new Map(),
    showFormulaSheets: new Set(),
    formulaMode: false,
    editJournal: { cells: new Map(), structuralOps: new Map(), bulkConstantFills: new Map() },
    flags: { preloadComplete: false },
    closure: { status: 'idle', pinned: new Map() },
    formulaText: new Map(),
    cachedFormulaValues: new Map(),
    pivotDefinitions: new Map(),
    outline: new Map(),
    recalc: {
      timer: null,
      generation: 0,
      failures: 0,
      formulaCells: new Map(),
      overlay: new Map(),
    },
  } as unknown as LazyWorkbookState
}

const bounds = { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }
const noMessage = () => {}

describe('readCopySourceDirect: journal semantics', () => {
  beforeEach(() => {
    readWorkbookRange.mockClear()
    vi.stubGlobal('window', { desktopApi: { readWorkbookRange } })
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('carries file formula text and the resolved xf on a clean sheet', async () => {
    const result = await readCopySourceDirect({ current: state() }, 's1', bounds, noMessage)
    expect(result?.[0]?.[0]).toMatchObject({ v: 7, fileFormula: '=B1*2' })
    expect(result?.[0]?.[0]?.s).toMatchObject({ bl: 1 })
  })

  it('refuses file-space formula text once structural edits shifted coordinates', async () => {
    const workbook = state()
    workbook.editJournal.structuralOps.set('s1', [{ kind: 'insert-rows', index: 50, count: 1 }])
    const result = await readCopySourceDirect({ current: workbook }, 's1', bounds, noMessage)
    expect(result?.[0]?.[0]?.v).toBe(7)
    expect(result?.[0]?.[0]?.fileFormula).toBeNull()
  })

  it('honors a journaled style reset over the file xf', async () => {
    const workbook = state()
    workbook.editJournal.cells.set(
      's1',
      new Map([['0:0', { row: 0, column: 0, hasValue: false, value: null, styleReset: true }]]),
    )
    const result = await readCopySourceDirect({ current: workbook }, 's1', bounds, noMessage)
    expect(result?.[0]?.[0]?.v).toBe(7)
    expect(result?.[0]?.[0]?.s).toBeNull()
  })
})
