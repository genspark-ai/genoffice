import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { readSheetRangeMapped } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

type RangeCall = { sessionId: string; sheetId: string; range: Record<string, number> }

const readWorkbookRange = vi.fn(async (call: RangeCall) => ({
  cells: [],
  rows: [],
  merges: [],
  hyperlinks: [],
  conditionalRules: [],
  dataValidations: [],
  indexedThroughRow: call.range.endRow,
}))

function state(): LazyWorkbookState {
  return {
    file: { sessionId: 'session-1', sheets: [] },
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
    editJournal: { cells: new Map(), structuralOps: new Map() },
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

const sheetMeta = {
  id: 's1',
  name: 'Sheet1',
  rowCount: 100_000,
  columnCount: 200,
} as unknown as LazyWorkbookState['file']['sheets'][number]

describe('readSheetRangeMapped: over-cap reads are row-batched', () => {
  beforeEach(() => {
    readWorkbookRange.mockClear()
    vi.stubGlobal('window', { desktopApi: { readWorkbookRange } })
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('splits a range that would exceed the sidecar cell budget', async () => {
    // 200 columns x 1500 rows = 300,000 cells; batches of floor(90_000/200)=450 rows.
    const result = await readSheetRangeMapped(
      state(),
      's1',
      { startRow: 0, endRow: 1_499, startColumn: 0, endColumn: 199 },
      sheetMeta,
    )
    expect(readWorkbookRange).toHaveBeenCalledTimes(4)
    const calls = readWorkbookRange.mock.calls.map(([call]) => call.range)
    expect(calls[0]).toMatchObject({ startRow: 0, endRow: 449 })
    expect(calls[1]).toMatchObject({ startRow: 450, endRow: 899 })
    expect(calls[2]).toMatchObject({ startRow: 900, endRow: 1_349 })
    expect(calls[3]).toMatchObject({ startRow: 1_350, endRow: 1_499 })
    expect(result?.fileEndRow).toBe(1_499)
    expect(result?.indexedThroughScreen).toBe(1_499)
  })

  it('stops early when indexing lags behind a batch', async () => {
    readWorkbookRange.mockImplementationOnce(async () => ({
      cells: [],
      rows: [],
      merges: [],
      hyperlinks: [],
      conditionalRules: [],
      dataValidations: [],
      // Indexing only reached row 20 of the requested 0..89 — later batches
      // cannot have data yet, so the read must not continue past this one.
      indexedThroughRow: 20,
    }))
    const result = await readSheetRangeMapped(
      state(),
      's1',
      { startRow: 0, endRow: 299, startColumn: 0, endColumn: 199 },
      sheetMeta,
    )
    expect(readWorkbookRange).toHaveBeenCalledTimes(1)
    expect(result?.indexedThroughScreen).toBe(20)
  })

  it('keeps small ranges on a single request', async () => {
    const result = await readSheetRangeMapped(
      state(),
      's1',
      { startRow: 5, endRow: 20, startColumn: 0, endColumn: 9 },
      sheetMeta,
    )
    expect(readWorkbookRange).toHaveBeenCalledTimes(1)
    expect(readWorkbookRange.mock.calls[0]?.[0]?.range).toMatchObject({
      startRow: 5,
      endRow: 20,
      startColumn: 0,
      endColumn: 9,
    })
    expect(result?.indexedThroughScreen).toBe(20)
  })
})
