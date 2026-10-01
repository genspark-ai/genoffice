import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { readSheetRangeMapped } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

/// The sidecar rejects a session its replacement has never heard of with this
/// exact message (see xlsx-engine lib.rs read_range, and the same wording
/// sheets-main's own session guard throws).
const SESSION_LOST = 'Unknown workbook session.'

const STALE = 'stale-session'
const FRESH = 'fresh-session'

type RangeCall = { sessionId: string; sheetId: string; range: Record<string, number> }

function cellsFor(call: RangeCall) {
  return {
    cells: [{ row: call.range.startRow, column: 0, value: 'ok' }],
    rows: [],
    merges: [],
    hyperlinks: [],
    conditionalRules: [],
    dataValidations: [],
    indexedThroughRow: call.range.endRow,
  }
}

/// Fails every read that carries the stale session id, exactly as a sidecar
/// replaced after a crash does: the new process has no such session.
const readWorkbookRange = vi.fn(async (call: RangeCall) => {
  if (call.sessionId !== FRESH) throw new Error(SESSION_LOST)
  return cellsFor(call)
})

const openWorkbooksForMerge = vi.fn(async (paths: string[]) => [
  { sessionId: FRESH, path: paths[0], sheets: [] },
])

function state(
  path: string | null = '/books/report.xlsx',
  journalCells?: Map<string, { hasValue: boolean; value: string }>,
): LazyWorkbookState {
  return {
    file: { sessionId: STALE, path: path ?? undefined, sheets: [] },
    generation: 1,
    loadedRanges: new Map([['s1', { startRow: 0, endRow: 49, startColumn: 0, endColumn: 9 }]]),
    loadingKeys: new Map([['s1', 'key-1']]),
    retryTimers: new Map(),
    frozenStripKeys: new Map([['s1', 's1:0:49:0:9']]),
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
    filterOrigins: new Map(),
    showFormulaSheets: new Set(),
    formulaMode: false,
    editJournal: { cells: journalCells ?? new Map(), structuralOps: new Map() },
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

const RANGE = { startRow: 0, endRow: 49, startColumn: 0, endColumn: 9 }

describe('a sidecar crash no longer strands the grid on a dead session', () => {
  beforeEach(() => {
    readWorkbookRange.mockClear()
    openWorkbooksForMerge.mockClear()
    vi.stubGlobal('window', { desktopApi: { readWorkbookRange, openWorkbooksForMerge } })
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('re-opens the file and serves the read through the new session', async () => {
    const lazy = state()
    const result = await readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)

    // The stale id was tried, the workbook re-opened for a live session, and
    // the read retried against it — instead of the grid staying empty forever.
    expect(readWorkbookRange.mock.calls[0]![0].sessionId).toBe(STALE)
    expect(openWorkbooksForMerge).toHaveBeenCalledWith(['/books/report.xlsx'])
    expect(readWorkbookRange.mock.calls.at(-1)![0].sessionId).toBe(FRESH)
    expect(result?.screen.cells).toEqual([{ row: 0, column: 0, value: 'ok' }])
    expect(lazy.file.sessionId).toBe(FRESH)
  })

  it('drops the streaming memos the dead session had already satisfied', async () => {
    const lazy = state()
    await readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)
    // Left in place these claim the window is loaded, so the next viewport
    // load returns early and the sheet stays blank.
    expect(lazy.loadedRanges.size).toBe(0)
    expect(lazy.loadingKeys.size).toBe(0)
    expect(lazy.frozenStripKeys.size).toBe(0)
  })

  it('keeps the session edits made before the crash', async () => {
    const journal = new Map([['0:0', { hasValue: true, value: 'typed' }]])
    const lazy = state('/books/report.xlsx', journal)

    await readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)
    // The journal is keyed by sheet, not by session: unsaved edits outlive the
    // crash and must not be dropped by the re-open.
    expect(lazy.editJournal.cells).toBe(journal)
  })

  it('propagates a read failure that is not a lost session', async () => {
    readWorkbookRange.mockRejectedValueOnce(new Error('worksheet part unreadable'))
    await expect(readSheetRangeMapped(state(), 's1', RANGE, sheetMeta)).rejects.toThrow(
      'worksheet part unreadable',
    )
    // A transient read error must not cost the user their session.
    expect(openWorkbooksForMerge).not.toHaveBeenCalled()
  })

  it('gives up rather than looping when there is no file to re-open', async () => {
    // An unsaved new workbook has no path: only a manual open can recover it.
    const lazy = state(null)
    await expect(readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)).rejects.toThrow(SESSION_LOST)
    expect(openWorkbooksForMerge).not.toHaveBeenCalled()
  })

  it('re-opens only once per workbook, so a second loss cannot loop', async () => {
    const lazy = state()
    await readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)
    expect(openWorkbooksForMerge).toHaveBeenCalledTimes(1)

    // A new crash after the recovered session: reported, not retried forever.
    readWorkbookRange.mockImplementation(async () => {
      throw new Error(SESSION_LOST)
    })
    await expect(readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)).rejects.toThrow(SESSION_LOST)
    expect(openWorkbooksForMerge).toHaveBeenCalledTimes(1)
  })
})
