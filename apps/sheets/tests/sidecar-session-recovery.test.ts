import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'

import { noteSidecarCrash, readSheetRangeMapped } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

/// Both the Rust sidecar and sheets-main's own session guard reject an unknown
/// session with this identical text (xlsx-engine lib.rs read_range;
/// sheets-main readWorkbookRange/saveWorkbook/closeWorkbook guards). The Save
/// swap and closeWorkbook raise it DELIBERATELY, on every ordinary save and
/// close, so the text alone cannot tell a crash from a teardown.
const SESSION_GONE = 'Unknown workbook session.'

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

/// Fails every read that carries the stale session id — what a sidecar replaced
/// after a crash does, and equally what the Save swap does to a read that
/// raced it. The error alone is the same in both cases.
const readWorkbookRange = vi.fn(async (call: RangeCall) => {
  if (call.sessionId !== FRESH) throw new Error(SESSION_GONE)
  return cellsFor(call)
})

/// The normal open path: what selectWorkbook re-opens a file through.
const reopenWorkbook = vi.fn(async (path: string) => ({
  sessionId: FRESH,
  path,
  sheets: [],
}))

/// The merge-source open, which recovery must never reach for: it opens a
/// second, read-only input session (re-running the csv/xls import) that
/// nothing would own or close.
const openWorkbooksForMerge = vi.fn(async (paths: string[]) => [
  { sessionId: 'merge-source', path: paths[0], sheets: [] },
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
    reopenWorkbook.mockClear()
    openWorkbooksForMerge.mockClear()
    vi.stubGlobal('window', {
      desktopApi: { readWorkbookRange, reopenWorkbook, openWorkbooksForMerge },
    })
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('re-opens through the normal open path and serves the read on the new session', async () => {
    // The positive crash signal from the main process (sidecar process death).
    noteSidecarCrash()
    const lazy = state()

    const result = await readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)

    expect(readWorkbookRange.mock.calls[0]![0].sessionId).toBe(STALE)
    // The NORMAL open path, not the merge-source open: a crash re-open is
    // re-opening this workbook, not opening a second read-only input.
    expect(reopenWorkbook).toHaveBeenCalledWith('/books/report.xlsx')
    expect(openWorkbooksForMerge).not.toHaveBeenCalled()
    expect(readWorkbookRange.mock.calls.at(-1)![0].sessionId).toBe(FRESH)
    expect(result?.screen.cells).toEqual([{ row: 0, column: 0, value: 'ok' }])
    expect(lazy.file.sessionId).toBe(FRESH)
  })

  it('drops the streaming memos the dead session had already satisfied', async () => {
    noteSidecarCrash()
    const lazy = state()
    await readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)
    // Left in place these claim the window is loaded, so the next viewport
    // load returns early and the sheet stays blank.
    expect(lazy.loadedRanges.size).toBe(0)
    expect(lazy.loadingKeys.size).toBe(0)
    expect(lazy.frozenStripKeys.size).toBe(0)
  })

  it('keeps the session edits made before the crash', async () => {
    noteSidecarCrash()
    const journal = new Map([['0:0', { hasValue: true, value: 'typed' }]])
    const lazy = state('/books/report.xlsx', journal)

    await readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)
    // The journal is keyed by sheet, not by session: unsaved edits outlive the
    // crash and must not be dropped by the re-open.
    expect(lazy.editJournal.cells).toBe(journal)
  })

  it('propagates a read failure that is not a lost session', async () => {
    // No crash signal: a transient read error must not cost the user their
    // session, whatever its wording.
    readWorkbookRange.mockRejectedValueOnce(new Error('worksheet part unreadable'))
    await expect(readSheetRangeMapped(state(), 's1', RANGE, sheetMeta)).rejects.toThrow(
      'worksheet part unreadable',
    )
    expect(reopenWorkbook).not.toHaveBeenCalled()
  })

  it('recovers a genuine crash even when the read fails for another reason', async () => {
    // A crash kills the whole session, so the first rejection after it can be
    // anything the dying process managed to say. The positive signal — not
    // the message — is what drives recovery.
    noteSidecarCrash()
    readWorkbookRange.mockRejectedValueOnce(new Error('worksheet part unreadable'))
    const lazy = state()

    const result = await readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)

    expect(reopenWorkbook).toHaveBeenCalledWith('/books/report.xlsx')
    expect(result?.screen.cells).toEqual([{ row: 0, column: 0, value: 'ok' }])
    expect(lazy.file.sessionId).toBe(FRESH)
  })

  it('gives up rather than looping when there is no file to re-open', async () => {
    // An unsaved new workbook has no path: only a manual open can recover it.
    noteSidecarCrash()
    const lazy = state(null)
    await expect(readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)).rejects.toThrow(SESSION_GONE)
    expect(reopenWorkbook).not.toHaveBeenCalled()
  })

  it('re-opens only once per workbook, so a second crash cannot loop', async () => {
    const lazy = state()
    noteSidecarCrash()
    await readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)
    expect(reopenWorkbook).toHaveBeenCalledTimes(1)

    // A new crash after the recovered session: reported, not retried forever.
    readWorkbookRange.mockImplementation(async () => {
      throw new Error(SESSION_GONE)
    })
    noteSidecarCrash()
    await expect(readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)).rejects.toThrow(SESSION_GONE)
    expect(reopenWorkbook).toHaveBeenCalledTimes(1)
  })
})

/// The defect the maintainer reported: the guard's message is shared by a real
/// crash and by the sessions the app tears down on purpose, so matching on it
/// re-opened the file during ordinary saves.
describe('a deliberate teardown is not a crash', () => {
  beforeEach(() => {
    readWorkbookRange.mockClear()
    reopenWorkbook.mockClear()
    openWorkbooksForMerge.mockClear()
    vi.stubGlobal('window', {
      desktopApi: { readWorkbookRange, reopenWorkbook, openWorkbooksForMerge },
    })
  })

  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('does not recover when a read races the Save swap', async () => {
    // The Save swap deletes the old session and opens a new one; a read still
    // carrying the old id is rejected with the SAME message a crash produces.
    // No crash signal: this must surface as an ordinary read error.
    const lazy = state()
    await expect(readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)).rejects.toThrow(SESSION_GONE)

    // Re-opening here is what orphaned a merge-source session: the swap's
    // replacement session then overwrote state.file.sessionId, leaving the
    // merge session alive and unowned, on every save that raced a read.
    expect(reopenWorkbook).not.toHaveBeenCalled()
    expect(openWorkbooksForMerge).not.toHaveBeenCalled()
    expect(lazy.file.sessionId).toBe(STALE)
  })

  it('does not recover when the post-closeWorkbook guard rejects a read', async () => {
    const lazy = state()
    await expect(readSheetRangeMapped(lazy, 's1', RANGE, sheetMeta)).rejects.toThrow(SESSION_GONE)

    expect(reopenWorkbook).not.toHaveBeenCalled()
    expect(openWorkbooksForMerge).not.toHaveBeenCalled()
    expect(lazy.file.sessionId).toBe(STALE)
  })

  it('does not treat a csv import failure as a lost session', async () => {
    // A different guard message entirely: the crash gate is a positive signal,
    // not a scan for session-shaped text.
    readWorkbookRange.mockRejectedValue(new Error('Merge source not found.'))
    await expect(readSheetRangeMapped(state(), 's1', RANGE, sheetMeta)).rejects.toThrow(
      'Merge source not found.',
    )
    expect(reopenWorkbook).not.toHaveBeenCalled()
  })
})
