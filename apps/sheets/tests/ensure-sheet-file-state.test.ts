import { afterEach, describe, expect, it, vi } from 'vitest'

import { ensureSheetFileState } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

function makeState(): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      sheets: [{ id: 'sheet-2', name: 'Data 2024', rowCount: 2, columnCount: 2 }],
    },
    editJournal: { structuralOps: new Map(), sheets: { added: new Map(), removed: new Set() } },
    sheetProtections: new Map(),
    sheetPageBreaks: new Map(),
    sheetFilePageSetups: new Map(),
    sheetProtectedRanges: new Map(),
  } as unknown as LazyWorkbookState
}

function rangeResult(indexingComplete: boolean) {
  return {
    cells: [],
    rows: [],
    indexingComplete,
    indexedThroughRow: 1,
    sheetProtection: indexingComplete ? { protected: true, hasPassword: false } : null,
    rowBreaks: [],
    colBreaks: [],
    pageSetup: null,
    protectedRanges: [],
  }
}

afterEach(() => {
  vi.unstubAllGlobals()
  vi.useRealTimers()
})

describe('ensureSheetFileState', () => {
  it('polls an unvisited sheet until indexing completes, then records its protection', async () => {
    vi.useFakeTimers()
    const readWorkbookRange = vi
      .fn()
      .mockResolvedValueOnce(rangeResult(false))
      .mockResolvedValueOnce(rangeResult(true))
    vi.stubGlobal('window', { desktopApi: { readWorkbookRange } })
    const state = makeState()
    const pending = ensureSheetFileState(state, 'sheet-2', { current: state })
    await vi.advanceTimersByTimeAsync(200)
    expect(await pending).toBe(true)
    expect(readWorkbookRange).toHaveBeenCalledTimes(2)
    expect(readWorkbookRange.mock.calls[0]![0]).toMatchObject({
      sessionId: 'session-1',
      sheetId: 'sheet-2',
    })
    expect(state.sheetProtections.get('sheet-2')).toMatchObject({ protected: true })
  })

  it('skips sheets whose state is already known or that the session added', async () => {
    const readWorkbookRange = vi.fn()
    vi.stubGlobal('window', { desktopApi: { readWorkbookRange } })
    const state = makeState()
    state.sheetProtections.set('sheet-2', {
      protected: false,
      hasPassword: false,
      allow: {} as never,
      password: null,
    })
    expect(await ensureSheetFileState(state, 'sheet-2', { current: state })).toBe(true)
    state.editJournal.sheets.added.set('sheet-9', { name: 'Scratch' })
    expect(await ensureSheetFileState(state, 'sheet-9', { current: state })).toBe(true)
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })

  it('gives up when the workbook was replaced meanwhile', async () => {
    const readWorkbookRange = vi.fn().mockResolvedValue(rangeResult(false))
    vi.stubGlobal('window', { desktopApi: { readWorkbookRange } })
    const state = makeState()
    expect(await ensureSheetFileState(state, 'sheet-2', { current: null })).toBe(false)
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })
})
