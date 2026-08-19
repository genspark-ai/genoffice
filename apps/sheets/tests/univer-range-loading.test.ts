import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  applyRangeInLoadedChunks,
  ensureLazyRangeLoaded,
  lazyCellEditable,
  lazyRangeEditable,
  loadWorkbookSkeleton,
  nextIndexWaitStall,
  normalizeVisibleRange,
} from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

/// loadWorkbookSkeleton clears the unit's undo history through the injector.
const undoStub = {
  __getInjector: () => ({ get: () => ({ clearUndoRedo: () => undefined }) }),
}

describe('normalizeVisibleRange', () => {
  it('uses the initial viewport when Univer has no scroll range yet', () => {
    expect(normalizeVisibleRange(null, 14_516, 16)).toEqual({
      startRow: 0,
      endRow: 79,
      startColumn: 0,
      endColumn: 15,
    })
  })

  it('falls back when a stale viewport is outside the replacement sheet', () => {
    expect(
      normalizeVisibleRange(
        {
          startRow: 20_000,
          endRow: 20_050,
          startColumn: 20,
          endColumn: 25,
        },
        14_516,
        16,
      ),
    ).toEqual({
      startRow: 0,
      endRow: 79,
      startColumn: 0,
      endColumn: 15,
    })
  })

  it('clamps a partially out-of-bounds viewport', () => {
    expect(
      normalizeVisibleRange(
        {
          startRow: -5,
          endRow: 200,
          startColumn: -2,
          endColumn: 30,
        },
        100,
        16,
      ),
    ).toEqual({
      startRow: 0,
      endRow: 99,
      startColumn: 0,
      endColumn: 15,
    })
  })
})

describe('loadWorkbookSkeleton', () => {
  it('shows a full starter grid for a blank A1-only workbook', () => {
    const created: Array<{
      sheets: Record<string, { rowCount: number; columnCount: number }>
    }> = []
    const runtime = {
      univer: undoStub,
      univerAPI: {
        getActiveWorkbook: () => null,
        disposeUnit: () => undefined,
        createWorkbook: (config: (typeof created)[number]) => {
          created.push(config)
          return { getSheetBySheetId: () => null, setActiveSheet: () => undefined }
        },
      },
    }
    const file = {
      sha256: 'blank',
      name: 'Book1.xlsx',
      visuals: [],
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          rowCount: 1,
          columnCount: 1,
          hidden: false,
          showGridLines: true,
          tabColor: null,
          defaultRowHeight: null,
          defaultColumnWidth: null,
          freeze: null,
          columnWidths: [],
        },
      ],
    }

    loadWorkbookSkeleton(runtime as never, file as never)

    expect(created[0]?.sheets['sheet-1']).toMatchObject({
      rowCount: 1000,
      columnCount: 26,
    })
  })

  it('preserves a file extent larger than the starter grid', () => {
    const created: Array<{
      sheets: Record<string, { rowCount: number; columnCount: number }>
    }> = []
    const runtime = {
      univer: undoStub,
      univerAPI: {
        getActiveWorkbook: () => null,
        disposeUnit: () => undefined,
        createWorkbook: (config: (typeof created)[number]) => {
          created.push(config)
          return { getSheetBySheetId: () => null, setActiveSheet: () => undefined }
        },
      },
    }
    const file = {
      sha256: 'large',
      name: 'Large.xlsx',
      visuals: [],
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          rowCount: 1004,
          columnCount: 13,
          hidden: false,
          showGridLines: true,
          tabColor: null,
          defaultRowHeight: null,
          defaultColumnWidth: null,
          freeze: null,
          columnWidths: [],
        },
      ],
    }

    loadWorkbookSkeleton(runtime as never, file as never)

    expect(created[0]?.sheets['sheet-1']).toMatchObject({
      rowCount: 1004,
      columnCount: 26,
    })
  })
})

interface Range {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
}

function streamedState(options: {
  loaded?: Range
  ops?: Array<{ kind: string; index: number; count: number }>
  preloadComplete?: boolean
  journalCells?: Array<{ row: number; column: number; value: string | number | null }>
  rows?: number
}): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      styles: [],
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          rowCount: options.rows ?? 10,
          columnCount: 5,
          tables: [],
          pivotTables: [],
          freeze: null,
        },
      ],
    },
    loadedRanges: new Map(options.loaded ? [['sheet-1', options.loaded]] : []),
    loadingKeys: new Map(),
    retryTimers: new Map(),
    decorationsPendingSheets: new Set(),
    editJournal: {
      structuralOps: new Map(options.ops ? [['sheet-1', options.ops]] : []),
      cells: new Map(
        options.journalCells
          ? [
              [
                'sheet-1',
                new Map(
                  options.journalCells.map((cell) => [
                    `${cell.row}:${cell.column}`,
                    { ...cell, hasValue: true },
                  ]),
                ),
              ],
            ]
          : [],
      ),
    },
    closure: { pinned: new Map() },
    recalc: { overlay: new Map() },
    flags: { preloadComplete: options.preloadComplete ?? false },
  } as unknown as LazyWorkbookState
}

const LOADED_TOP: Range = { startRow: 0, endRow: 4, startColumn: 0, endColumn: 4 }

describe('lazyCellEditable', () => {
  it('allows loaded cells and truly-beyond-data cells, blocks unstreamed ones', () => {
    const state = streamedState({ loaded: LOADED_TOP })
    expect(lazyCellEditable(state, 'sheet-1', 2, 2)).toBe(true)
    expect(lazyCellEditable(state, 'sheet-1', 7, 0)).toBe(false)
    expect(lazyCellEditable(state, 'sheet-1', 12, 0)).toBe(true)
  })

  it('blocks the shifted unstreamed tail after insert_rows', () => {
    const state = streamedState({
      loaded: LOADED_TOP,
      ops: [{ kind: 'insert-rows', index: 2, count: 3 }],
    })
    // Screen rows 10-12 now hold file rows 7-9, which never streamed in.
    expect(lazyCellEditable(state, 'sheet-1', 10, 0)).toBe(false)
    expect(lazyCellEditable(state, 'sheet-1', 12, 0)).toBe(false)
    expect(lazyCellEditable(state, 'sheet-1', 13, 0)).toBe(true)
  })

  it('allows rows inserted this session (journal-owned, nothing streams in)', () => {
    const state = streamedState({ ops: [{ kind: 'insert-rows', index: 6, count: 3 }] })
    expect(lazyCellEditable(state, 'sheet-1', 6, 0)).toBe(true)
    expect(lazyCellEditable(state, 'sheet-1', 8, 0)).toBe(true)
    expect(lazyCellEditable(state, 'sheet-1', 9, 0)).toBe(false)
  })

  it('shrinks the editable-beyond-data bound after delete_rows', () => {
    const state = streamedState({
      loaded: LOADED_TOP,
      ops: [{ kind: 'remove-rows', index: 0, count: 3 }],
    })
    expect(lazyCellEditable(state, 'sheet-1', 7, 0)).toBe(true)
    expect(lazyCellEditable(state, 'sheet-1', 5, 0)).toBe(false)
    expect(lazyCellEditable(state, 'sheet-1', 3, 0)).toBe(true)
  })

  it('handles column inserts the same way', () => {
    const state = streamedState({
      loaded: LOADED_TOP,
      ops: [{ kind: 'insert-cols', index: 1, count: 2 }],
    })
    expect(lazyCellEditable(state, 'sheet-1', 0, 5)).toBe(false)
    expect(lazyCellEditable(state, 'sheet-1', 0, 7)).toBe(true)
    expect(lazyCellEditable(state, 'sheet-1', 0, 1)).toBe(true)
  })

  it('always allows edits once the workbook is fully loaded', () => {
    const state = streamedState({ preloadComplete: true })
    expect(lazyCellEditable(state, 'sheet-1', 7, 0)).toBe(true)
  })
})

describe('ensureLazyRangeLoaded', () => {
  const worksheet = { getSheetId: () => 'sheet-1' }

  function ensure(state: LazyWorkbookState, range: Range): Promise<boolean> {
    return ensureLazyRangeLoaded(
      { univerAPI: { getActiveWorkbook: () => null } } as never,
      { current: state },
      worksheet as never,
      range,
      (message) => {
        throw new Error(message)
      },
    )
  }

  it('accepts the shifted tail after insert_rows once loaded', async () => {
    const state = streamedState({
      loaded: { startRow: 0, endRow: 12, startColumn: 0, endColumn: 4 },
      ops: [{ kind: 'insert-rows', index: 2, count: 3 }],
    })
    await expect(
      ensure(state, { startRow: 10, endRow: 12, startColumn: 0, endColumn: 4 }),
    ).resolves.toBe(true)
    await expect(
      ensure(state, { startRow: 10, endRow: 13, startColumn: 0, endColumn: 4 }),
    ).resolves.toBe(false)
  })

  it('rejects ranges beyond the shrunk extent after delete_rows', async () => {
    const state = streamedState({
      loaded: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 4 },
      ops: [{ kind: 'remove-rows', index: 0, count: 4 }],
    })
    await expect(
      ensure(state, { startRow: 0, endRow: 7, startColumn: 0, endColumn: 4 }),
    ).resolves.toBe(false)
    await expect(
      ensure(state, { startRow: 0, endRow: 5, startColumn: 0, endColumn: 4 }),
    ).resolves.toBe(true)
  })

  it('handles column inserts the same way', async () => {
    const state = streamedState({
      loaded: { startRow: 0, endRow: 9, startColumn: 0, endColumn: 6 },
      ops: [{ kind: 'insert-cols', index: 1, count: 2 }],
    })
    await expect(
      ensure(state, { startRow: 0, endRow: 0, startColumn: 5, endColumn: 6 }),
    ).resolves.toBe(true)
    await expect(
      ensure(state, { startRow: 0, endRow: 0, startColumn: 5, endColumn: 7 }),
    ).resolves.toBe(false)
  })

  it('replays journal cells when the range is entirely journal-owned', async () => {
    // Column 5 was inserted this session and its header/seed cells written,
    // then the streaming window moved away (evicting them from the grid).
    // Re-loading the inserted column must reinstall the journal cells, or
    // reads see empty cells where this session's edits live.
    const state = streamedState({
      loaded: { startRow: 100, endRow: 180, startColumn: 0, endColumn: 4 },
      ops: [{ kind: 'insert-cols', index: 5, count: 1 }],
      journalCells: [
        { row: 0, column: 5, value: 'Owner' },
        { row: 1, column: 5, value: 'merrick' },
      ],
    })
    const written: Array<{ row: number; column: number; data: unknown }> = []
    const capturingWorksheet = {
      getSheetId: () => 'sheet-1',
      getSheet: () => ({ getRowManager: () => ({ getRow: () => undefined }) }),
      getRange: (row: number, column: number) => ({
        setValues: (values: unknown[][]) => written.push({ row, column, data: values[0]?.[0] }),
        clearContent: () => written.push({ row, column, data: null }),
        clearFormat: () => undefined,
      }),
    }
    await expect(
      ensureLazyRangeLoaded(
        {
          univerAPI: {
            getActiveWorkbook: () => ({ getActiveSheet: () => capturingWorksheet }),
          },
        } as never,
        { current: state },
        capturingWorksheet as never,
        { startRow: 0, endRow: 1, startColumn: 5, endColumn: 5 },
        (message) => {
          throw new Error(message)
        },
      ),
    ).resolves.toBe(true)
    expect(written).toEqual([
      // Evict the old streaming window before installing the journal-owned one.
      { row: 100, column: 0, data: null },
      { row: 0, column: 5, data: {} },
      { row: 0, column: 5, data: { v: 'Owner' } },
      { row: 1, column: 5, data: { v: 'merrick' } },
    ])
  })
})

describe('lazyRangeEditable', () => {
  it('requires the loaded window for file-backed ranges', () => {
    const state = streamedState({ loaded: LOADED_TOP })
    expect(
      lazyRangeEditable(state, 'sheet-1', {
        startRow: 0,
        endRow: 4,
        startColumn: 0,
        endColumn: 4,
      }),
    ).toBe(true)
    expect(
      lazyRangeEditable(state, 'sheet-1', {
        startRow: 0,
        endRow: 7,
        startColumn: 0,
        endColumn: 4,
      }),
    ).toBe(false)
  })

  it('allows a range fully inside a column inserted this session', () => {
    // The fill-source case: CT2 lives in a freshly inserted column while the
    // streaming window sits somewhere else entirely.
    const state = streamedState({
      loaded: { startRow: 100, endRow: 180, startColumn: 0, endColumn: 4 },
      ops: [{ kind: 'insert-cols', index: 3, count: 1 }],
    })
    expect(
      lazyRangeEditable(state, 'sheet-1', {
        startRow: 0,
        endRow: 1,
        startColumn: 3,
        endColumn: 3,
      }),
    ).toBe(true)
  })

  it('allows a range fully inside rows inserted this session', () => {
    const state = streamedState({ ops: [{ kind: 'insert-rows', index: 2, count: 3 }] })
    expect(
      lazyRangeEditable(state, 'sheet-1', {
        startRow: 2,
        endRow: 4,
        startColumn: 0,
        endColumn: 4,
      }),
    ).toBe(true)
  })

  it('still requires the window for the file-backed part of a mixed range', () => {
    const ops = [{ kind: 'insert-cols', index: 2, count: 1 }]
    const bounds = { startRow: 0, endRow: 1, startColumn: 1, endColumn: 3 }
    // Screen columns 1 and 3 are file-backed; column 2 is journal-owned.
    expect(
      lazyRangeEditable(
        streamedState({ loaded: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 3 }, ops }),
        'sheet-1',
        bounds,
      ),
    ).toBe(true)
    expect(
      lazyRangeEditable(
        streamedState({ loaded: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 }, ops }),
        'sheet-1',
        bounds,
      ),
    ).toBe(false)
  })

  it('allows ranges entirely beyond the screen extent', () => {
    const state = streamedState({})
    expect(
      lazyRangeEditable(state, 'sheet-1', {
        startRow: 20,
        endRow: 25,
        startColumn: 0,
        endColumn: 4,
      }),
    ).toBe(true)
  })
})

describe('nextIndexWaitStall', () => {
  it('resets the stall count whenever indexing advances', () => {
    expect(nextIndexWaitStall(39, 100, 150)).toBe(0)
    expect(nextIndexWaitStall(5, null, 1)).toBe(0)
  })

  it('accumulates consecutive stalls and gives up at the limit', () => {
    expect(nextIndexWaitStall(0, 100, 100)).toBe(1)
    expect(nextIndexWaitStall(38, 100, 100)).toBe(39)
    expect(nextIndexWaitStall(39, 100, 100)).toBeNull()
  })

  it('treats a stream that has not reached the sheet as a stall', () => {
    expect(nextIndexWaitStall(0, null, null)).toBe(1)
  })
})

describe('applyRangeInLoadedChunks', () => {
  const runtime = { univerAPI: { getActiveWorkbook: () => null } }
  const worksheet = {
    getSheetId: () => 'sheet-1',
    getSheet: () => ({ getRowManager: () => ({ getRow: () => undefined }) }),
    getRange: () => ({
      setValues: () => undefined,
      clearContent: () => undefined,
      clearFormat: () => undefined,
    }),
  }

  afterEach(() => {
    delete (globalThis as { window?: unknown }).window
    vi.useRealTimers()
  })

  /// Runs a chunked apply that ends up polling the stalled indexer, fast-
  /// forwarding through the 250ms waits until it gives up.
  async function applyUntilSettled(
    state: LazyWorkbookState,
    bounds: Range,
    options: { neighborColumns: boolean },
  ): Promise<void> {
    vi.useFakeTimers()
    const promise = applyRangeInLoadedChunks(
      runtime as never,
      { current: state },
      worksheet as never,
      bounds,
      () => undefined,
      () => undefined,
      options,
    ).catch(() => undefined)
    await vi.advanceTimersByTimeAsync(15_000)
    await promise
  }

  function stubSidecar(indexedThroughRow: number): Array<Range> {
    const requests: Array<Range> = []
    ;(globalThis as { window?: unknown }).window = {
      desktopApi: {
        readWorkbookRange: async ({ range }: { range: Range }) => {
          requests.push(range)
          return {
            cells: [],
            rows: [],
            merges: [],
            hyperlinks: [],
            conditionalRules: [],
            autoFilter: null,
            dataValidations: [],
            sheetProtection: null,
            indexedThroughRow,
            indexingComplete: false,
          }
        },
      },
    }
    return requests
  }

  it('applies a value-only edit into an inserted column without any sidecar read', async () => {
    // The "insert a column, fill 88k rows with a default" scenario: the
    // target column is journal-owned, so with neighborColumns: false the
    // chunk loads resolve instantly no matter how far indexing has gotten.
    const requests = stubSidecar(0)
    const state = streamedState({
      loaded: { startRow: 100, endRow: 180, startColumn: 0, endColumn: 4 },
      ops: [{ kind: 'insert-cols', index: 5, count: 1 }],
    })
    const chunks: Array<Range> = []
    await applyRangeInLoadedChunks(
      runtime as never,
      { current: state },
      worksheet as never,
      { startRow: 0, endRow: 9, startColumn: 5, endColumn: 5 },
      (chunk) => chunks.push(chunk),
      () => undefined,
      { neighborColumns: false },
    )
    expect(requests).toEqual([])
    expect(chunks).toEqual([{ startRow: 0, endRow: 9, startColumn: 5, endColumn: 5 }])
  })

  it('evicts each prior journal-owned chunk during a whole-column fill', async () => {
    const requests = stubSidecar(0)
    const state = streamedState({
      rows: 50_000,
      loaded: { startRow: 0, endRow: 79, startColumn: 0, endColumn: 4 },
      ops: [{ kind: 'insert-cols', index: 5, count: 1 }],
    })
    const cleared: Range[] = []
    const evictingWorksheet = {
      getSheetId: () => 'sheet-1',
      getSheet: () => ({ getRowManager: () => ({ getRow: () => undefined }) }),
      getRange: (row: number, column: number, rows: number, columns: number) => ({
        setValues: () => undefined,
        clearContent: () =>
          cleared.push({
            startRow: row,
            endRow: row + rows - 1,
            startColumn: column,
            endColumn: column + columns - 1,
          }),
        clearFormat: () => undefined,
      }),
    }
    const chunks: Range[] = []
    await applyRangeInLoadedChunks(
      runtime as never,
      { current: state },
      evictingWorksheet as never,
      { startRow: 0, endRow: 49_999, startColumn: 5, endColumn: 5 },
      (chunk) => chunks.push(chunk),
      () => undefined,
      { neighborColumns: false },
    )
    expect(requests).toEqual([])
    expect(chunks).toHaveLength(3)
    expect(cleared).toEqual([
      { startRow: 0, endRow: 79, startColumn: 0, endColumn: 4 },
      { startRow: 0, endRow: 19_999, startColumn: 5, endColumn: 5 },
      { startRow: 20_000, endRow: 39_999, startColumn: 5, endColumn: 5 },
    ])
    expect(state.loadedRanges.get('sheet-1')).toEqual({
      startRow: 40_000,
      endRow: 49_999,
      startColumn: 5,
      endColumn: 5,
    })
  })

  it('loads only the target columns for value-only edits on file-backed cells', async () => {
    const requests = stubSidecar(0)
    const state = streamedState({
      loaded: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 4 },
    })
    await applyUntilSettled(
      state,
      { startRow: 0, endRow: 9, startColumn: 2, endColumn: 2 },
      { neighborColumns: false },
    )
    expect(requests.length).toBeGreaterThan(0)
    expect(requests[0]).toMatchObject({ startColumn: 2, endColumn: 2 })
  })

  it('widens chunk loads to the full sheet width when neighbor columns are requested', async () => {
    const requests = stubSidecar(0)
    const state = streamedState({
      loaded: { startRow: 0, endRow: 4, startColumn: 0, endColumn: 4 },
    })
    await applyUntilSettled(
      state,
      { startRow: 0, endRow: 9, startColumn: 2, endColumn: 2 },
      { neighborColumns: true },
    )
    expect(requests.length).toBeGreaterThan(0)
    expect(requests[0]).toMatchObject({ startColumn: 0, endColumn: 4 })
  })
})
