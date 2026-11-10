import { beforeEach, describe, expect, it, vi } from 'vitest'

import {
  pickNextError,
  runStreamedErrorCheck,
  scanStreamedWorkbookErrors,
  type SheetError,
} from '../src/renderer/error-checking'
import { FILE_READ_BATCH_CELLS, MAX_SCAN_CELLS } from '../src/renderer/ai/workbook-search'
import { ensureLazyRangeLoaded, readSheetRangeMapped } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

vi.mock('../src/renderer/univer-sync', () => ({
  readSheetRangeMapped: vi.fn(),
  ensureLazyRangeLoaded: vi.fn().mockResolvedValue(true),
}))

vi.mock('../src/renderer/i18n/locale', () => ({
  t: (key: string, params?: Record<string, string>) =>
    `${key}${params ? ` ${JSON.stringify(params)}` : ''}`,
}))

const mockRead = vi.mocked(readSheetRangeMapped)
const mockEnsure = vi.mocked(ensureLazyRangeLoaded)

type MappedResult = Awaited<ReturnType<typeof readSheetRangeMapped>>

function mapped(
  cells: {
    row: number
    column: number
    value: string | number | boolean | null
    formula?: string
  }[],
): MappedResult {
  return {
    raw: { indexingComplete: true },
    indexedThroughScreen: 499,
    fileEndRow: 499,
    screen: { cells, rows: [], merges: [], hyperlinks: [] },
  } as unknown as MappedResult
}

function state(): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      sheets: [
        { id: 's1', name: 'Sheet1', rowCount: 500, columnCount: 4 },
        { id: 's2', name: 'Data', rowCount: 500, columnCount: 4 },
      ],
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
    editJournal: {
      cells: new Map(),
      structuralOps: new Map(),
    },
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

describe('pickNextError', () => {
  const order = new Map([
    ['s1', 0],
    ['s2', 1],
  ])
  const err = (sheetId: string, row: number, column: number): SheetError => ({
    sheetId,
    row,
    column,
    value: '#REF!',
  })

  it('picks the first error after the active cell', () => {
    const errors = [err('s1', 2, 0), err('s1', 40, 1), err('s2', 3, 3)]
    expect(pickNextError(errors, order, 's1', 10, 0)).toEqual(err('s1', 40, 1))
    expect(pickNextError(errors, order, 's1', 40, 1)).toEqual(err('s2', 3, 3))
  })

  it('wraps to the first error after the last one', () => {
    const errors = [err('s1', 2, 0), err('s2', 3, 3)]
    expect(pickNextError(errors, order, 's2', 3, 3)).toEqual(err('s1', 2, 0))
  })

  it('returns null without errors', () => {
    expect(pickNextError([], order, 's1', 0, 0)).toBeNull()
  })
})

describe('scanStreamedWorkbookErrors', () => {
  beforeEach(() => {
    mockRead.mockReset()
  })

  it('collects journal and file errors, shadowing edited cells', async () => {
    const lazyState = state()
    lazyState.editJournal.cells.set(
      's1',
      new Map([
        ['0', { row: 5, column: 1, value: '#REF!', hasValue: true }],
        // An edit that overwrote a file-side error must hide it.
        ['1', { row: 20, column: 2, value: 'fixed', hasValue: true }],
      ]),
    )
    mockRead.mockImplementation(async (_state, sheetId) => {
      if (sheetId !== 's1') {
        return mapped([{ row: 300, column: 0, value: '#DIV/0!' }])
      }
      return mapped([
        { row: 20, column: 2, value: '#VALUE!' },
        { row: 400, column: 3, value: 'ok' },
      ])
    })

    const scan = await scanStreamedWorkbookErrors(lazyState)
    expect(scan.truncated).toBe(false)
    expect(scan.errors).toHaveLength(2)
    expect(scan.errors[0]).toMatchObject({ sheetId: 's1', row: 5, column: 1 })
    expect(scan.errors[1]).toMatchObject({ sheetId: 's2', row: 300, column: 0 })
  })

  it('flags truncation when a page read fails', async () => {
    mockRead.mockRejectedValue(new Error('sidecar gone'))
    const scan = await scanStreamedWorkbookErrors(state())
    expect(scan.errors).toHaveLength(0)
    expect(scan.truncated).toBe(true)
  })

  it('reads back journal formula results and reports their errors', async () => {
    const lazyState = state()
    lazyState.editJournal.cells.set(
      's1',
      new Map([['0', { row: 3, column: 1, value: null, formula: '=1/0', hasValue: true }]]),
    )
    mockRead.mockImplementation(async () => mapped([]))
    const scan = await scanStreamedWorkbookErrors(lazyState, {
      resolveFormulaValue: () => '#DIV/0!',
    })
    expect(scan.errors).toEqual([{ sheetId: 's1', row: 3, column: 1, value: '#DIV/0!' }])
  })

  it('does not let style-only journal entries hide file errors', async () => {
    const lazyState = state()
    lazyState.editJournal.cells.set(
      's1',
      new Map([['0', { row: 20, column: 2, value: null, hasValue: false }]]),
    )
    mockRead.mockImplementation(async (_state, sheetId) =>
      sheetId === 's1' ? mapped([{ row: 20, column: 2, value: '#VALUE!' }]) : mapped([]),
    )
    const scan = await scanStreamedWorkbookErrors(lazyState)
    expect(scan.errors).toEqual([{ sheetId: 's1', row: 20, column: 2, value: '#VALUE!' }])
  })

  it('follows live membership: session-added sheets scanned, deleted ones skipped', async () => {
    const lazyState = state()
    lazyState.editJournal.cells.set(
      's3',
      new Map([['0', { row: 1, column: 0, value: '#REF!', hasValue: true }]]),
    )
    mockRead.mockImplementation(async () => mapped([]))
    const scan = await scanStreamedWorkbookErrors(lazyState, { liveSheetIds: ['s2', 's3'] })
    expect(scan.errors).toEqual([{ sheetId: 's3', row: 1, column: 0, value: '#REF!' }])
    // s1 was deleted this session: its file pages are never read.
    expect(mockRead.mock.calls.every(([, sheetId]) => sheetId === 's2')).toBe(true)
  })

  it('stops at the scan-cell budget instead of paging the whole file', async () => {
    const lazyState = state()
    ;(lazyState.file.sheets as { rowCount: number }[])[0]!.rowCount = 300_000
    mockRead.mockImplementation(async () => mapped([]))
    const scan = await scanStreamedWorkbookErrors(lazyState)
    expect(scan.truncated).toBe(true)
    const cellsPerBatch = Math.floor(FILE_READ_BATCH_CELLS / 4) * 4
    expect(mockRead).toHaveBeenCalledTimes(Math.ceil(MAX_SCAN_CELLS / cellsPerBatch))
  })
})

describe('runStreamedErrorCheck', () => {
  function harness(lazyState: LazyWorkbookState | null, errors: SheetError[]) {
    const worksheet = {
      getSheetId: () => 's1',
      getRange: vi.fn(() => ({ activate: vi.fn() })),
    }
    const targetActivate = vi.fn()
    const target = {
      getSheetId: () => 's2',
      getRange: vi.fn(() => ({ activate: targetActivate })),
    }
    const workbook = {
      getSheets: () => [worksheet, target],
      getActiveSheet: () => worksheet,
      getSheetBySheetId: (sheetId: string) => (sheetId === 's2' ? target : worksheet),
      getActiveRange: () => ({ getRow: () => 0, getColumn: () => 0 }),
      setActiveSheet: vi.fn(),
    }
    return {
      runtime: {
        univerAPI: {
          getActiveWorkbook: () => workbook,
          executeCommand: vi.fn(async () => true),
        },
      } as unknown as Parameters<typeof runStreamedErrorCheck>[0]['runtime'],
      lazyWorkbookRef: { current: lazyState },
      setMessage: vi.fn(),
      refreshSelectionEcho: vi.fn(),
      workbook,
      target,
      targetActivate,
      errors,
    }
  }

  beforeEach(() => {
    mockRead.mockReset()
    mockEnsure.mockReset()
    mockEnsure.mockResolvedValue(true)
  })

  it('scans, jumps to the next error across sheets, and reports the count', async () => {
    mockRead.mockImplementation(async (_state, sheetId) =>
      sheetId === 's2' ? mapped([{ row: 300, column: 0, value: '#DIV/0!' }]) : mapped([]),
    )
    const harnessData = harness(state(), [])
    await runStreamedErrorCheck(harnessData)

    const messages = harnessData.setMessage.mock.calls.map(([text]) => String(text))
    expect(messages.some((text) => text.startsWith('appCheckingErrors'))).toBe(true)
    const finalMessage = messages.find((text) => text.startsWith('appErrorsFound'))
    expect(finalMessage).toContain('"count":1')
    // Cross-sheet jump activated the right sheet and loaded the hit's range.
    expect(harnessData.workbook.setActiveSheet).toHaveBeenCalledWith(harnessData.target)
    expect(mockEnsure).toHaveBeenCalled()
    expect(harnessData.target.getRange).toHaveBeenCalledWith(300, 0, 1, 1)
    expect(harnessData.targetActivate).toHaveBeenCalledTimes(1)
  })

  it('reports a clean sheet without jumping', async () => {
    mockRead.mockResolvedValue(mapped([{ row: 1, column: 1, value: 'fine' }]))
    const harnessData = harness(state(), [])
    await runStreamedErrorCheck(harnessData)

    const messages = harnessData.setMessage.mock.calls.map(([text]) => String(text))
    expect(messages[messages.length - 1]).toBe('appNoErrorsFound')
    expect(mockEnsure).not.toHaveBeenCalled()
  })

  it('ignores a second click while a scan is in flight', async () => {
    const pending: { release?: () => void } = {}
    // The first page read blocks until released; every later read resolves
    // immediately so the scan can finish.
    mockRead.mockImplementationOnce(
      async () =>
        new Promise<MappedResult>((resolve) => {
          pending.release = () => resolve(mapped([]))
        }),
    )
    mockRead.mockResolvedValue(mapped([]))
    const harnessData = harness(state(), [])
    const first = runStreamedErrorCheck(harnessData)
    await Promise.resolve()
    await runStreamedErrorCheck(harnessData)
    pending.release?.()
    await first
    expect(
      harnessData.setMessage.mock.calls.filter(([text]) => text === 'appCheckingErrors'),
    ).toHaveLength(1)
  })

  it('bails out when the workbook is swapped mid-scan', async () => {
    const harnessData = harness(state(), [])
    mockRead.mockImplementation(async () => {
      harnessData.lazyWorkbookRef.current = state() // file switch during the paged read
      return mapped([{ row: 300, column: 0, value: '#DIV/0!' }])
    })
    await runStreamedErrorCheck(harnessData)
    expect(mockEnsure).not.toHaveBeenCalled()
    expect(harnessData.workbook.setActiveSheet).not.toHaveBeenCalled()
    // Only the initial scanning status — no stale found/not-found result.
    expect(harnessData.setMessage).toHaveBeenCalledTimes(1)
    // The in-flight guard is released: the new workbook scans normally.
    mockRead.mockImplementation(async () => mapped([]))
    await runStreamedErrorCheck(harnessData)
    expect(harnessData.setMessage.mock.calls.map(([text]) => text)).toEqual([
      'appCheckingErrors',
      'appCheckingErrors',
      'appNoErrorsFound',
    ])
  })
})
