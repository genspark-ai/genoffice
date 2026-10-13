// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import {
  handleExportPdf,
  openPrintPreviewHost,
  type PageLayoutContext,
} from '../src/renderer/page-layout-actions'
import { createEditJournal } from '../src/renderer/edit-journal'
import { startFullLoad } from '../src/renderer/full-load'
import { workbookCellCounts } from '../src/renderer/load-budget'
import { setModuleLang } from '../src/renderer/i18n/locale'
import { fullLoadGate } from '../src/renderer/reorder-gate'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'
import { preloadEntireWorkbook } from '../src/renderer/univer-sync'

vi.mock('../src/renderer/univer-sync', async (importOriginal) => {
  const actual = await importOriginal<typeof import('../src/renderer/univer-sync')>()
  return {
    ...actual,
    preloadEntireWorkbook: vi.fn(),
  }
})

type FileSheet = LazyWorkbookState['file']['sheets'][number]

/// Behaves like the real preloadEntireWorkbook as far as the callers can see:
/// flags the workbook running, finishes after `ms`, then flags it complete.
function fakePreload(ms: number) {
  vi.mocked(preloadEntireWorkbook).mockImplementation(async (_runtime, ref) => {
    const state = ref.current
    if (!state) return
    state.flags.preloadRunning = true
    await new Promise((resolve) => setTimeout(resolve, ms))
    state.flags.preloadRunning = false
    state.flags.preloadComplete = true
  })
}

describe('print and PDF workbook readiness coordination', () => {
  let messages: string[] = []
  let printWorkbookMock: ReturnType<typeof vi.fn>
  let exportPdfMock: ReturnType<typeof vi.fn>

  beforeEach(() => {
    vi.mocked(preloadEntireWorkbook).mockReset()
    setModuleLang('en')
    messages = []
    printWorkbookMock = vi.fn().mockResolvedValue({ ok: true })
    exportPdfMock = vi.fn().mockResolvedValue({ canceled: false, path: 'D:/test.pdf' })

    vi.stubGlobal('window', {
      desktopApi: {
        printWorkbook: printWorkbookMock,
        exportPdf: exportPdfMock,
      },
    })
  })

  afterEach(() => {
    vi.useRealTimers()
  })

  function sheet(overrides: Partial<FileSheet> = {}): FileSheet {
    return {
      id: 'sheet-1',
      name: 'Sheet1',
      rowCount: 10,
      columnCount: 10,
      ...overrides,
    } as FileSheet
  }

  function createMockState(
    sheets: readonly FileSheet[] = [sheet()],
    preloadRunning = false,
  ): LazyWorkbookState {
    return {
      file: { name: 'test.xlsx', sessionId: 'session-123', sheets, visuals: [] },
      flags: { preloadComplete: false, preloadRunning },
      editJournal: createEditJournal(),
      sheetFilePageSetups: new Map(),
      sheetPageBreaks: new Map(),
    } as unknown as LazyWorkbookState
  }

  function createMockContext(state: LazyWorkbookState | null): PageLayoutContext {
    const worksheet = {
      getSheetId: () => 'sheet-1',
      getSheetName: () => 'Sheet1',
      getLastRow: () => 1,
      getLastColumn: () => 1,
      getRowHeight: () => 20,
      getColumnWidth: () => 100,
      getMergedRanges: () => [],
      getRange: () => ({
        getValues: () => [['A1']],
        getDisplayValues: () => [['A1']],
        getCellStyleData: () => null,
      }),
      getMaxRows: () => 10,
      getMaxColumns: () => 10,
      isSheetHidden: () => false,
    }
    const workbook = {
      getActiveSheet: () => worksheet,
      getActiveRange: () => null,
      getSheets: () => [worksheet],
      getSheetBySheetId: () => worksheet,
    }
    const runtime = {
      univerAPI: { getActiveWorkbook: () => workbook },
    } as unknown as UniverRuntime

    return {
      univerRef: { current: runtime },
      lazyWorkbookRef: { current: state },
      setMessage: (message) => {
        messages.push(message)
      },
      setPendingEdits: vi.fn(),
      requestVisualInstall: vi.fn(),
      runOps: () => Promise.reject(new Error('page layout ops are not used here')),
    }
  }

  it('PDF export of a streamed workbook preloads it first and exports', async () => {
    const state = createMockState()
    const ctx = createMockContext(state)
    fakePreload(10)

    expect(await handleExportPdf(ctx)).toBe(true)
    expect(preloadEntireWorkbook).toHaveBeenCalledTimes(1)
    expect(exportPdfMock).toHaveBeenCalledTimes(1)
    expect(messages).toContain(
      'PDF export needs the fully loaded workbook — wait for loading to finish.',
    )
    expect(messages).toContain('Exported D:/test.pdf.')
  })

  it('Print of a streamed workbook preloads it first and opens the dialog host', async () => {
    const state = createMockState()
    const ctx = createMockContext(state)
    fakePreload(10)

    const host = await openPrintPreviewHost(ctx)
    expect(host).not.toBeNull()
    expect(preloadEntireWorkbook).toHaveBeenCalledTimes(1)
    expect(state.flags.preloadComplete).toBe(true)
  })

  it('opens the Print dialog at once for a fully loaded workbook', async () => {
    const state = createMockState()
    state.flags.preloadComplete = true

    expect(await openPrintPreviewHost(createMockContext(state))).not.toBeNull()
    expect(preloadEntireWorkbook).not.toHaveBeenCalled()
  })

  it('joins a load the user already started instead of starting a second one', async () => {
    const state = createMockState()
    const ctx = createMockContext(state)
    fakePreload(20)
    const userLoad = startFullLoad(ctx.univerRef.current!, ctx.lazyWorkbookRef, () => undefined)

    expect(await handleExportPdf(ctx)).toBe(true)
    await userLoad
    expect(preloadEntireWorkbook).toHaveBeenCalledTimes(1)
    expect(exportPdfMock).toHaveBeenCalledTimes(1)
  })

  it('a Load-all click during a print-triggered preload joins it', async () => {
    const state = createMockState()
    const ctx = createMockContext(state)
    fakePreload(20)

    const pdf = handleExportPdf(ctx)
    expect(state.flags.preloadRunning).toBe(true)
    const userLoad = startFullLoad(ctx.univerRef.current!, ctx.lazyWorkbookRef, () => undefined)

    expect(await pdf).toBe(true)
    await userLoad
    expect(preloadEntireWorkbook).toHaveBeenCalledTimes(1)
  })

  it('a filter click during a print-triggered preload gets the in-progress notice', async () => {
    const state = createMockState()
    const ctx = createMockContext(state)
    const cellCounts = workbookCellCounts(state.file.sheets)
    const gate = () =>
      fullLoadGate({ formulaMode: false, preloadRunning: state.flags.preloadRunning, cellCounts })
    fakePreload(20)

    expect(gate()).toBe('offerFullLoad')
    const pdf = handleExportPdf(ctx)
    expect(gate()).toBe('loading')
    await pdf
    expect(state.flags.preloadComplete).toBe(true)
  })

  it('a failed preload cancels the export and the Print dialog with the failure message', async () => {
    const state = createMockState()
    const ctx = createMockContext(state)
    vi.mocked(preloadEntireWorkbook).mockRejectedValue(new Error('Preload disk error'))

    expect(await handleExportPdf(ctx)).toBe(false)
    expect(await openPrintPreviewHost(ctx)).toBeNull()
    expect(exportPdfMock).not.toHaveBeenCalled()
    expect(messages).toContain('Unable to export the PDF.')
    expect(messages).toContain('Unable to print.')
  })

  it('a preload that ends without completing cancels the export', async () => {
    const state = createMockState()
    const ctx = createMockContext(state)
    vi.mocked(preloadEntireWorkbook).mockResolvedValue(undefined)

    expect(await handleExportPdf(ctx)).toBe(false)
    expect(exportPdfMock).not.toHaveBeenCalled()
    expect(messages.at(-1)).toBe('Unable to export the PDF.')
  })

  it('a preload that never finishes gives up after the timeout', async () => {
    vi.useFakeTimers()
    const state = createMockState()
    const ctx = createMockContext(state)
    vi.mocked(preloadEntireWorkbook).mockReturnValue(new Promise(() => undefined))

    const pdf = handleExportPdf(ctx)
    await vi.advanceTimersByTimeAsync(180_000)

    expect(await pdf).toBe(false)
    expect(exportPdfMock).not.toHaveBeenCalled()
    expect(messages.at(-1)).toBe('Unable to export the PDF.')
  })

  it('a workbook over the full-load budget is refused without preloading, with a wording that fits', async () => {
    const state = createMockState([sheet({ rowCount: 10_000, columnCount: 200 })])
    const ctx = createMockContext(state)

    expect(await handleExportPdf(ctx)).toBe(false)
    expect(await openPrintPreviewHost(ctx)).toBeNull()
    expect(preloadEntireWorkbook).not.toHaveBeenCalled()
    expect(exportPdfMock).not.toHaveBeenCalled()
    expect(messages).toEqual([
      'This workbook is too large to load fully into memory, so it cannot be printed or exported to PDF.',
      'This workbook is too large to load fully into memory, so it cannot be printed or exported to PDF.',
    ])
  })

  it('a sparse workbook inside the shared budget is loaded, not refused', async () => {
    // 5M-cell bounding box but only 100k stored cells: the Load-all prompt
    // offers this, so Print and PDF export must too.
    const state = createMockState([
      sheet({ rowCount: 50_000, columnCount: 100, storedCellCount: 100_000 }),
    ])
    const ctx = createMockContext(state)
    fakePreload(5)

    expect(await handleExportPdf(ctx)).toBe(true)
    expect(preloadEntireWorkbook).toHaveBeenCalledTimes(1)
  })

  it('switching workbooks while preloading aborts the export quietly', async () => {
    const stateA = createMockState()
    const stateB = createMockState([])
    const ctx = createMockContext(stateA)
    vi.mocked(preloadEntireWorkbook).mockImplementation(async () => {
      await new Promise((resolve) => setTimeout(resolve, 10))
      ctx.lazyWorkbookRef.current = stateB
      stateA.flags.preloadComplete = true
    })
    messages.length = 0

    expect(await handleExportPdf(ctx)).toBe(false)
    expect(exportPdfMock).not.toHaveBeenCalled()
    expect(messages).not.toContain('Unable to export the PDF.')
  })

  it('an untitled workbook (no lazy state) prints and exports without a preload', async () => {
    const ctx = createMockContext(null)

    expect(await openPrintPreviewHost(ctx)).not.toBeNull()
    expect(await handleExportPdf(ctx)).toBe(true)
    expect(exportPdfMock).toHaveBeenCalledTimes(1)
    expect(preloadEntireWorkbook).not.toHaveBeenCalled()
  })

  it('concurrent Print and PDF export share one preload', async () => {
    const state = createMockState()
    const ctx = createMockContext(state)
    fakePreload(30)

    const [host, pdfOk] = await Promise.all([openPrintPreviewHost(ctx), handleExportPdf(ctx)])
    expect(host).not.toBeNull()
    expect(pdfOk).toBe(true)
    expect(preloadEntireWorkbook).toHaveBeenCalledTimes(1)
    expect(exportPdfMock).toHaveBeenCalledTimes(1)
  })
})
