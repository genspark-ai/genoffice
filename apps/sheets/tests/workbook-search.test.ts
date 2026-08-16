import { describe, expect, it, vi } from 'vitest'

import { findWorkbookCells, selectWorkbookRange } from '../src/renderer/ai/workbook-search'
import { ensureLazyRangeLoaded, readSheetRangeMapped } from '../src/renderer/univer-sync'
import type { FindCellsOptions } from '../src/renderer/ai/tools'
import type { WorkbookReadContext } from '../src/renderer/ai/workbook-readers'

vi.mock('../src/renderer/univer-sync', () => ({
  readSheetRangeMapped: vi.fn(),
  ensureLazyRangeLoaded: vi.fn().mockResolvedValue(true),
}))

function options(overrides: Partial<FindCellsOptions> = {}): FindCellsOptions {
  return {
    query: '',
    regex: false,
    lookIn: 'both',
    errorsOnly: false,
    maxResults: 50,
    ...overrides,
  }
}

type DemoSheet = {
  id: string
  name: string
  cells: Record<string, { value: string | number | boolean | null; formula?: string }>
}

function demoCtx(sheets: DemoSheet[]): WorkbookReadContext {
  return {
    univerRef: { current: null },
    lazyWorkbookRef: { current: null },
    adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets }) } },
  } as unknown as WorkbookReadContext
}

const DEMO_SHEETS: DemoSheet[] = [
  {
    id: 'sheet-1',
    name: 'Sheet1',
    cells: {
      A1: { value: 'Total revenue' },
      A2: { value: 120 },
      B2: { value: 240, formula: '=A2*2' },
      C1: { value: '#REF!', formula: '=Gone!A1' },
    },
  },
  { id: 'sheet-2', name: 'Summary', cells: { A1: { value: 'grand TOTAL' } } },
]

describe('findWorkbookCells: demo workbook', () => {
  it('matches values case-insensitively across sheets', async () => {
    const result = await findWorkbookCells(demoCtx(DEMO_SHEETS), options({ query: 'total' }))
    expect(result.error).toBeUndefined()
    expect(result.matches.map((m) => `${m.sheetName}!${m.address}`)).toEqual([
      'Sheet1!A1',
      'Summary!A1',
    ])
    expect(result.truncated).toBe(false)
  })

  it('restricts matching to formulas when asked', async () => {
    const result = await findWorkbookCells(
      demoCtx(DEMO_SHEETS),
      options({ query: 'a2', lookIn: 'formulas' }),
    )
    expect(result.matches).toHaveLength(1)
    expect(result.matches[0]?.address).toBe('B2')
  })

  it('finds formula error values with errors_only', async () => {
    const result = await findWorkbookCells(demoCtx(DEMO_SHEETS), options({ errorsOnly: true }))
    expect(result.matches).toHaveLength(1)
    expect(result.matches[0]?.value).toBe('#REF!')
  })

  it('supports regex matching and reports invalid patterns', async () => {
    const regexResult = await findWorkbookCells(
      demoCtx(DEMO_SHEETS),
      options({ query: '^grand', regex: true }),
    )
    expect(regexResult.matches.map((m) => m.sheetName)).toEqual(['Summary'])
    const invalid = await findWorkbookCells(
      demoCtx(DEMO_SHEETS),
      options({ query: '(', regex: true }),
    )
    expect(invalid.error).toContain('Invalid regex')
  })

  it('truncates at maxResults and flags it', async () => {
    const result = await findWorkbookCells(
      demoCtx(DEMO_SHEETS),
      options({ query: 'total', maxResults: 1 }),
    )
    expect(result.matches).toHaveLength(1)
    expect(result.truncated).toBe(true)
  })

  it('rejects an unknown sheetId', async () => {
    const result = await findWorkbookCells(
      demoCtx(DEMO_SHEETS),
      options({ query: 'total', sheetId: 'ghost' }),
    )
    expect(result.error).toContain('Unknown sheet: ghost')
  })
})

function lazyCtx(state: unknown, worksheets: unknown[]): WorkbookReadContext {
  return {
    univerRef: {
      current: {
        univerAPI: {
          getActiveWorkbook: () => ({
            getSheets: () => worksheets,
            getActiveSheet: () => worksheets[0],
          }),
        },
      },
    },
    lazyWorkbookRef: { current: state },
    adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } },
  } as unknown as WorkbookReadContext
}

function lazyState(journalCells: Map<string, Map<string, unknown>>) {
  return {
    file: {
      sessionId: 'session-1',
      sheets: [{ id: 'sh1', name: 'Data', rowCount: 4, columnCount: 2 }],
    },
    editJournal: { cells: journalCells, structuralOps: new Map() },
  }
}

const DATA_SHEET = {
  getSheetId: () => 'sh1',
  getSheetName: () => 'Data',
  getRange: (address: string) => ({ getValue: () => (address === 'A3' ? '#DIV/0!' : null) }),
}

describe('findWorkbookCells: lazy workbook', () => {
  it('overlays journal edits and shadows the file cell underneath', async () => {
    const journal = new Map([
      [
        'sh1',
        new Map([
          ['0:0', { row: 0, column: 0, hasValue: true, value: 'edited total' }],
          ['0:1', { row: 0, column: 1, hasValue: false, value: null }],
        ]),
      ],
    ])
    vi.mocked(readSheetRangeMapped).mockResolvedValue({
      screen: {
        cells: [
          { row: 0, column: 0, value: 'file total' },
          { row: 1, column: 0, value: 'total again' },
          { row: 1, column: 1, value: 'unrelated' },
        ],
        rows: [],
        merges: [],
        hyperlinks: [],
      },
      raw: { indexingComplete: true },
      indexedThroughScreen: 3,
      fileEndRow: 3,
    } as never)
    const result = await findWorkbookCells(
      lazyCtx(lazyState(journal), [DATA_SHEET]),
      options({ query: 'total' }),
    )
    expect(result.matches.map((m) => `${m.address}: ${String(m.value)}`)).toEqual([
      'A1: edited total',
      'A2: total again',
    ])
    expect(result.incompleteSheets).toEqual([])
  })

  it('backfills computed values for journal formula cells (errors_only sees them)', async () => {
    const journal = new Map([
      [
        'sh1',
        new Map([['2:0', { row: 2, column: 0, hasValue: true, value: null, formula: '=A1/0' }]]),
      ],
    ])
    vi.mocked(readSheetRangeMapped).mockResolvedValue({
      screen: { cells: [], rows: [], merges: [], hyperlinks: [] },
      raw: { indexingComplete: true },
      indexedThroughScreen: 3,
      fileEndRow: 3,
    } as never)
    const result = await findWorkbookCells(
      lazyCtx(lazyState(journal), [DATA_SHEET]),
      options({ errorsOnly: true }),
    )
    expect(result.matches).toEqual([
      { sheetName: 'Data', address: 'A3', value: '#DIV/0!', formula: '=A1/0' },
    ])
  })

  it('flags sheets whose indexing has not caught up', async () => {
    vi.mocked(readSheetRangeMapped).mockResolvedValue({
      screen: { cells: [], rows: [], merges: [], hyperlinks: [] },
      raw: { indexingComplete: false },
      indexedThroughScreen: 0,
      fileEndRow: 3,
    } as never)
    const result = await findWorkbookCells(
      lazyCtx(lazyState(new Map()), [DATA_SHEET]),
      options({ query: 'total' }),
    )
    expect(result.matches).toEqual([])
    expect(result.incompleteSheets).toEqual(['Data'])
  })
})

describe('selectWorkbookRange', () => {
  function selectionCtx() {
    const range = { activate: vi.fn() }
    const worksheet = {
      getSheetId: () => 'sh1',
      getSheetName: () => 'Data',
      getRange: vi.fn().mockReturnValue(range),
      scrollToCell: vi.fn(),
    }
    const workbook = {
      getActiveSheet: () => worksheet,
      getSheetBySheetId: (id: string) => (id === 'sh1' ? worksheet : null),
      setActiveSheet: vi.fn(),
    }
    const ctx = {
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: { current: null },
      adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } },
    } as unknown as WorkbookReadContext
    return { ctx, worksheet, workbook, range }
  }

  it('selects and scrolls to the range on the active sheet', async () => {
    const { ctx, worksheet, range } = selectionCtx()
    const result = await selectWorkbookRange(
      ctx,
      undefined,
      { startRow: 1, startColumn: 1, endRow: 3, endColumn: 2 },
      () => {},
    )
    expect(result).toEqual({ ok: true, sheetName: 'Data' })
    expect(worksheet.getRange).toHaveBeenCalledWith(1, 1, 3, 2)
    expect(range.activate).toHaveBeenCalled()
    expect(worksheet.scrollToCell).toHaveBeenCalledWith(1, 1)
  })

  it('loads the target range first on lazy workbooks', async () => {
    const { ctx, worksheet } = selectionCtx()
    ;(ctx.lazyWorkbookRef as { current: unknown }).current = lazyState(new Map())
    const result = await selectWorkbookRange(
      ctx,
      'sh1',
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
      () => {},
    )
    expect(result.ok).toBe(true)
    expect(vi.mocked(ensureLazyRangeLoaded)).toHaveBeenCalled()
    expect(worksheet.scrollToCell).toHaveBeenCalledWith(0, 0)
  })

  it('reports an unknown sheet', async () => {
    const { ctx } = selectionCtx()
    const result = await selectWorkbookRange(
      ctx,
      'ghost',
      { startRow: 0, startColumn: 0, endRow: 0, endColumn: 0 },
      () => {},
    )
    expect(result).toEqual({ ok: false, error: 'Unknown sheet: ghost' })
  })
})
