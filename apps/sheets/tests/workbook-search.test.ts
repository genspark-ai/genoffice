import { afterEach, describe, expect, it, vi } from 'vitest'

import {
  ERROR_VALUE_RE,
  findWorkbookCells,
  selectWorkbookRange,
} from '../src/renderer/ai/workbook-search'
import { ensureLazyRangeLoaded, readSheetRangeMapped } from '../src/renderer/univer-sync'
import type { FindCellsOptions } from '../src/renderer/ai/tools'
import type { WorkbookReadContext } from '../src/renderer/ai/workbook-readers'
import { buildRedactionIndex } from '../src/renderer/ai/redact'

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

  it('recognizes every error value Excel can display, modern data-type errors included', () => {
    for (const value of [
      '#REF!',
      '#DIV/0!',
      '#VALUE!',
      '#NAME?',
      '#N/A',
      '#NUM!',
      '#NULL!',
      '#SPILL!',
      '#CALC!',
      '#FIELD!',
      '#CONNECT!',
      '#BLOCKED!',
      '#UNKNOWN!',
      '#GETTING_DATA',
    ]) {
      expect(ERROR_VALUE_RE.test(value)).toBe(true)
    }
    for (const value of ['#FOO!', 'N/A', '#REF', '', '#12345!']) {
      expect(ERROR_VALUE_RE.test(value)).toBe(false)
    }
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

describe('findWorkbookCells: lazy workbook via sidecar', () => {
  afterEach(() => {
    vi.unstubAllGlobals()
  })

  it('searches the file in the sidecar and overlays journal edits', async () => {
    const sidecarFind = vi.fn().mockResolvedValue({
      matches: [
        { sheetId: 'sh1', row: 0, column: 0, value: 'file total', valueText: 'file total' },
        { sheetId: 'sh1', row: 1, column: 0, value: 'total again', valueText: 'total again' },
        { sheetId: 'sh1', row: 2, column: 1, value: 9, valueText: '9', formulaText: '=TOTAL()' },
      ],
      complete: true,
      indexingComplete: true,
    })
    vi.stubGlobal('window', { desktopApi: { findWorkbookCells: sidecarFind } })
    vi.mocked(readSheetRangeMapped).mockReset()
    const journal = new Map([
      ['sh1', new Map([['0:0', { row: 0, column: 0, hasValue: true, value: 'edited total' }]])],
    ])
    const result = await findWorkbookCells(
      lazyCtx(lazyState(journal), [DATA_SHEET]),
      options({ query: 'total' }),
    )
    expect(result.matches.map((m) => `${m.address}: ${String(m.value)}`)).toEqual([
      'A1: edited total',
      'A2: total again',
      'B3: 9',
    ])
    expect(result.truncated).toBe(false)
    expect(result.incompleteSheets).toEqual([])
    expect(readSheetRangeMapped).not.toHaveBeenCalled()
    expect(sidecarFind.mock.calls[0]![0]).toMatchObject({
      sessionId: 'session-1',
      sheetId: 'sh1',
      query: 'total',
      matchCase: false,
      matchEntireCell: false,
      lookIn: 'both',
      wildcards: false,
      limit: 51,
    })
  })

  it('truncates at maxResults and keeps regex queries on the range scan', async () => {
    const sidecarFind = vi.fn().mockResolvedValue({
      matches: [0, 1, 2].map((row) => ({
        sheetId: 'sh1',
        row,
        column: 0,
        value: 'total',
        valueText: 'total',
      })),
      complete: false,
      nextCursor: { sheetId: 'sh1', row: 3, column: 0 },
      indexingComplete: true,
    })
    vi.stubGlobal('window', { desktopApi: { findWorkbookCells: sidecarFind } })
    const capped = await findWorkbookCells_(options({ query: 'total', maxResults: 2 }))
    expect(capped.matches).toHaveLength(2)
    expect(capped.truncated).toBe(true)
    expect(sidecarFind.mock.calls[0]![0]).toMatchObject({ limit: 3 })

    sidecarFind.mockClear()
    await findWorkbookCells_(options({ query: 'total', lookIn: 'formulas' }))
    expect(sidecarFind.mock.calls[0]![0]).toMatchObject({ lookIn: 'formulas_only' })

    vi.mocked(readSheetRangeMapped).mockResolvedValue({
      screen: {
        cells: [{ row: 0, column: 0, value: 'total' }],
        rows: [],
        merges: [],
        hyperlinks: [],
      },
      raw: { indexingComplete: true },
      indexedThroughScreen: 3,
      fileEndRow: 3,
    } as never)
    sidecarFind.mockClear()
    const viaRegex = await findWorkbookCells_(options({ query: 'tot+al', regex: true }))
    expect(viaRegex.matches).toHaveLength(1)
    expect(sidecarFind).not.toHaveBeenCalled()
  })

  it('drops a withheld cell the sidecar returns', async () => {
    // The sidecar reads the file, so it is the one search path that never saw
    // the index until now. `shadowed` cannot help either: it only holds cells
    // edited this session, and a mark made weeks ago touches neither.
    const sidecarFind = vi.fn().mockResolvedValue({
      matches: [
        { sheetId: 'sh1', row: 1, column: 1, value: '13800138000', valueText: '13800138000' },
        { sheetId: 'sh1', row: 2, column: 1, value: '13800138001', valueText: '13800138001' },
      ],
      complete: true,
      indexingComplete: true,
    })
    vi.stubGlobal('window', { desktopApi: { findWorkbookCells: sidecarFind } })
    const ctx = {
      ...lazyCtx(lazyState(new Map()), [DATA_SHEET]),
      redactionIndexRef: {
        current: buildRedactionIndex(
          [
            {
              sheetName: 'Data',
              marks: [
                { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' },
              ],
            },
          ],
          [{ name: 'Data', id: 'sh1' }],
        ),
      },
    }
    const result = await findWorkbookCells(ctx, options({ query: '1380013800' }))
    // B2 is withheld, B3 holds the neighbouring number and is not: if both
    // vanished the search would be broken rather than merely quiet.
    expect(result.matches.map((m) => m.address)).toEqual(['B3'])
    expect(JSON.stringify(result)).not.toContain('13800138000')
  })

  function findWorkbookCells_(opts: FindCellsOptions) {
    return findWorkbookCells(lazyCtx(lazyState(new Map()), [DATA_SHEET]), opts)
  }
})

describe('find_cells treats a withheld cell as absent', () => {
  /**
   * The leak is not only the matched content. Reporting "Customers!B2 matches"
   * is an oracle: the model can test candidate values one after another and
   * learn which one the reader hid, without ever seeing a cell. A withheld
   * cell therefore never matches at all.
   */
  function ctxWithRedaction(): WorkbookReadContext {
    return {
      univerRef: { current: null },
      lazyWorkbookRef: { current: null },
      adapterRef: {
        current: {
          getSnapshot: () => ({
            revision: 0,
            sheets: [
              {
                id: 'sh1',
                name: 'Customers',
                cells: {
                  A1: { value: 'name' },
                  B1: { value: 'phone' },
                  A2: { value: 'Acme' },
                  B2: { value: '13800138000' },
                  B3: { value: '13800138001' },
                },
              },
            ],
          }),
        },
      },
      redactionIndexRef: {
        current: buildRedactionIndex(
          [
            {
              sheetName: 'Customers',
              marks: [
                { startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' },
              ],
            },
          ],
          [{ id: 'sh1', name: 'Customers' }],
        ),
      },
    } as unknown as WorkbookReadContext
  }

  it('does not confirm that a guessed value sits in a withheld cell', async () => {
    const result = await findWorkbookCells(ctxWithRedaction(), options({ query: '13800138000' }))
    expect(result.matches).toEqual([])
  })

  it('still finds the same value in a cell that is not withheld', async () => {
    // B3 holds the same number and is not marked: if it stopped matching too,
    // the search would be silently broken rather than merely quiet.
    const result = await findWorkbookCells(ctxWithRedaction(), options({ query: '1380013800' }))
    expect(result.matches.map((match) => match.address)).toEqual(['B3'])
  })

  it('never puts the withheld value in a result', async () => {
    const result = await findWorkbookCells(ctxWithRedaction(), options({ query: '1' }))
    expect(JSON.stringify(result)).not.toContain('13800138000')
  })
})

describe('find_cells on a streaming workbook treats a withheld cell as absent', () => {
  /// A mark over B2 on the `Data` fixture sheet.
  const withheldB2 = () =>
    buildRedactionIndex(
      [
        {
          sheetName: 'Data',
          marks: [{ startRow: 1, endRow: 1, startColumn: 1, endColumn: 1, label: 'client phone' }],
        },
      ],
      [{ id: 'sh1', name: 'Data' }],
    )

  function lazyRedactionCtx(state: unknown): WorkbookReadContext {
    return {
      ...lazyCtx(state, [DATA_SHEET]),
      redactionIndexRef: { current: withheldB2() },
    } as unknown as WorkbookReadContext
  }

  it('does not confirm a guessed value sitting in a streamed cell', async () => {
    vi.mocked(readSheetRangeMapped).mockResolvedValue({
      screen: {
        cells: [
          { row: 1, column: 0, value: 'findable' },
          { row: 1, column: 1, value: '13800138000' },
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
      lazyRedactionCtx(lazyState(new Map())),
      options({ query: '13800138000' }),
    )
    expect(result.matches).toEqual([])
  })

  it('still searches the cells around a withheld one', async () => {
    vi.mocked(readSheetRangeMapped).mockResolvedValue({
      screen: {
        cells: [
          { row: 1, column: 0, value: 'findable' },
          { row: 1, column: 1, value: '13800138000' },
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
      lazyRedactionCtx(lazyState(new Map())),
      options({ query: 'findable' }),
    )
    expect(result.matches.map((match) => match.address)).toEqual(['A2'])
  })

  it('does not confirm a guessed value in a cell written this session', async () => {
    vi.mocked(readSheetRangeMapped).mockResolvedValue({
      screen: { cells: [], rows: [], merges: [], hyperlinks: [] },
      raw: { indexingComplete: true },
      indexedThroughScreen: 3,
      fileEndRow: 3,
    } as never)
    const journal = new Map([
      ['sh1', new Map([['1:1', { row: 1, column: 1, hasValue: true, value: '13800138000' }]])],
    ])

    const result = await findWorkbookCells(
      lazyRedactionCtx(lazyState(journal)),
      options({ query: '13800138000' }),
    )
    expect(result.matches).toEqual([])
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
