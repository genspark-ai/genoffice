import { afterEach, describe, expect, it, vi } from 'vitest'

import { traceWorkbookDependents, traceWorkbookPrecedents } from '../src/renderer/ai/formula-audit'
import type { WorkbookReadContext } from '../src/renderer/ai/workbook-readers'

vi.mock('../src/renderer/univer-sync', () => ({
  readSheetRangeMapped: vi.fn(),
  ensureLazyRangeLoaded: vi.fn(),
}))

afterEach(() => {
  vi.unstubAllGlobals()
  vi.clearAllMocks()
})

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

const DEMO: DemoSheet[] = [
  {
    id: 's1',
    name: 'Sheet1',
    cells: {
      A1: { value: 10 },
      A2: { value: 0 },
      B1: { value: '#DIV/0!', formula: '=A1/A2' },
      C1: { value: '#DIV/0!', formula: '=B1*2' },
      D1: { value: 'text' },
      E1: { value: 10, formula: '=SUM(A:A)' },
    },
  },
  {
    id: 's2',
    name: 'Summary',
    cells: { A1: { value: '#DIV/0!', formula: '=Sheet1!B1' } },
  },
]

describe('traceWorkbookPrecedents: demo workbook', () => {
  it('lists single-cell precedents with values', async () => {
    const result = await traceWorkbookPrecedents(demoCtx(DEMO), undefined, 'B1')
    expect(result.formula).toBe('=A1/A2')
    expect(result.refs.map((ref) => ref.label)).toEqual(['A1', 'A2'])
    expect(result.refs[0]?.samples[0]).toMatchObject({ address: 'A1', value: 10 })
    expect(result.refs.every((ref) => !ref.hasError)).toBe(true)
  })

  it('flags precedents that hold error values', async () => {
    const result = await traceWorkbookPrecedents(demoCtx(DEMO), undefined, 'C1')
    expect(result.refs).toHaveLength(1)
    expect(result.refs[0]).toMatchObject({ label: 'B1', hasError: true })
  })

  it('reports non-formula cells without refs', async () => {
    const result = await traceWorkbookPrecedents(demoCtx(DEMO), undefined, 'D1')
    expect(result.formula).toBeUndefined()
    expect(result.value).toBe('text')
    expect(result.refs).toEqual([])
  })

  it('clamps whole-column references to the data extent', async () => {
    const result = await traceWorkbookPrecedents(demoCtx(DEMO), undefined, 'E1')
    expect(result.refs[0]?.label).toBe('A1:A2')
    expect(result.refs[0]?.cellCount).toBe(2)
  })

  it('resolves sheet-qualified references', async () => {
    const result = await traceWorkbookPrecedents(demoCtx(DEMO), 's2', 'A1')
    expect(result.refs[0]).toMatchObject({ label: 'Sheet1!B1', hasError: true })
  })

  it('surfaces defined-name usage it cannot expand', async () => {
    const sheets: DemoSheet[] = [
      { id: 's1', name: 'Sheet1', cells: { A1: { value: 5, formula: '=Total*2' } } },
    ]
    const result = await traceWorkbookPrecedents(demoCtx(sheets), undefined, 'A1')
    expect(result.usesNames).toBe(true)
    expect(result.refs).toEqual([])
  })
})

describe('traceWorkbookDependents: demo workbook', () => {
  it('finds direct and whole-column dependents on the same sheet', async () => {
    const result = await traceWorkbookDependents(demoCtx(DEMO), undefined, 'A2')
    expect(result.dependents.map((dep) => `${dep.sheetName}!${dep.address}`)).toEqual([
      'Sheet1!B1',
      'Sheet1!E1',
    ])
  })

  it('finds cross-sheet dependents through sheet qualifiers', async () => {
    const result = await traceWorkbookDependents(demoCtx(DEMO), undefined, 'B1')
    expect(result.dependents.map((dep) => `${dep.sheetName}!${dep.address}`)).toEqual([
      'Sheet1!C1',
      'Summary!A1',
    ])
  })

  it('rejects an unknown sheet', async () => {
    const result = await traceWorkbookDependents(demoCtx(DEMO), 'ghost', 'A1')
    expect(result.error).toContain('Unknown sheet: ghost')
  })
})

describe('traceWorkbookDependents: lazy workbook', () => {
  function lazyCtx(
    journalCells: Map<string, Map<string, unknown>>,
    structuralOps: Map<string, unknown[]> = new Map(),
  ): WorkbookReadContext {
    const worksheet = { getSheetId: () => 'sh1', getSheetName: () => 'Data' }
    return {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              getActiveSheet: () => worksheet,
              getSheetBySheetId: () => null,
            }),
          },
        },
      },
      lazyWorkbookRef: {
        current: {
          file: {
            sessionId: 'session-1',
            sheets: [{ id: 'sh1', name: 'Data', rowCount: 5, columnCount: 3 }],
          },
          editJournal: { cells: journalCells, structuralOps },
        },
      },
      adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } },
    } as unknown as WorkbookReadContext
  }

  it('merges file formulas with journal overrides and flags incomplete indexing', async () => {
    const journal = new Map([
      [
        'sh1',
        new Map([
          // C1 overwritten this session: no longer a dependent of A1
          ['0:2', { row: 0, column: 2, hasValue: true, value: 7 }],
          // C2 written this session as a new dependent
          ['1:2', { row: 1, column: 2, hasValue: true, value: null, formula: '=A1*3' }],
        ]),
      ],
    ])
    vi.stubGlobal('window', {
      desktopApi: {
        readWorkbookFormulas: vi.fn().mockResolvedValue({
          cells: [
            { row: 0, column: 2, value: 20, formula: '=A1*2' },
            { row: 2, column: 2, value: 30, formula: '=A1+B1' },
            { row: 3, column: 2, value: 1, formula: '=B2' },
          ],
          indexingComplete: false,
          truncated: false,
        }),
      },
    })
    const result = await traceWorkbookDependents(lazyCtx(journal), undefined, 'A1')
    expect(result.dependents.map((dep) => dep.address)).toEqual(['C2', 'C3'])
    expect(result.incompleteSheets).toEqual(['Data'])
  })

  it('maps file-formula refs through journaled structural ops before matching', async () => {
    // One row inserted at the top: file formula "=A1" at file C1 now sits at
    // screen C2 and semantically reads screen A2
    const ops = new Map([['sh1', [{ kind: 'insert-rows', index: 0, count: 1 }]]])
    vi.stubGlobal('window', {
      desktopApi: {
        readWorkbookFormulas: vi.fn().mockResolvedValue({
          cells: [{ row: 0, column: 2, value: 20, formula: '=A1' }],
          indexingComplete: true,
          truncated: false,
        }),
      },
    })
    const hit = await traceWorkbookDependents(lazyCtx(new Map(), ops), undefined, 'A2')
    expect(hit.dependents.map((dep) => dep.address)).toEqual(['C2'])
    const miss = await traceWorkbookDependents(lazyCtx(new Map(), ops), undefined, 'A1')
    expect(miss.dependents).toEqual([])
  })

  it('keeps whole-column refs covering rows inserted this session', async () => {
    const ops = new Map([['sh1', [{ kind: 'insert-rows', index: 0, count: 1 }]]])
    vi.stubGlobal('window', {
      desktopApi: {
        readWorkbookFormulas: vi.fn().mockResolvedValue({
          cells: [{ row: 0, column: 2, value: 20, formula: '=SUM(A:A)' }],
          indexingComplete: true,
          truncated: false,
        }),
      },
    })
    // Screen A1 is the inserted row — A:A semantically covers it
    const result = await traceWorkbookDependents(lazyCtx(new Map(), ops), undefined, 'A1')
    expect(result.dependents.map((dep) => dep.address)).toEqual(['C2'])
  })

  it('matches file-formula qualifiers against pre-rename sheet names, never session-added sheets', async () => {
    const sheet1 = { getSheetId: () => 'sh1', getSheetName: () => 'Data' }
    // Renamed this session: screen name differs from the file name
    const sheet2 = { getSheetId: () => 'sh2', getSheetName: () => 'Renamed' }
    // Session-added sheet reusing the renamed sheet's old name
    const sheet3 = { getSheetId: () => 'sh3', getSheetName: () => 'Old' }
    const ctx = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [sheet1, sheet2, sheet3],
              getActiveSheet: () => sheet1,
              getSheetBySheetId: () => null,
            }),
          },
        },
      },
      lazyWorkbookRef: {
        current: {
          file: {
            sessionId: 'session-1',
            sheets: [
              { id: 'sh1', name: 'Data', rowCount: 5, columnCount: 3 },
              { id: 'sh2', name: 'Old', rowCount: 5, columnCount: 3 },
            ],
          },
          editJournal: { cells: new Map(), structuralOps: new Map() },
        },
      },
      adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } },
    } as unknown as WorkbookReadContext
    vi.stubGlobal('window', {
      desktopApi: {
        readWorkbookFormulas: vi.fn().mockImplementation(({ sheetId }: { sheetId: string }) =>
          Promise.resolve({
            cells: sheetId === 'sh1' ? [{ row: 0, column: 2, value: 1, formula: '=Old!A1' }] : [],
            indexingComplete: true,
            truncated: false,
          }),
        ),
      },
    })
    const renamed = await traceWorkbookDependents(ctx, 'sh2', 'A1')
    expect(renamed.dependents.map((dep) => `${dep.sheetName}!${dep.address}`)).toEqual(['Data!C1'])
    // The file "=Old!A1" refers to the renamed sheet, not the new one that took its name
    const recreated = await traceWorkbookDependents(ctx, 'sh3', 'A1')
    expect(recreated.dependents).toEqual([])
  })
})

describe('traceWorkbookPrecedents: lazy workbook with structural ops', () => {
  it('re-opens whole-column refs to the screen extent after mapping', async () => {
    const ops = new Map([['sh1', [{ kind: 'insert-rows', index: 0, count: 1 }]]])
    const worksheet = { getSheetId: () => 'sh1', getSheetName: () => 'Data' }
    const ctx = {
      univerRef: {
        current: {
          univerAPI: {
            getActiveWorkbook: () => ({
              getSheets: () => [worksheet],
              getActiveSheet: () => worksheet,
              getSheetBySheetId: () => null,
            }),
          },
        },
      },
      lazyWorkbookRef: {
        current: {
          file: {
            sessionId: 'session-1',
            sheets: [{ id: 'sh1', name: 'Data', rowCount: 5, columnCount: 3 }],
          },
          editJournal: { cells: new Map(), structuralOps: ops },
        },
      },
      adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } },
    } as unknown as WorkbookReadContext
    const { readSheetRangeMapped } = await import('../src/renderer/univer-sync')
    vi.mocked(readSheetRangeMapped).mockResolvedValue({
      screen: {
        cells: [{ row: 1, column: 2, value: 20, formula: '=SUM(A:A)' }],
        rows: [],
        merges: [],
        hyperlinks: [],
      },
      raw: { indexingComplete: true },
      indexedThroughScreen: 5,
      fileEndRow: 4,
    } as never)
    const result = await traceWorkbookPrecedents(ctx, undefined, 'C2')
    // 5 file rows + 1 inserted = 6 screen rows: the open axis spans A1:A6
    expect(result.refs[0]?.label).toBe('A1:A6')
    expect(result.refs[0]?.cellCount).toBe(6)
  })
})
