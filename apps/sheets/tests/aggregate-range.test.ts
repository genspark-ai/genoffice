import { afterEach, describe, expect, it, vi } from 'vitest'

import { parseRange } from '../src/domain/cell-address'
import { aggregateWorkbookRange } from '../src/renderer/ai/aggregate-range'
import type { WorkbookReadContext } from '../src/renderer/ai/workbook-readers'

afterEach(() => {
  vi.unstubAllGlobals()
  vi.clearAllMocks()
})

type DemoSheet = {
  id: string
  name: string
  cells: Record<string, { value: string | number | boolean | null; formula?: string }>
}

type UniverStub = Record<string, { getValue(): unknown }>

function univerRef(sheetValues: Record<string, UniverStub>, activeSheetId = 'sheet-1') {
  return {
    current: {
      univerAPI: {
        getActiveWorkbook: () => ({
          getActiveSheet: () => ({ getSheetId: () => activeSheetId }),
          getSheetBySheetId: (id: string) => {
            const values = sheetValues[id]
            if (!values) return null
            return { getRange: (address: string) => values[address] ?? { getValue: () => null } }
          },
        }),
      },
    },
  }
}

function demoCtx(
  sheets: DemoSheet[],
  univer: Record<string, UniverStub> = {},
): WorkbookReadContext {
  return {
    univerRef: univerRef(univer),
    lazyWorkbookRef: { current: null },
    adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets }) } },
  } as unknown as WorkbookReadContext
}

type JournalEntryStub = {
  row: number
  column: number
  hasValue: boolean
  value: string | number | boolean | null
  formula?: string
}

function lazyCtx(options: {
  rowCount: number
  columnCount: number
  journal?: JournalEntryStub[]
  univer?: Record<string, UniverStub>
}): WorkbookReadContext {
  const entries = new Map(
    (options.journal ?? []).map((entry) => [`${entry.row}:${entry.column}`, entry]),
  )
  return {
    univerRef: univerRef(options.univer ?? {}),
    lazyWorkbookRef: {
      current: {
        file: {
          sessionId: 'session-1',
          sheets: [
            {
              id: 'sheet-1',
              name: 'Data',
              rowCount: options.rowCount,
              columnCount: options.columnCount,
            },
          ],
        },
        editJournal: {
          cells: new Map([['sheet-1', entries]]),
          structuralOps: new Map(),
          sheets: { added: new Set(), removed: new Set() },
        },
      },
    },
    adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } },
  } as unknown as WorkbookReadContext
}

function stubReadWorkbookRange(
  cells: { row: number; column: number; value: string | number | boolean | null }[],
) {
  const readWorkbookRange = vi
    .fn()
    .mockImplementation(({ range }: { range: { startRow: number; endRow: number } }) =>
      Promise.resolve({
        cells: cells.filter((cell) => cell.row >= range.startRow && cell.row <= range.endRow),
        indexedThroughRow: Number.MAX_SAFE_INTEGER,
      }),
    )
  vi.stubGlobal('window', { desktopApi: { readWorkbookRange } })
  return readWorkbookRange
}

describe('aggregateWorkbookRange: demo workbook', () => {
  const SHEETS: DemoSheet[] = [
    { id: 'sheet-1', name: 'Sheet1', cells: { A1: { value: 1 }, A2: { value: 2 } } },
    { id: 'sheet-2', name: 'Summary', cells: { A1: { value: 100 } } },
  ]

  it('errors on an unknown sheetId instead of falling back to the first sheet', async () => {
    const result = await aggregateWorkbookRange(demoCtx(SHEETS), 'sheet-9', parseRange('A1:A10'))
    expect(result).toEqual({
      ok: false,
      error: 'Unknown sheet: sheet-9 (use an id from get_workbook_context)',
    })
  })

  it('aggregates the explicitly requested sheet', async () => {
    const result = await aggregateWorkbookRange(demoCtx(SHEETS), 'sheet-2', parseRange('A1:A10'))
    expect(result.ok).toBe(true)
    if (result.ok) expect(result.aggregate.sum).toBe(100)
  })

  it('backfills formula cells (value:null) from the Univer grid', async () => {
    const sheets: DemoSheet[] = [
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: { A1: { value: 1 }, A2: { value: null, formula: '=A1*2' } },
      },
    ]
    const result = await aggregateWorkbookRange(
      demoCtx(sheets, { 'sheet-1': { A2: { getValue: () => 2 } } }),
      undefined,
      parseRange('A1:A2'),
    )
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.aggregate.nonEmpty).toBe(2)
      expect(result.aggregate.sum).toBe(3)
    }
  })
})

describe('aggregateWorkbookRange: lazy workbook', () => {
  it('includes journaled rows appended below the file extent', async () => {
    stubReadWorkbookRange([
      { row: 0, column: 0, value: 10 },
      { row: 1, column: 0, value: 20 },
    ])
    const ctx = lazyCtx({
      rowCount: 2,
      columnCount: 1,
      journal: [
        { row: 2, column: 0, hasValue: true, value: 30 },
        { row: 3, column: 0, hasValue: true, value: 40 },
      ],
    })
    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A10'))
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.aggregate.nonEmpty).toBe(4)
      expect(result.aggregate.sum).toBe(100)
      expect(result.aggregate.cells).toBe(4)
    }
  })

  it('aggregates a range that lies entirely below the file extent', async () => {
    const readWorkbookRange = stubReadWorkbookRange([])
    const ctx = lazyCtx({
      rowCount: 2,
      columnCount: 1,
      journal: [{ row: 4, column: 0, hasValue: true, value: 99 }],
    })
    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A4:A6'))
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.aggregate.nonEmpty).toBe(1)
      expect(result.aggregate.sum).toBe(99)
    }
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })

  it('still errors when the range is beyond both the file and session extents', async () => {
    stubReadWorkbookRange([])
    const ctx = lazyCtx({ rowCount: 2, columnCount: 1 })
    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A5:A9'))
    expect(result).toMatchObject({ ok: false })
    if (!result.ok) expect(result.error).toContain('outside the sheet data extent')
  })

  it('backfills journal formula entries (value:null) from the Univer grid', async () => {
    stubReadWorkbookRange([{ row: 0, column: 0, value: 5 }])
    const ctx = lazyCtx({
      rowCount: 1,
      columnCount: 1,
      journal: [{ row: 1, column: 0, hasValue: true, value: null, formula: '=A1*3' }],
      univer: { 'sheet-1': { A2: { getValue: () => 15 } } },
    })
    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A2'))
    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.aggregate.nonEmpty).toBe(2)
      expect(result.aggregate.sum).toBe(20)
    }
  })
})
