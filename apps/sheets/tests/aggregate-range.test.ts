import { afterEach, describe, expect, it, vi } from 'vitest'

import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import { aggregateWorkbookRange } from '../src/renderer/ai/aggregate-range'
import type { WorkbookReadContext } from '../src/renderer/ai/workbook-readers'
import { NO_REDACTIONS, buildRedactionIndex, type RedactionIndex } from '../src/renderer/ai/redact'
import type { SheetRedactionState } from '@genoffice/xlsx-gateway/gateway/xlsx-redaction'

/// A mark over a cell of the fixture sheet, which the file calls `Data`.
function markOnData(
  startRow: number,
  endRow: number,
  startColumn: number,
  endColumn: number,
  label = 'secret',
): RedactionIndex {
  const states: SheetRedactionState[] = [
    { sheetName: 'Data', marks: [{ startRow, endRow, startColumn, endColumn, label }] },
  ]
  return buildRedactionIndex(states, [{ id: 'sheet-1', name: 'Data' }])
}

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
  style?: Record<string, unknown>
}

type BulkFillStub = {
  startRow: number
  endRow: number
  startColumn: number
  endColumn: number
  value: string | number | boolean | null
}

function lazyCtx(options: {
  rowCount: number
  columnCount: number
  journal?: JournalEntryStub[]
  fills?: BulkFillStub[]
  univer?: Record<string, UniverStub>
  redactions?: RedactionIndex
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
          bulkConstantFills: new Map([['sheet-1', options.fills ?? []]]),
          structuralOps: new Map(),
          sheets: { added: new Set(), removed: new Set() },
        },
      },
    },
    adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } },
    redactionIndexRef: { current: options.redactions ?? NO_REDACTIONS },
  } as unknown as WorkbookReadContext
}

function stubReadWorkbookRange(
  cells: { row: number; column: number; value: string | number | boolean | null }[],
  hiddenRows: number[] = [],
) {
  const readWorkbookRange = vi
    .fn()
    .mockImplementation(
      ({
        range,
      }: {
        range: { startRow: number; endRow: number; startColumn: number; endColumn: number }
      }) =>
        Promise.resolve({
          cells: cells.filter(
            (cell) =>
              cell.row >= range.startRow &&
              cell.row <= range.endRow &&
              cell.column >= range.startColumn &&
              cell.column <= range.endColumn,
          ),
          rows: hiddenRows
            .filter((row) => row >= range.startRow && row <= range.endRow)
            .map((row) => ({ row, hidden: true })),
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
  it('drops fill-band cells on file-hidden rows and still reads fully filled batches for row properties', async () => {
    const readWorkbookRange = stubReadWorkbookRange(
      [
        { row: 0, column: 0, value: 1 },
        { row: 3, column: 0, value: 4 },
      ],
      [1],
    )
    const ctx = lazyCtx({
      rowCount: 4,
      columnCount: 1,
      fills: [{ startRow: 0, endRow: 3, startColumn: 0, endColumn: 0, value: 10 }],
    })

    const all = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A4'))
    expect(all.ok && all.aggregate.sum).toBe(40)
    expect(readWorkbookRange).not.toHaveBeenCalled()

    const visible = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A4'), {
      skipFileHiddenRows: true,
    })
    expect(readWorkbookRange).toHaveBeenCalledOnce()
    expect(visible.ok).toBe(true)
    if (visible.ok) {
      expect(visible.aggregate.cells).toBe(3)
      expect(visible.aggregate.nonEmpty).toBe(3)
      expect(visible.aggregate.sum).toBe(30)
    }
  })

  it('skips file-hidden rows only when asked, using the per-batch row properties', async () => {
    const cells = [
      { row: 0, column: 0, value: 1 },
      { row: 1, column: 0, value: 2 },
      { row: 2, column: 0, value: 3 },
      { row: 3, column: 0, value: 4 },
    ]
    stubReadWorkbookRange(cells, [1, 3])
    const ctx = lazyCtx({ rowCount: 4, columnCount: 1, fills: [] })

    const all = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A4'))
    expect(all.ok && all.aggregate.sum).toBe(10)

    const visible = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A4'), {
      skipFileHiddenRows: true,
    })
    expect(visible.ok).toBe(true)
    if (visible.ok) {
      expect(visible.aggregate.cells).toBe(2)
      expect(visible.aggregate.nonEmpty).toBe(2)
      expect(visible.aggregate.numericCount).toBe(2)
      expect(visible.aggregate.sum).toBe(4)
    }
  })

  it('overlays bulk fills on streamed file values while retaining uncovered cells', async () => {
    const readWorkbookRange = stubReadWorkbookRange([
      { row: 0, column: 0, value: 1 },
      { row: 1, column: 0, value: 2 },
      { row: 2, column: 0, value: 3 },
      { row: 3, column: 0, value: 4 },
    ])
    const ctx = lazyCtx({
      rowCount: 4,
      columnCount: 1,
      fills: [
        {
          startRow: 1,
          endRow: 2,
          startColumn: 0,
          endColumn: 0,
          value: 10,
        },
      ],
    })

    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A4'))

    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.aggregate.cells).toBe(4)
      expect(result.aggregate.nonEmpty).toBe(4)
      expect(result.aggregate.sum).toBe(25)
      expect(result.aggregate.topValues).toContainEqual({ value: '10', count: 2 })
    }
    expect(readWorkbookRange).toHaveBeenCalledOnce()
  })

  it('aggregates a fully covered million-cell fill without reading or expanding cells', async () => {
    const readWorkbookRange = stubReadWorkbookRange([])
    const ctx = lazyCtx({
      rowCount: 1_000_000,
      columnCount: 1,
      fills: [
        {
          startRow: 0,
          endRow: 999_999,
          startColumn: 0,
          endColumn: 0,
          value: 3,
        },
      ],
    })

    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A1000000'))

    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.aggregate).toMatchObject({
        cells: 1_000_000,
        nonEmpty: 1_000_000,
        numericCount: 1_000_000,
        sum: 3_000_000,
        average: 3,
        topValues: [{ value: '3', count: 1_000_000 }],
      })
    }
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })

  it('uses later overlapping fills for their disjoint winning spans', async () => {
    const readWorkbookRange = stubReadWorkbookRange([])
    const ctx = lazyCtx({
      rowCount: 5,
      columnCount: 1,
      fills: [
        {
          startRow: 0,
          endRow: 4,
          startColumn: 0,
          endColumn: 0,
          value: 'base',
        },
        {
          startRow: 1,
          endRow: 3,
          startColumn: 0,
          endColumn: 0,
          value: 'later',
        },
      ],
    })

    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A5'))

    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.aggregate.topValues).toEqual([
        { value: 'later', count: 3 },
        { value: 'base', count: 2 },
      ])
    }
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })

  it('lets explicit scalar, formula, and empty edits override fills but ignores style-only edits', async () => {
    const readWorkbookRange = stubReadWorkbookRange([])
    const ctx = lazyCtx({
      rowCount: 5,
      columnCount: 1,
      fills: [
        {
          startRow: 0,
          endRow: 4,
          startColumn: 0,
          endColumn: 0,
          value: 10,
        },
      ],
      journal: [
        { row: 0, column: 0, hasValue: true, value: 2 },
        { row: 1, column: 0, hasValue: true, value: null, formula: '=A1+5' },
        { row: 2, column: 0, hasValue: true, value: null },
        { row: 3, column: 0, hasValue: true, value: '' },
        { row: 4, column: 0, hasValue: false, value: null, style: { bold: true } },
      ],
      univer: { 'sheet-1': { A2: { getValue: () => 7 } } },
    })

    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A5'))

    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.aggregate.cells).toBe(5)
      expect(result.aggregate.nonEmpty).toBe(3)
      expect(result.aggregate.numericCount).toBe(3)
      expect(result.aggregate.sum).toBe(19)
      expect(result.aggregate.distinct).toBe(3)
      expect(result.aggregate.topValues).toEqual(
        expect.arrayContaining([
          { value: '10', count: 1 },
          { value: '2', count: 1 },
          { value: '7', count: 1 },
        ]),
      )
    }
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })

  it('uses bulk fills to expand the session data extent', async () => {
    const readWorkbookRange = stubReadWorkbookRange([])
    const ctx = lazyCtx({
      rowCount: 2,
      columnCount: 1,
      fills: [
        {
          startRow: 4,
          endRow: 5,
          startColumn: 0,
          endColumn: 0,
          value: 'appended',
        },
      ],
    })

    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A4:A6'))

    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.aggregate.cells).toBe(3)
      expect(result.aggregate.nonEmpty).toBe(2)
      expect(result.aggregate.topValues).toEqual([{ value: 'appended', count: 2 }])
    }
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })

  it('counts null and empty-string fills as cells but empty aggregate values', async () => {
    const readWorkbookRange = stubReadWorkbookRange([])
    const ctx = lazyCtx({
      rowCount: 4,
      columnCount: 1,
      fills: [
        {
          startRow: 0,
          endRow: 1,
          startColumn: 0,
          endColumn: 0,
          value: null,
        },
        {
          startRow: 2,
          endRow: 3,
          startColumn: 0,
          endColumn: 0,
          value: '',
        },
      ],
    })

    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A4'))

    expect(result.ok).toBe(true)
    if (result.ok) {
      expect(result.aggregate.cells).toBe(4)
      expect(result.aggregate.nonEmpty).toBe(0)
      expect(result.aggregate.distinct).toBe(0)
      expect(result.aggregate.topValues).toEqual([])
    }
    expect(readWorkbookRange).not.toHaveBeenCalled()
  })

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

describe('a withheld cell is kept out of the arithmetic, whichever source it came from', () => {
  /**
   * `sum` is exact, so one withheld cell inside a column is solvable by
   * subtraction. Masking the cell in the output cannot help — the value is
   * recoverable from the total regardless. The cell has to be excluded from
   * the sum itself, which means every source that can contribute a value has
   * to route it away. There are three, and missing one leaks.
   */

  it('for a cell streamed from the file', async () => {
    stubReadWorkbookRange([
      { row: 0, column: 0, value: 10 },
      { row: 1, column: 0, value: 999_999 },
      { row: 2, column: 0, value: 40 },
    ])
    const ctx = lazyCtx({
      rowCount: 3,
      columnCount: 1,
      redactions: markOnData(1, 1, 0, 0, 'client phone'),
    })

    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A3'))

    expect(result.ok).toBe(true)
    if (!result.ok) return
    // Without the routing the sum would be 1000049 and 999_999 recoverable.
    expect(result.aggregate.sum).toBe(50)
    expect(result.aggregate.max).toBe(40)
    expect(result.aggregate.numericCount).toBe(2)
    expect(result.aggregate.withheld).toBe(1)
    expect(result.aggregate.cells).toBe(3)
    expect(JSON.stringify(result.aggregate)).not.toContain('999999')
  })

  it('for a cell covered by a bulk fill', async () => {
    // A fill band is one value repeated over a rectangle, so the leak is a
    // repetition count, not a single cell: without the subtraction the band
    // would add the reader's own value once per row it covers.
    stubReadWorkbookRange([])
    const ctx = lazyCtx({
      rowCount: 5,
      columnCount: 1,
      fills: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 0, value: 7 }],
      redactions: markOnData(1, 3, 0, 0, 'bulk'),
    })

    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A5'))

    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.aggregate.withheld).toBe(3)
    expect(result.aggregate.sum).toBe(14) // rows 0 and 4 only
    expect(result.aggregate.cells).toBe(5)
    expect(result.aggregate.topValues).toEqual([{ value: '7', count: 2 }])
  })

  it('for a cell written this session', async () => {
    stubReadWorkbookRange([
      { row: 0, column: 0, value: 10 },
      { row: 1, column: 0, value: 999_999 },
      { row: 2, column: 0, value: 40 },
    ])
    const ctx = lazyCtx({
      rowCount: 3,
      columnCount: 1,
      journal: [
        { row: 1, column: 0, hasValue: true, value: 999_999 },
        { row: 2, column: 0, hasValue: true, value: 40 },
      ],
      redactions: markOnData(1, 1, 0, 0, 'fresh'),
    })

    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A3'))

    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.aggregate.sum).toBe(50)
    expect(result.aggregate.withheld).toBe(1)
    expect(JSON.stringify(result.aggregate)).not.toContain('999999')
  })

  it('reports a withheld range that is entirely empty of statistics', async () => {
    // Every cell in the range withheld: the answer is "3 cells, all withheld",
    // never a total the model could subtract its way back to the values.
    stubReadWorkbookRange([
      { row: 0, column: 0, value: 5 },
      { row: 1, column: 0, value: 6 },
      { row: 2, column: 0, value: 7 },
    ])
    const ctx = lazyCtx({
      rowCount: 3,
      columnCount: 1,
      redactions: markOnData(0, 2, 0, 0, 'whole column'),
    })

    const result = await aggregateWorkbookRange(ctx, 'sheet-1', parseRange('A1:A3'))

    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.aggregate.withheld).toBe(3)
    expect(result.aggregate.sum).toBe(0)
    expect(result.aggregate.numericCount).toBe(0)
    expect(result.aggregate.topValues).toEqual([])
  })
})

describe('the demo-workbook path withholds too', () => {
  it('keeps a withheld cell out of a snapshot-backed aggregate', async () => {
    // The demo workbook never goes through the streaming reader, so it has its
    // own routing. Missing it would leave a whole code path unprotected.
    const redactions = buildRedactionIndex(
      [
        {
          sheetName: 'Demo',
          marks: [{ startRow: 1, endRow: 1, startColumn: 0, endColumn: 0, label: 'demo secret' }],
        },
      ],
      [{ id: 'demo-1', name: 'Demo' }],
    )
    const ctx: WorkbookReadContext = {
      univerRef: univerRef({}),
      lazyWorkbookRef: { current: null },
      adapterRef: {
        current: {
          getSnapshot: () => ({
            revision: 0,
            sheets: [
              {
                id: 'demo-1',
                name: 'Demo',
                cells: {
                  A1: { value: 10 },
                  A2: { value: 999_999 },
                  A3: { value: 40 },
                },
              },
            ],
          }),
        },
      },
      redactionIndexRef: { current: redactions },
    } as unknown as WorkbookReadContext

    const result = await aggregateWorkbookRange(ctx, 'demo-1', parseRange('A1:A3'))

    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.aggregate.sum).toBe(50)
    expect(result.aggregate.withheld).toBe(1)
    expect(result.aggregate.cells).toBe(3)
    expect(JSON.stringify(result.aggregate)).not.toContain('999999')
  })
})
