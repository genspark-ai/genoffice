import { describe, expect, it, vi } from 'vitest'
import { parseRange } from '@genoffice/xlsx-gateway/domain/cell-address'
import {
  getSourceRange,
  handleCreatePivot,
  pivotFieldOptions,
  type PivotActionContext,
} from '../src/renderer/pivot-actions'
import type { OoXmlPivotConfig } from '../src/renderer/PivotDialog'
import { applyAiPivotAdd } from '../src/renderer/workbook-ops'
import { createEditJournal } from '../src/renderer/edit-journal'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

vi.mock('../src/renderer/workbook-ops', async (importOriginal) => ({
  ...(await importOriginal<typeof import('../src/renderer/workbook-ops')>()),
  applyAiPivotAdd: vi.fn(),
}))

type TestRange = {
  getRow(): number
  getColumn(): number
  getHeight(): number
  getWidth(): number
  getValues(): unknown[][]
  getValue(): unknown
}

function fixture(
  cells: unknown[][] = [
    ['Region', 'Revenue', 'Units'],
    ['East', 10, 2],
    ['West', 20, 3],
  ],
) {
  const reads: { row: number; height: number }[] = []
  const getRange = (row: number | string, col = 0, height = 1, width = 1): TestRange => {
    if (typeof row === 'string') {
      const b = parseRange(row)
      return getRange(
        b.startRow,
        b.startColumn,
        b.endRow - b.startRow + 1,
        b.endColumn - b.startColumn + 1,
      )
    }
    const values = () => {
      reads.push({ row, height })
      return Array.from({ length: height }, (_, r) =>
        Array.from({ length: width }, (_, c) => cells[row + r]?.[col + c] ?? null),
      )
    }
    return {
      getRow: () => row,
      getColumn: () => col,
      getHeight: () => height,
      getWidth: () => width,
      getValues: values,
      getValue: () => values()[0]?.[0],
    }
  }
  let selection = getRange(1, 1)
  const worksheet = {
    getSheetId: () => 's1',
    getRange,
    getLastRow: () => cells.length - 1,
    getLastColumn: () => Math.max(...cells.map((r) => r.length)) - 1,
  }
  const workbook = { getActiveSheet: () => worksheet, getActiveRange: () => selection }
  const state = {
    flags: { preloadComplete: true },
    file: { sheets: [{ id: 's1', tables: [] }] },
    editJournal: createEditJournal(),
  } as unknown as LazyWorkbookState
  const ctx = {
    univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
    lazyWorkbookRef: { current: state },
    setPendingEdits: vi.fn(),
    setMessage: vi.fn(),
  } as unknown as PivotActionContext
  return {
    ctx,
    state,
    reads,
    select: (ref: string) => {
      selection = getRange(ref)
    },
  }
}

function addTable(f: ReturnType<typeof fixture>, ref: string, totalsRowCount = 0) {
  f.state.file.sheets[0]!.tables.push({
    range: parseRange(ref),
    headerRowCount: 1,
    totalsRowCount,
    showRowStripes: false,
    showColumnStripes: false,
  })
}

describe('pivot source selection', () => {
  it('uses the whole continuous region for a single data cell and its headers', () => {
    const { ctx, reads } = fixture()
    const source = getSourceRange(ctx)
    expect(source).toBe('A1:C3')
    reads.length = 0
    expect(pivotFieldOptions(ctx, source)).toEqual([
      { label: 'Region', colIndex: 0 },
      { label: 'Revenue', colIndex: 1 },
      { label: 'Units', colIndex: 2 },
    ])
    expect(reads).toEqual([{ row: 0, height: 1 }])
    // The subtotal dialog still uses the original single-cell selection.
    expect(pivotFieldOptions(ctx)).toEqual([])
  })

  it('keeps explicit selections, even when they cut across a larger region', () => {
    const f = fixture()
    f.select('B1:C3')
    expect(getSourceRange(f.ctx)).toBe('B1:C3')
    expect(pivotFieldOptions(f.ctx, 'B1:C3')).toEqual([
      { label: 'Revenue', colIndex: 1 },
      { label: 'Units', colIndex: 2 },
    ])
    f.select('A1:C1')
    expect(getSourceRange(f.ctx)).toBe('A1:C1')
    expect(pivotFieldOptions(f.ctx, 'A1:C1')).toEqual([])
  })

  it('stops at blank rows and columns and retains absolute column indices', () => {
    const f = fixture([
      [null, 'Name', 'Amount', null, 'Other'],
      [null, 'A', 0, null, 99],
      [],
      [null, 'B', 12],
    ])
    f.select('C2')
    expect(getSourceRange(f.ctx)).toBe('B1:C2')
    expect(pivotFieldOptions(f.ctx, 'B1:C2')).toEqual([
      { label: 'Name', colIndex: 1 },
      { label: 'Amount', colIndex: 2 },
    ])
  })

  it('keeps empty and isolated cells unchanged', () => {
    const f = fixture([[null], [], [false]])
    f.select('A1')
    expect(getSourceRange(f.ctx)).toBe('A1:A1')
    f.select('A3')
    expect(getSourceRange(f.ctx)).toBe('A3:A3')
  })

  it('does not infer a boundary from an incompletely loaded workbook', () => {
    const f = fixture()
    f.state.flags.preloadComplete = false
    expect(getSourceRange(f.ctx)).toBe('B2:B2')
    expect(f.reads).toEqual([])
    f.select('A1:C3')
    expect(getSourceRange(f.ctx)).toBe('A1:C3')
  })

  it('prefers table boundaries over adjacent data and includes internal empty rows', () => {
    const f = fixture([['Name', 'Amount', 'Adjacent'], ['A', 2, 100], [], ['B', 3, 200]])
    addTable(f, 'A1:B4')
    f.select('B3')
    expect(getSourceRange(f.ctx)).toBe('A1:B4')
  })

  it('excludes table totals even when the totals cell is selected', () => {
    const f = fixture([
      ['Name', 'Amount'],
      ['A', 2],
      ['Total', 2],
    ])
    addTable(f, 'A1:B3', 1)
    f.select('B3')
    expect(getSourceRange(f.ctx)).toBe('A1:B2')
  })

  it('keeps a data row inserted at the totals row inside the source', () => {
    const f = fixture([
      ['Name', 'Amount'],
      ['A', 2],
      ['New', 5],
      ['Total', 7],
    ])
    addTable(f, 'A1:B3', 1)
    f.state.editJournal.structuralOps.set('s1', [{ kind: 'insert-rows', index: 2, count: 1 }])
    f.select('B2')
    expect(getSourceRange(f.ctx)).toBe('A1:B3')
  })

  it('maps saved table bounds after a row insertion', () => {
    const f = fixture([[], ['Region', 'Revenue', 'Units'], ['East', 10, 2], ['West', 20, 3]])
    addTable(f, 'A1:C3')
    f.state.editJournal.structuralOps.set('s1', [{ kind: 'insert-rows', index: 0, count: 1 }])
    f.select('B3')
    expect(getSourceRange(f.ctx)).toBe('A2:C4')
  })

  it('uses newly created table bounds from the edit journal', () => {
    const f = fixture()
    f.state.editJournal.tableAdds.push({
      sheetId: 's1',
      area: parseRange('A1:B3'),
      name: 'Sales',
      columnNames: ['Region', 'Revenue'],
      bandedRows: true,
    })
    expect(getSourceRange(f.ctx)).toBe('A1:B3')
  })

  it('includes newly inserted data rows inside a saved table', () => {
    const f = fixture([
      ['Region', 'Revenue'],
      ['New', 5],
      ['East', 10],
      ['West', 20],
    ])
    addTable(f, 'A1:B3')
    f.state.editJournal.structuralOps.set('s1', [{ kind: 'insert-rows', index: 1, count: 1 }])
    f.select('B2')
    expect(getSourceRange(f.ctx)).toBe('A1:B4')
  })

  it('preserves an explicit range containing a table totals row', () => {
    const f = fixture([
      ['Name', 'Amount'],
      ['A', 2],
      ['Total', 2],
    ])
    addTable(f, 'A1:B3', 1)
    f.select('A1:B3')
    expect(getSourceRange(f.ctx)).toBe('A1:B3')
  })

  it('does not infer a header row for a headerless table', () => {
    const f = fixture()
    addTable(f, 'A1:C3')
    f.state.file.sheets[0]!.tables[0]!.headerRowCount = 0
    expect(getSourceRange(f.ctx)).toBe('B2:B2')
  })

  it('creates from the resolved source, not the selected data cell', () => {
    const f = fixture()
    const config: OoXmlPivotConfig = {
      sourceRange: getSourceRange(f.ctx),
      targetCell: 'E1',
      rowFieldIndices: [0],
      colFieldIndices: [],
      groupings: [],
      labelFilters: [],
      valueFilters: [],
      values: [{ fieldIndex: 1, agg: 'sum' }],
    }
    expect(handleCreatePivot(f.ctx, config)).toBeNull()
    expect(applyAiPivotAdd).toHaveBeenLastCalledWith(
      f.ctx.univerRef.current,
      f.state,
      expect.objectContaining({
        sourceRange: 'A1:C3',
        targetCell: 'E1',
        rowFields: ['Region'],
        values: [{ field: 'Revenue', agg: 'sum' }],
      }),
    )
  })
})
