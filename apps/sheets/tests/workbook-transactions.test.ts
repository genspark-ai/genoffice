import { describe, expect, it } from 'vitest'

import { InMemoryWorkbookAdapter, WorkbookConflictError } from '../src/domain/in-memory-workbook'
import type { WorkbookSnapshot } from '../src/domain/workbook.types'

function createAdapter(): InMemoryWorkbookAdapter {
  const snapshot: WorkbookSnapshot = {
    revision: 0,
    sheets: [{ id: 'sheet-1', name: 'Sheet1', cells: { A1: { value: 'before' } } }],
  }
  return new InMemoryWorkbookAdapter(snapshot)
}

function createChartAdapter(): InMemoryWorkbookAdapter {
  const snapshot: WorkbookSnapshot = {
    revision: 0,
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet1',
        cells: {
          A1: { value: 'Region' },
          B1: { value: 'Revenue' },
          A2: { value: 'North' },
          B2: { value: 128000 },
          A3: { value: 'South' },
          B3: { value: 96000 },
        },
      },
    ],
  }
  return new InMemoryWorkbookAdapter(snapshot)
}

describe('InMemoryWorkbookAdapter', () => {
  it('keeps dry-run planning side-effect free', () => {
    const adapter = createAdapter()
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-1',
      baseRevision: 0,
      summary: 'Change A1',
      operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'after' }],
    })

    expect(plan.cellChanges[0]?.before.value).toBe('before')
    expect(adapter.getSnapshot().sheets[0]?.cells.A1?.value).toBe('before')
    expect(adapter.getSnapshot().revision).toBe(0)
  })

  it('commits one transaction and restores content through undo', () => {
    const adapter = createAdapter()
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-1',
      baseRevision: 0,
      summary: 'Change A1',
      operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'after' }],
    })

    expect(adapter.apply(plan).revision).toBe(1)
    expect(adapter.getSnapshot().sheets[0]?.cells.A1?.value).toBe('after')
    expect(adapter.undo().revision).toBe(2)
    expect(adapter.getSnapshot().sheets[0]?.cells.A1?.value).toBe('before')
    expect(adapter.getAuditLog()).toEqual([
      {
        transactionId: 'tx-1',
        action: 'apply',
        previousRevision: 0,
        revision: 1,
        changedCellCount: 1,
        renamedSheetCount: 0,
        structuralChangeCount: 0,
      },
      {
        transactionId: 'undo-1',
        action: 'undo',
        previousRevision: 1,
        revision: 2,
        changedCellCount: 0,
        renamedSheetCount: 0,
        structuralChangeCount: 0,
      },
    ])
  })

  it('rejects stale and repeated transactions', () => {
    const adapter = createAdapter()
    const command = {
      dslVersion: 1 as const,
      transactionId: 'tx-1',
      baseRevision: 0,
      summary: 'Change A1',
      operations: [{ op: 'set_cell' as const, sheetId: 'sheet-1', address: 'A1', value: 'after' }],
    }
    const plan = adapter.plan(command)
    adapter.apply(plan)

    expect(() => adapter.apply(plan)).toThrow(WorkbookConflictError)
    expect(() => adapter.plan({ ...command, transactionId: 'tx-2' })).toThrow(WorkbookConflictError)
  })

  it('expands set_range into per-cell changes with formula detection', () => {
    const adapter = createAdapter()
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-range',
      baseRevision: 0,
      summary: 'Fill B2:C3',
      operations: [
        {
          op: 'set_range',
          sheetId: 'sheet-1',
          start: 'B2',
          values: [
            [1, 2],
            ['x', '=SUM(B2:C2)'],
          ],
        },
      ],
    })

    expect(plan.cellChanges.map((c) => c.address)).toEqual(['B2', 'C2', 'B3', 'C3'])
    expect(plan.cellChanges[3]?.after).toEqual({ value: null, formula: '=SUM(B2:C2)' })
    adapter.apply(plan)
    const cells = adapter.getSnapshot().sheets[0]?.cells
    expect(cells?.B2?.value).toBe(1)
    expect(cells?.C3?.formula).toBe('=SUM(B2:C2)')
  })

  it('clears every cell in a clear_range', () => {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [{ id: 'sheet-1', name: 'Sheet1', cells: { A1: { value: 1 }, B2: { value: 2 } } }],
    })
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-clear',
      baseRevision: 0,
      summary: 'Clear A1:B2',
      operations: [{ op: 'clear_range', sheetId: 'sheet-1', range: 'A1:B2' }],
    })
    adapter.apply(plan)
    const cells = adapter.getSnapshot().sheets[0]?.cells
    expect(cells?.A1).toEqual({ value: null })
    expect(cells?.B2).toEqual({ value: null })
  })

  it('shifts cells down on insert_rows and restores them through undo', () => {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: { A1: { value: 'header' }, A2: { value: 'data' } },
        },
      ],
    })
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-insert',
      baseRevision: 0,
      summary: 'Insert a row before row 2',
      operations: [{ op: 'insert_rows', sheetId: 'sheet-1', row: 2, count: 1 }],
    })

    expect(plan.structuralChanges[0]?.label).toBe('Insert 1 row before row 2')
    expect(plan.warnings).toEqual([])
    adapter.apply(plan)
    const cells = adapter.getSnapshot().sheets[0]?.cells
    expect(cells?.A1?.value).toBe('header')
    expect(cells?.A2).toBeUndefined()
    expect(cells?.A3?.value).toBe('data')

    adapter.undo()
    expect(adapter.getSnapshot().sheets[0]?.cells.A2?.value).toBe('data')
  })

  it('drops deleted rows and shifts the rest up', () => {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: { A1: { value: 'r1' }, A2: { value: 'r2' }, A3: { value: 'r3' } },
        },
      ],
    })
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-delete',
      baseRevision: 0,
      summary: 'Delete row 2',
      operations: [{ op: 'delete_rows', sheetId: 'sheet-1', row: 2, count: 1 }],
    })
    adapter.apply(plan)
    const cells = adapter.getSnapshot().sheets[0]?.cells
    expect(cells?.A1?.value).toBe('r1')
    expect(cells?.A2?.value).toBe('r3')
    expect(cells?.A3).toBeUndefined()
  })

  it('shifts columns on insert_cols and delete_cols', () => {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [
        { id: 'sheet-1', name: 'Sheet1', cells: { A1: { value: 'a' }, B1: { value: 'b' } } },
      ],
    })
    const inserted = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-cols-1',
      baseRevision: 0,
      summary: 'Insert a column before B',
      operations: [{ op: 'insert_cols', sheetId: 'sheet-1', column: 'B', count: 1 }],
    })
    adapter.apply(inserted)
    expect(adapter.getSnapshot().sheets[0]?.cells.C1?.value).toBe('b')

    const deleted = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-cols-2',
      baseRevision: 1,
      summary: 'Delete column A',
      operations: [{ op: 'delete_cols', sheetId: 'sheet-1', column: 'A', count: 1 }],
    })
    adapter.apply(deleted)
    const cells = adapter.getSnapshot().sheets[0]?.cells
    expect(cells?.B1?.value).toBe('b')
    expect(cells?.A1).toBeUndefined()
  })

  it('adds a sheet with a unique generated id and rejects duplicate names', () => {
    const adapter = createAdapter()
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-sheet',
      baseRevision: 0,
      summary: 'Add Summary sheet',
      operations: [{ op: 'add_sheet', name: 'Summary' }],
    })
    adapter.apply(plan)
    const sheets = adapter.getSnapshot().sheets
    expect(sheets).toHaveLength(2)
    expect(sheets[1]).toEqual({ id: 'sheet-2', name: 'Summary', cells: {} })

    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-sheet-2',
        baseRevision: 1,
        summary: 'Duplicate name',
        operations: [{ op: 'add_sheet', name: 'Sheet1' }],
      }),
    ).toThrow(WorkbookConflictError)
  })

  it('rewrites formula references when rows shift and flags #REF! deletions', () => {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: {
            B2: { value: 10 },
            B3: { value: 20 },
            B4: { value: null, formula: '=SUM(B2:B3)' },
            C4: { value: null, formula: '=B2*2' },
          },
        },
      ],
    })
    const inserted = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-shift-1',
      baseRevision: 0,
      summary: 'Insert row before row 2',
      operations: [{ op: 'insert_rows', sheetId: 'sheet-1', row: 2, count: 1 }],
    })
    expect(inserted.warnings).toEqual([])
    adapter.apply(inserted)
    let cells = adapter.getSnapshot().sheets[0]?.cells
    expect(cells?.B5?.formula).toBe('=SUM(B3:B4)')
    expect(cells?.C5?.formula).toBe('=B3*2')

    const deleted = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-shift-2',
      baseRevision: 1,
      summary: 'Delete row 3 (first data row)',
      operations: [{ op: 'delete_rows', sheetId: 'sheet-1', row: 3, count: 1 }],
    })
    expect(deleted.warnings[0]).toContain('#REF!')
    expect(deleted.warnings[0]).toContain('Sheet1!C4')
    adapter.apply(deleted)
    cells = adapter.getSnapshot().sheets[0]?.cells
    expect(cells?.B4?.formula).toBe('=SUM(B3:B3)')
    expect(cells?.C4?.formula).toBe('=#REF!*2')
  })

  it('merges format_range patches into per-cell styles and undoes them', () => {
    const adapter = createAdapter()
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-fmt-1',
      baseRevision: 0,
      summary: 'Bold the header',
      operations: [
        {
          op: 'format_range',
          sheetId: 'sheet-1',
          range: 'A1:B1',
          format: { bold: true, fillColor: '#FFF2CC' },
        },
      ],
    })
    expect(plan.formatChanges[0]?.label).toBe('A1:B1: bold, fill #FFF2CC')
    expect(plan.cellChanges).toEqual([])
    adapter.apply(plan)
    const styled = adapter.getSnapshot().sheets[0]?.styles
    expect(styled?.A1).toEqual({ bold: true, fillColor: '#FFF2CC' })
    expect(styled?.B1).toEqual({ bold: true, fillColor: '#FFF2CC' })

    const cleared = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-fmt-2',
      baseRevision: 1,
      summary: 'Remove the fill',
      operations: [
        {
          op: 'format_range',
          sheetId: 'sheet-1',
          range: 'A1:B1',
          format: { fillColor: null },
        },
      ],
    })
    adapter.apply(cleared)
    expect(adapter.getSnapshot().sheets[0]?.styles?.A1).toEqual({ bold: true })

    adapter.undo()
    expect(adapter.getSnapshot().sheets[0]?.styles?.A1).toEqual({
      bold: true,
      fillColor: '#FFF2CC',
    })
  })

  it('re-keys styles when rows shift and rejects an empty format patch', () => {
    const adapter = createAdapter()
    adapter.apply(
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-fmt-3',
        baseRevision: 0,
        summary: 'Style A1',
        operations: [
          { op: 'format_range', sheetId: 'sheet-1', range: 'A1', format: { bold: true } },
        ],
      }),
    )
    adapter.apply(
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-fmt-4',
        baseRevision: 1,
        summary: 'Insert a row on top',
        operations: [{ op: 'insert_rows', sheetId: 'sheet-1', row: 1, count: 1 }],
      }),
    )
    const sheet = adapter.getSnapshot().sheets[0]
    expect(sheet?.styles?.A2).toEqual({ bold: true })
    expect(sheet?.styles?.A1).toBeUndefined()

    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-fmt-5',
        baseRevision: 2,
        summary: 'Empty patch',
        operations: [{ op: 'format_range', sheetId: 'sheet-1', range: 'A1', format: {} }],
      }),
    ).toThrow()
  })

  it('sorts a range into per-cell changes with CAS protection', () => {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet1',
          cells: {
            A1: { value: 'Name' },
            B1: { value: 'Score' },
            A2: { value: 'Cara' },
            B2: { value: 70 },
            A3: { value: 'Ann' },
            B3: { value: 90 },
          },
        },
      ],
    })
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-sort',
      baseRevision: 0,
      summary: 'Sort by score desc',
      operations: [
        {
          op: 'sort_range',
          sheetId: 'sheet-1',
          range: 'A1:B3',
          byColumn: 'B',
          order: 'desc',
          hasHeader: true,
        },
      ],
    })
    expect(plan.cellChanges.map((c) => [c.address, c.after.value])).toEqual([
      ['A2', 'Ann'],
      ['B2', 90],
      ['A3', 'Cara'],
      ['B3', 70],
    ])
    adapter.apply(plan)
    expect(adapter.getSnapshot().sheets[0]?.cells.A2?.value).toBe('Ann')
  })

  it('tracks merges, row heights, and column widths and re-keys them on shifts', () => {
    const adapter = createAdapter()
    adapter.apply(
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-layout',
        baseRevision: 0,
        summary: 'Merge title, size header',
        operations: [
          { op: 'merge_cells', sheetId: 'sheet-1', range: 'A1:C1' },
          { op: 'set_row_height', sheetId: 'sheet-1', row: 1, heightPoints: 24 },
          { op: 'set_col_width', sheetId: 'sheet-1', column: 'B', count: 2, widthPx: 120 },
        ],
      }),
    )
    let sheet = adapter.getSnapshot().sheets[0]
    expect(sheet?.merges).toEqual(['A1:C1'])
    expect(sheet?.rowHeights).toEqual({ '1': 24 })
    expect(sheet?.colWidths).toEqual({ B: 120, C: 120 })

    adapter.apply(
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-layout-shift',
        baseRevision: 1,
        summary: 'Insert a row on top',
        operations: [{ op: 'insert_rows', sheetId: 'sheet-1', row: 1, count: 1 }],
      }),
    )
    sheet = adapter.getSnapshot().sheets[0]
    expect(sheet?.merges).toEqual(['A2:C2'])
    expect(sheet?.rowHeights).toEqual({ '2': 24 })
    expect(sheet?.colWidths).toEqual({ B: 120, C: 120 })

    adapter.undo()
    adapter.undo()
    sheet = adapter.getSnapshot().sheets[0]
    expect(sheet?.merges).toBeUndefined()
    expect(sheet?.rowHeights).toBeUndefined()
  })

  it('rejects overlapping merges and unmerges intersecting ranges', () => {
    const adapter = createAdapter()
    adapter.apply(
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-merge-1',
        baseRevision: 0,
        summary: 'Merge A1:B2',
        operations: [{ op: 'merge_cells', sheetId: 'sheet-1', range: 'A1:B2' }],
      }),
    )
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-merge-2',
        baseRevision: 1,
        summary: 'Overlapping merge',
        operations: [{ op: 'merge_cells', sheetId: 'sheet-1', range: 'B2:C3' }],
      }),
    ).toThrow(WorkbookConflictError)

    adapter.apply(
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-merge-3',
        baseRevision: 1,
        summary: 'Unmerge',
        operations: [{ op: 'unmerge_cells', sheetId: 'sheet-1', range: 'A1:A1' }],
      }),
    )
    expect(adapter.getSnapshot().sheets[0]?.merges).toBeUndefined()
  })

  it('deletes a sheet, warning about formulas that reference it', () => {
    const adapter = new InMemoryWorkbookAdapter({
      revision: 0,
      sheets: [
        { id: 'sheet-1', name: 'Data', cells: { A1: { value: 1 } } },
        { id: 'sheet-2', name: 'Summary', cells: { B1: { value: null, formula: '=Data!A1*2' } } },
      ],
    })
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-del-sheet',
      baseRevision: 0,
      summary: 'Delete the Data sheet',
      operations: [{ op: 'delete_sheet', sheetId: 'sheet-1' }],
    })
    expect(plan.structuralChanges[0]?.label).toBe('Delete sheet sheet-1')
    expect(plan.warnings[0]).toContain('Summary!B1')
    adapter.apply(plan)
    expect(adapter.getSnapshot().sheets.map((sheet) => sheet.id)).toEqual(['sheet-2'])

    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-del-last',
        baseRevision: 1,
        summary: 'Delete the last sheet',
        operations: [{ op: 'delete_sheet', sheetId: 'sheet-2' }],
      }),
    ).toThrow(WorkbookConflictError)
  })

  it('edits a demo chart by visual id and rejects unknown or empty chart edits', () => {
    const adapter = createChartAdapter()
    const addPlan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-chart-add',
      baseRevision: 0,
      summary: 'Add a chart',
      operations: [
        { op: 'add_chart', sheetId: 'sheet-1', chartType: 'column', dataRange: 'A1:B3' },
      ],
    })
    adapter.apply(addPlan)
    const visualId = adapter.getSnapshot().sheets[0]?.visuals?.[0]?.id ?? ''
    const editPlan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-chart-edit',
      baseRevision: 1,
      summary: 'Edit the chart',
      operations: [
        {
          op: 'edit_chart',
          chartPath: visualId,
          title: 'Sales',
          chartType: 'pie',
          legend: 'bottom',
          dataLabels: 'category-percent',
        },
      ],
    })
    adapter.apply(editPlan)
    const chart = adapter.getSnapshot().sheets[0]?.visuals?.[0]?.chart
    expect(chart?.title).toBe('Sales')
    expect(chart?.chartTypes).toEqual(['pieChart'])
    expect(chart?.legend).toBe('bottom')
    expect(chart?.dataLabels).toBe('category-percent')

    expect(() =>
      adapter.apply(
        adapter.plan({
          dslVersion: 1,
          transactionId: 'tx-chart-unknown',
          baseRevision: 2,
          summary: 'Edit a missing chart',
          operations: [{ op: 'edit_chart', chartPath: 'demo-chart-sheet-1-99', title: 'Sales' }],
        }),
      ),
    ).toThrow(/Unknown chart/)
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-chart-empty',
        baseRevision: 2,
        summary: 'Empty edit',
        operations: [{ op: 'edit_chart', chartPath: 'xl/charts/chart1.xml' }],
      }),
    ).toThrow(/at least one/)
  })

  it('deletes a demo visual by id and rejects unknown ids', () => {
    const adapter = createChartAdapter()
    adapter.apply(
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-add-for-delete',
        baseRevision: 0,
        summary: 'Add a chart',
        operations: [{ op: 'add_chart', sheetId: 'sheet-1', chartType: 'pie', dataRange: 'A1:B3' }],
      }),
    )
    // the value-label default is bar/column only — pie keeps its legend
    expect(adapter.getSnapshot().sheets[0]?.visuals?.[0]?.chart.dataLabels).toBeUndefined()
    const visualId = adapter.getSnapshot().sheets[0]?.visuals?.[0]?.id ?? ''
    adapter.apply(
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-delete-visual',
        baseRevision: 1,
        summary: 'Delete the chart',
        operations: [{ op: 'delete_visual', visualId }],
      }),
    )
    expect(adapter.getSnapshot().sheets[0]?.visuals ?? []).toHaveLength(0)
    expect(() =>
      adapter.apply(
        adapter.plan({
          dslVersion: 1,
          transactionId: 'tx-delete-missing',
          baseRevision: 2,
          summary: 'Delete a missing visual',
          operations: [{ op: 'delete_visual', visualId: 'demo-chart-nope' }],
        }),
      ),
    ).toThrow(/Unknown visual/)
  })

  it('inserts a chart into the demo workbook and removes it through undo', () => {
    const adapter = createChartAdapter()
    const plan = adapter.plan({
      dslVersion: 1,
      transactionId: 'tx-add-chart',
      baseRevision: 0,
      summary: 'Add a chart',
      operations: [
        {
          op: 'add_chart',
          sheetId: 'sheet-1',
          chartType: 'column',
          dataRange: 'A1:B3',
          title: 'Regional Revenue',
        },
      ],
    })
    expect(plan.structuralChanges[0]?.label).toBe(
      'Insert column chart from A1:B3 "Regional Revenue"',
    )
    // plan is a dry run — nothing lands until apply
    expect(adapter.getSnapshot().sheets[0]?.visuals).toBeUndefined()

    adapter.apply(plan)
    const visual = adapter.getSnapshot().sheets[0]?.visuals?.[0]
    expect(visual?.kind).toBe('chart')
    expect(visual?.chart.title).toBe('Regional Revenue')
    expect(visual?.chart.chartTypes).toEqual(['barChart'])
    expect(visual?.chart.barDirection).toBe('col')
    // new bar/column charts label each bar by default
    expect(visual?.chart.dataLabels).toBe('value')
    expect(visual?.chart.series[0]).toMatchObject({
      name: 'Revenue',
      categories: ['North', 'South'],
      values: [128000, 96000],
      valuesRef: "'Sheet1'!$B$2:$B$3",
      categoriesRef: "'Sheet1'!$A$2:$A$3",
    })
    // default anchor: two columns right of the data
    expect(visual?.anchor.fromRow).toBe(0)
    expect(visual?.anchor.fromColumn).toBe(3)

    adapter.undo()
    expect(adapter.getSnapshot().sheets[0]?.visuals).toBeUndefined()
  })

  it('rejects add_chart without a numeric column or over 2000 cells', () => {
    const adapter = createChartAdapter()
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-chart-text',
        baseRevision: 0,
        summary: 'Chart over text',
        operations: [{ op: 'add_chart', sheetId: 'sheet-1', chartType: 'pie', dataRange: 'A1:A3' }],
      }),
    ).toThrow(/numeric column/)
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-chart-huge',
        baseRevision: 0,
        summary: 'Chart over everything',
        operations: [
          { op: 'add_chart', sheetId: 'sheet-1', chartType: 'line', dataRange: 'A1:Z100' },
        ],
      }),
    ).toThrow(/2000 cells/)
  })

  it('rejects a batch mixing structural and cell-content operations', () => {
    const adapter = createAdapter()
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-mixed',
        baseRevision: 0,
        summary: 'Insert then write',
        operations: [
          { op: 'insert_rows', sheetId: 'sheet-1', row: 1, count: 1 },
          { op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'x' },
        ],
      }),
    ).toThrow()
  })

  it('rejects invalid DSL before it reaches workbook state', () => {
    const adapter = createAdapter()
    expect(() =>
      adapter.plan({
        dslVersion: 1,
        transactionId: 'tx-1',
        baseRevision: 0,
        summary: 'Out of bounds',
        operations: [{ op: 'set_cell', sheetId: 'sheet-1', address: '../../A1', value: 1 }],
      }),
    ).toThrow()
    expect(adapter.getSnapshot()).toEqual(createAdapter().getSnapshot())
  })
})
