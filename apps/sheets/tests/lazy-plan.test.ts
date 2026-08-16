import { describe, expect, it } from 'vitest'

import { planPrompt } from '../src/ai/deterministic-planner'
import { workbookCommandBatchSchema } from '../src/domain/workbook-dsl'
import type { CellState } from '../src/domain/workbook.types'
import { buildLazyChangePlan, planStillMatches } from '../src/renderer/lazy-plan'

const CELLS: Record<string, CellState> = {
  B2: { value: 'existing' },
  C3: { value: null, formula: '=SUM(A1:A2)' },
}
const readCell = (address: string): CellState => CELLS[address] ?? { value: null }

function plan(prompt: string) {
  const batch = workbookCommandBatchSchema.parse(
    planPrompt(prompt, { revision: 0, sheetId: 'sheet-1' }),
  )
  return buildLazyChangePlan(batch, readCell, () => 'Data')
}

describe('buildLazyChangePlan', () => {
  it('previews a value set with the live cell as before-state', () => {
    const result = plan('set B2 to 4242')
    expect(result.cellChanges).toEqual([
      {
        sheetId: 'sheet-1',
        address: 'B2',
        before: { value: 'existing' },
        after: { value: 4242 },
      },
    ])
    expect(result.sheetRenames).toEqual([])
  })

  it('previews a formula set over a formula cell', () => {
    const result = plan('formula C3 = MAX(A1:A9)')
    expect(result.cellChanges).toEqual([
      {
        sheetId: 'sheet-1',
        address: 'C3',
        before: { value: null, formula: '=SUM(A1:A2)' },
        after: { value: null, formula: '=MAX(A1:A9)' },
      },
    ])
  })

  it('previews a sheet rename with the current name as before-state', () => {
    const result = plan('rename sheet to Budget 2027')
    expect(result.sheetRenames).toEqual([
      {
        sheetId: 'sheet-1',
        before: 'Data',
        after: 'Budget 2027',
      },
    ])
    expect(result.cellChanges).toEqual([])
  })

  it('rejects an invalid sheet name through the DSL schema', () => {
    expect(() => plan('rename sheet to bad[name]')).toThrow()
  })
})

describe('buildLazyChangePlan: DSL v2 operations', () => {
  function planBatch(operations: unknown[]) {
    const batch = workbookCommandBatchSchema.parse({
      dslVersion: 1,
      transactionId: 'tx-v2',
      baseRevision: 0,
      summary: 'v2 ops',
      operations,
    })
    return buildLazyChangePlan(batch, readCell, () => 'Data')
  }

  it('expands set_range against live before-states', () => {
    const result = planBatch([
      { op: 'set_range', sheetId: 'sheet-1', start: 'B2', values: [['new', '=A1+1']] },
    ])
    expect(result.cellChanges).toEqual([
      { sheetId: 'sheet-1', address: 'B2', before: { value: 'existing' }, after: { value: 'new' } },
      {
        sheetId: 'sheet-1',
        address: 'C2',
        before: { value: null },
        after: { value: null, formula: '=A1+1' },
      },
    ])
  })

  it('previews structural operations as labeled entries without before-states', () => {
    const result = planBatch([
      { op: 'insert_rows', sheetId: 'sheet-1', row: 2, count: 2 },
      { op: 'add_sheet', name: 'Summary' },
    ])
    expect(result.cellChanges).toEqual([])
    expect(result.structuralChanges.map((change) => change.label)).toEqual([
      'Insert 2 rows before row 2',
      'Add sheet "Summary"',
    ])
  })

  it('a structural-only plan always still matches (no cell CAS to verify)', () => {
    const result = planBatch([{ op: 'delete_rows', sheetId: 'sheet-1', row: 5, count: 1 }])
    expect(planStillMatches(result, () => ({ value: 'anything' }))).toBe(true)
  })

  it('previews format_range as a labeled format change', () => {
    const result = planBatch([
      {
        op: 'format_range',
        sheetId: 'sheet-1',
        range: 'A1:C1',
        format: { bold: true, numberFormat: '0.00%' },
      },
    ])
    expect(result.cellChanges).toEqual([])
    expect(result.formatChanges).toEqual([
      {
        sheetId: 'sheet-1',
        range: 'A1:C1',
        format: { bold: true, numberFormat: '0.00%' },
        label: 'A1:C1: bold, number format 0.00%',
      },
    ])
  })

  it('expands sort_range against live cells into CAS-protected changes', () => {
    const cells: Record<string, CellState> = {
      A1: { value: 30 },
      A2: { value: 10 },
      A3: { value: 20 },
    }
    const result = buildLazyChangePlan(
      workbookCommandBatchSchema.parse({
        dslVersion: 1,
        transactionId: 'tx-sort',
        baseRevision: 0,
        summary: 'sort asc',
        operations: [
          { op: 'sort_range', sheetId: 'sheet-1', range: 'A1:A3', byColumn: 'A', order: 'asc' },
        ],
      }),
      (address) => cells[address] ?? { value: null },
      () => 'Data',
    )
    expect(result.cellChanges.map((c) => [c.address, c.before.value, c.after.value])).toEqual([
      ['A1', 30, 10],
      ['A2', 10, 20],
      ['A3', 20, 30],
    ])
  })

  it('previews layout operations as labeled entries', () => {
    const result = planBatch([
      { op: 'merge_cells', sheetId: 'sheet-1', range: 'A1:C1' },
      { op: 'set_col_width', sheetId: 'sheet-1', column: 'B', widthPx: 140 },
    ])
    expect(result.structuralChanges.map((change) => change.label)).toEqual([
      'Merge A1:C1',
      'Set column B width to 140px',
    ])
  })

  it('previews chart edits with a descriptive label, mixable with content ops', () => {
    const result = planBatch([
      {
        op: 'edit_chart',
        chartPath: 'xl/charts/chart2.xml',
        title: 'Q3',
        chartType: 'line',
        seriesColors: { '0': '#4472C4' },
      },
      { op: 'set_cell', sheetId: 'sheet-1', address: 'A1', value: 'x' },
    ])
    expect(result.structuralChanges[0]?.label).toBe(
      'Edit chart xl/charts/chart2.xml: title "Q3", type line, series colors #0→#4472C4',
    )
    expect(result.cellChanges).toHaveLength(1)
  })

  it('rejects format_range mixed with structural ops in one batch', () => {
    expect(() =>
      planBatch([
        { op: 'insert_rows', sheetId: 'sheet-1', row: 1, count: 1 },
        { op: 'format_range', sheetId: 'sheet-1', range: 'A1', format: { bold: true } },
      ]),
    ).toThrow(/separate batches/)
  })
})

describe('buildLazyChangePlan: cross-sheet routing', () => {
  it("reads before-states and sheet names from each operation's own sheet", () => {
    const grids: Record<string, Record<string, CellState>> = {
      'sheet-1': { A1: { value: 'active' } },
      'sheet-2': { A1: { value: 'other' } },
    }
    const result = buildLazyChangePlan(
      workbookCommandBatchSchema.parse({
        dslVersion: 1,
        transactionId: 'tx-cross',
        baseRevision: 0,
        summary: 'cross-sheet batch',
        operations: [
          { op: 'set_cell', sheetId: 'sheet-2', address: 'A1', value: 'new' },
          { op: 'rename_sheet', sheetId: 'sheet-2', name: 'Renamed' },
        ],
      }),
      (address, sheetId) => grids[sheetId]?.[address] ?? { value: null },
      (sheetId) => (sheetId === 'sheet-2' ? 'Other' : 'Data'),
    )
    expect(result.cellChanges).toEqual([
      { sheetId: 'sheet-2', address: 'A1', before: { value: 'other' }, after: { value: 'new' } },
    ])
    expect(result.sheetRenames).toEqual([{ sheetId: 'sheet-2', before: 'Other', after: 'Renamed' }])
  })
})

describe('planStillMatches', () => {
  it('accepts when previewed cells are unchanged and rejects drift', () => {
    const result = plan('set B2 to 4242')
    expect(planStillMatches(result, readCell)).toBe(true)
    expect(planStillMatches(result, () => ({ value: 'changed meanwhile' }))).toBe(false)
  })

  it('treats formula changes as drift', () => {
    const result = plan('formula C3 = MAX(A1:A9)')
    expect(planStillMatches(result, readCell)).toBe(true)
    expect(planStillMatches(result, () => ({ value: null, formula: '=SUM(A1:A3)' }))).toBe(false)
  })

  it("checks drift against each change's own sheet", () => {
    const result = buildLazyChangePlan(
      workbookCommandBatchSchema.parse({
        dslVersion: 1,
        transactionId: 'tx-cross-drift',
        baseRevision: 0,
        summary: 'cross-sheet drift',
        operations: [{ op: 'set_cell', sheetId: 'sheet-2', address: 'A1', value: 'new' }],
      }),
      (_address, sheetId) => (sheetId === 'sheet-2' ? { value: 'other' } : { value: 'active' }),
      () => 'Other',
    )
    expect(
      planStillMatches(result, (_address, sheetId) =>
        sheetId === 'sheet-2' ? { value: 'other' } : { value: 'changed' },
      ),
    ).toBe(true)
    expect(
      planStillMatches(result, (_address, sheetId) =>
        sheetId === 'sheet-2' ? { value: 'changed' } : { value: 'active' },
      ),
    ).toBe(false)
  })
})
