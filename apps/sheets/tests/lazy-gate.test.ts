import { describe, expect, it, vi } from 'vitest'

import { lazyGateError, proposeOperations, type PlanContext } from '../src/renderer/plan-operations'
import type { LazyWorkbookState } from '../src/renderer/univer-state'
import type { WorkbookOperation } from '../src/domain/workbook-dsl'

/// The BeforeCommandExecute gates in App.tsx cancel these facade commands
/// silently; lazyGateError mirrors them so propose/apply fail loud instead.

function lazyState(overrides: Record<string, unknown> = {}): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      sheets: [
        { id: 'sh1', name: 'Data', rowCount: 10, columnCount: 5, pivotRanges: [] },
        {
          id: 'sh2',
          name: 'Pivot',
          rowCount: 10,
          columnCount: 5,
          pivotRanges: [{ startRow: 0, endRow: 4, startColumn: 0, endColumn: 2 }],
        },
      ],
      visuals: [],
    },
    editJournal: {
      cells: new Map(),
      structuralOps: new Map(),
      sheets: { added: new Set(), removed: new Set() },
      visualAdds: [],
      tableAdds: [],
    },
    loadedRanges: new Map(),
    formulaMode: true,
    flags: { preloadComplete: true },
    filterOrigins: new Map(),
    appliedDvSheets: new Set(['sh1', 'sh2']),
    ...overrides,
  } as unknown as LazyWorkbookState
}

describe('lazyGateError', () => {
  it.each(['insert_rows', 'delete_rows'] as const)(
    'blocks %s on a pivot sheet and names it',
    (op) => {
      const state = lazyState()
      const error = lazyGateError(state, { op, sheetId: 'sh2', row: 6, count: 1 })
      expect(error).toContain('Pivot')
      expect(error).toContain('PivotTable')
      expect(lazyGateError(state, { op, sheetId: 'sh1', row: 6, count: 1 })).toBeNull()
    },
  )

  it('blocks merges on a pivot sheet regardless of the pivot region', () => {
    const error = lazyGateError(lazyState(), {
      op: 'merge_cells',
      sheetId: 'sh2',
      range: 'D8:E9',
    })
    expect(error).toContain('PivotTable')
  })

  it('blocks filter ops until the workbook is fully loaded', () => {
    const loading = lazyState({ flags: { preloadComplete: false } })
    expect(lazyGateError(loading, { op: 'set_filter', sheetId: 'sh1', range: 'A1:B5' })).toContain(
      'still loading',
    )
    const streamed = lazyState({ formulaMode: false, flags: { preloadComplete: false } })
    expect(lazyGateError(streamed, { op: 'clear_filter', sheetId: 'sh1' })).toContain(
      'fully-loaded mode',
    )
    expect(
      lazyGateError(lazyState(), {
        op: 'set_filter_criteria',
        sheetId: 'sh1',
        column: 'A',
        values: null,
      }),
    ).toBeNull()
  })

  it('blocks writes on a sheet whose XML is above the save-patch cap', () => {
    const state = lazyState({
      file: {
        sessionId: 'session-1',
        sheets: [
          {
            id: 'sh1',
            name: 'Huge',
            rowCount: 10,
            columnCount: 5,
            pivotRanges: [],
            sourceXmlBytes: 501 * 1024 * 1024,
          },
        ],
        visuals: [],
      },
    })
    const error = lazyGateError(state, { op: 'set_cell', sheetId: 'sh1', address: 'A1', value: 1 })
    expect(error).toContain('read-only')
    expect(error).toContain('500MB')
    // workbook.xml-only ops stay allowed
    expect(lazyGateError(state, { op: 'rename_sheet', sheetId: 'sh1', name: 'X' })).toBeNull()
    expect(
      lazyGateError(state, { op: 'set_sheet_hidden', sheetId: 'sh1', hidden: true }),
    ).toBeNull()
    // at/below the cap: no gate
    const under = lazyState()
    expect(
      lazyGateError(under, { op: 'set_cell', sheetId: 'sh1', address: 'A1', value: 1 }),
    ).toBeNull()
  })

  it('gates add_pivot on its output sheet, not its source', () => {
    const state = lazyState({
      file: {
        sessionId: 'session-1',
        sheets: [
          { id: 'small', name: 'Small', rowCount: 10, columnCount: 5, pivotRanges: [] },
          {
            id: 'huge',
            name: 'Huge',
            rowCount: 10,
            columnCount: 5,
            pivotRanges: [],
            sourceXmlBytes: 501 * 1024 * 1024,
          },
        ],
        visuals: [],
      },
    })
    const pivot = (sheetId: string, targetSheetId?: string): WorkbookOperation => ({
      op: 'add_pivot',
      sheetId,
      ...(targetSheetId === undefined ? {} : { targetSheetId }),
      sourceRange: 'A1:B5',
      targetCell: 'D1',
      rowFields: 'h',
      values: [{ field: 'h', agg: 'count' }],
    })
    // small source → oversized output: blocked
    expect(lazyGateError(state, pivot('small', 'huge'))).toContain('read-only')
    // oversized source read → small output: allowed
    expect(lazyGateError(state, pivot('huge', 'small'))).toBeNull()
    // default output = the source sheet
    expect(lazyGateError(state, pivot('huge'))).toContain('read-only')
  })

  it('allows filter ops on sheets added this session even while streaming', () => {
    const state = lazyState({
      formulaMode: false,
      flags: { preloadComplete: false },
      editJournal: {
        cells: new Map(),
        structuralOps: new Map(),
        sheets: { added: new Set(['new1']), removed: new Set() },
        visualAdds: [],
        tableAdds: [],
      },
    })
    expect(lazyGateError(state, { op: 'set_filter', sheetId: 'new1', range: 'A1:B5' })).toBeNull()
  })

  it('blocks filter edits on table-owned filters', () => {
    const state = lazyState({
      filterOrigins: new Map([
        [
          'sh1',
          { origin: 'table', range: { startRow: 0, endRow: 5, startColumn: 0, endColumn: 3 } },
        ],
      ]),
    })
    expect(
      lazyGateError(state, {
        op: 'set_filter_criteria',
        sheetId: 'sh1',
        column: 'A',
        values: ['x'],
      }),
    ).toContain('table')
  })

  it('blocks set_data_validation until the sheet is indexed', () => {
    const state = lazyState({ appliedDvSheets: new Set() })
    expect(
      lazyGateError(state, {
        op: 'set_data_validation',
        sheetId: 'sh1',
        range: 'A1:A5',
        validation: null,
      }),
    ).toContain('still being indexed')
    expect(
      lazyGateError(lazyState(), {
        op: 'set_data_validation',
        sheetId: 'sh1',
        range: 'A1:A5',
        validation: null,
      }),
    ).toBeNull()
  })
})

describe('proposeOperations: gate rejection (lazy workbook)', () => {
  function lazyContext(state: LazyWorkbookState): PlanContext {
    const worksheets = new Map(
      state.file.sheets.map((sheet) => [
        sheet.id,
        {
          getSheetId: () => sheet.id,
          getSheetName: () => sheet.name,
          getMaxRows: () => sheet.rowCount,
          getMaxColumns: () => sheet.columnCount,
          getRange: () => ({
            getValue: () => null,
            getRawValue: () => null,
            getRawValues: () => [[null]],
            getFormula: () => '',
            getCellDatas: () => [[null]],
          }),
        },
      ]),
    )
    const workbook = {
      getActiveSheet: () => worksheets.get('sh1'),
      getSheetBySheetId: (id: string) => worksheets.get(id) ?? null,
    }
    return {
      adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } },
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: { current: state },
      lazyPreviewRef: { current: null },
      setPreview: vi.fn(),
      autoApplySafePlan: vi.fn().mockResolvedValue({ ok: true }),
    } as unknown as PlanContext
  }

  function propose(state: LazyWorkbookState, operation: WorkbookOperation) {
    return proposeOperations(lazyContext(state), [operation], 'test')
  }

  it('rejects structural ops on a pivot sheet at propose time', () => {
    const outcome = propose(lazyState(), { op: 'insert_rows', sheetId: 'sh2', row: 8, count: 1 })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('PivotTable')
  })

  it('rejects merge_cells on a pivot sheet at propose time', () => {
    const outcome = propose(lazyState(), { op: 'merge_cells', sheetId: 'sh2', range: 'D8:E9' })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('PivotTable')
  })

  it('rejects filter ops while the workbook is still loading', () => {
    const state = lazyState({ formulaMode: false, flags: { preloadComplete: false } })
    const outcome = propose(state, { op: 'set_filter', sheetId: 'sh1', range: 'A1:B5' })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('fully-loaded mode')
  })

  it('rejects set_data_validation while the sheet is still indexing', () => {
    const state = lazyState({ appliedDvSheets: new Set() })
    const outcome = propose(state, {
      op: 'set_data_validation',
      sheetId: 'sh1',
      range: 'A1:A5',
      validation: { kind: 'checkbox' },
    })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('still being indexed')
  })

  it('still proposes the same ops on non-gated sheets', () => {
    const outcome = propose(lazyState(), { op: 'insert_rows', sheetId: 'sh1', row: 2, count: 1 })
    expect(outcome.ok).toBe(true)
  })

  it('rejects a copy_range source beyond the source sheet grid at propose time', () => {
    // Source sh1 is 10×5, target sh3 is 100×26: A1:A20 fits the target but
    // reads source rows that do not exist. This used to fail only at apply
    // time (the target-bounds check cannot catch it).
    const state = lazyState({
      file: {
        sessionId: 'session-1',
        sheets: [
          { id: 'sh1', name: 'Data', rowCount: 10, columnCount: 5, pivotRanges: [] },
          { id: 'sh3', name: 'Big', rowCount: 100, columnCount: 26, pivotRanges: [] },
        ],
        visuals: [],
      },
    })
    const outcome = propose(state, {
      op: 'copy_range',
      sheetId: 'sh3',
      sourceSheetId: 'sh1',
      source: 'A1:A20',
      target: 'C1',
    })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('source range')
  })

  it('rejects duplicate_sheet at propose when the source has sheet-scoped defined names', () => {
    // The save clones the part and fails closed on these — without the
    // propose gate the tool reported success and only ⌘S failed.
    const state = lazyState({
      file: {
        sessionId: 'session-1',
        sheets: [
          { id: 'sh1', name: 'Data', rowCount: 10, columnCount: 5, pivotRanges: [] },
          { id: 'sh2', name: 'Other', rowCount: 10, columnCount: 5, pivotRanges: [] },
        ],
        visuals: [],
        definedNames: [{ name: 'LocalName', formula: 'Data!$A$1', sheetIndex: 0 }],
      },
    })
    const outcome = propose(state, { op: 'duplicate_sheet', sheetId: 'sh1' })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('sheet-scoped defined names')
    // scoped to a different sheet: duplicating this one stays allowed
    expect(propose(state, { op: 'duplicate_sheet', sheetId: 'sh2' }).ok).toBe(true)
  })

  it('rejects duplicate_sheet via the sidecar flag when only hidden/_xlnm names are scoped', () => {
    // Hidden and _xlnm.* built-ins never reach file.definedNames; the sidecar
    // reports them per sheet so the gate matches the save-side check exactly.
    const state = lazyState({
      file: {
        sessionId: 'session-1',
        sheets: [
          {
            id: 'sh1',
            name: 'Data',
            rowCount: 10,
            columnCount: 5,
            pivotRanges: [],
            hasScopedDefinedNames: true,
          },
        ],
        visuals: [],
        definedNames: [],
      },
    })
    const outcome = propose(state, { op: 'duplicate_sheet', sheetId: 'sh1' })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('sheet-scoped defined names')
  })

  it('accepts a copy_range whose source fits the source sheet grid', () => {
    const outcome = propose(lazyState(), {
      op: 'copy_range',
      sheetId: 'sh1',
      source: 'A1:A5',
      target: 'C1',
    })
    expect(outcome.ok).toBe(true)
  })

  it('allows add_chart over data an earlier fill in the same batch writes', () => {
    // fill/copy run in the same apply lane before the chart; the pre-apply
    // grid cannot show their data, so the numeric probe is skipped.
    const outcome = proposeOperations(
      lazyContext(lazyState()),
      [
        { op: 'set_cell', sheetId: 'sh1', address: 'A1', value: 7 },
        { op: 'fill_range', sheetId: 'sh1', source: 'A1:A1', target: 'A1:B3' },
        { op: 'add_chart', sheetId: 'sh1', chartType: 'column', dataRange: 'A1:B3' },
      ],
      'test',
    )
    expect(outcome.ok).toBe(true)
  })

  it('rejects add_chart over set_range data from the same batch', () => {
    // cell changes apply AFTER the visual lane — the chart would read an
    // empty grid, so this stays rejected with split-batch guidance.
    const outcome = proposeOperations(
      lazyContext(lazyState()),
      [
        {
          op: 'set_range',
          sheetId: 'sh1',
          range: 'A1:B3',
          values: [
            ['a', 1],
            ['b', 2],
            ['c', 3],
          ],
        },
        { op: 'add_chart', sheetId: 'sh1', chartType: 'column', dataRange: 'A1:B3' },
      ],
      'test',
    )
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('next')
  })

  it('rejects add_chart when nothing in the batch writes its empty range', () => {
    const outcome = propose(lazyState(), {
      op: 'add_chart',
      sheetId: 'sh1',
      chartType: 'column',
      dataRange: 'A1:B3',
    })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('numeric column')
  })
})

describe('proposeOperations: structural anchor bounds (10 rows × 5 columns grid)', () => {
  // Out-of-bounds anchors used to flow through to Univer, whose permission
  // interceptor rejects them AFTER the plan is accepted: the command silently
  // no-ops and the grid pops a misleading "range is protected" dialog.
  function lazyContext(state: LazyWorkbookState): PlanContext {
    const worksheets = new Map(
      state.file.sheets.map((sheet) => [
        sheet.id,
        {
          getSheetId: () => sheet.id,
          getSheetName: () => sheet.name,
          getMaxRows: () => sheet.rowCount,
          getMaxColumns: () => sheet.columnCount,
          getRange: () => ({
            getValue: () => null,
            getRawValue: () => null,
            getRawValues: () => [[null]],
            getFormula: () => '',
            getCellDatas: () => [[null]],
          }),
        },
      ]),
    )
    const workbook = {
      getActiveSheet: () => worksheets.get('sh1'),
      getSheetBySheetId: (id: string) => worksheets.get(id) ?? null,
    }
    return {
      adapterRef: { current: { getSnapshot: () => ({ revision: 0, sheets: [] }) } },
      univerRef: { current: { univerAPI: { getActiveWorkbook: () => workbook } } },
      lazyWorkbookRef: { current: state },
      lazyPreviewRef: { current: null },
      setPreview: vi.fn(),
      autoApplySafePlan: vi.fn().mockResolvedValue({ ok: true }),
    } as unknown as PlanContext
  }

  function propose(operation: WorkbookOperation) {
    return proposeOperations(lazyContext(lazyState()), [operation], 'test')
  }

  it('rejects insert_cols anchored beyond the last grid column', () => {
    const outcome = propose({ op: 'insert_cols', sheetId: 'sh1', column: 'F', count: 1 })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) {
      expect(outcome.error).toContain('beyond the sheet grid')
      expect(outcome.error).toContain('insert before the last column E')
    }
  })

  it('accepts insert_cols before the last column', () => {
    expect(propose({ op: 'insert_cols', sheetId: 'sh1', column: 'E', count: 1 }).ok).toBe(true)
  })

  it('rejects insert_rows anchored beyond the last grid row', () => {
    const outcome = propose({ op: 'insert_rows', sheetId: 'sh1', row: 11, count: 1 })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('beyond the sheet grid')
    expect(propose({ op: 'insert_rows', sheetId: 'sh1', row: 10, count: 1 }).ok).toBe(true)
  })

  it('rejects delete_cols spans that run past the last column', () => {
    const outcome = propose({ op: 'delete_cols', sheetId: 'sh1', column: 'D', count: 3 })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('extend beyond the sheet grid')
    expect(propose({ op: 'delete_cols', sheetId: 'sh1', column: 'D', count: 2 }).ok).toBe(true)
  })

  it('rejects delete_rows spans that run past the last row', () => {
    const outcome = propose({ op: 'delete_rows', sheetId: 'sh1', row: 9, count: 3 })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('extend beyond the sheet grid')
    expect(propose({ op: 'delete_rows', sheetId: 'sh1', row: 9, count: 2 }).ok).toBe(true)
  })
})
