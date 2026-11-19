import { afterEach, describe, expect, it } from 'vitest'

import {
  structuralDeleteFormulaError,
  structuralDeleteFormulaErrorSync,
} from '../src/renderer/plan-operations'
import type { WorkbookSharedFormulaGroup } from '../src/shared/desktop-api'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

/// The save aborts when a formula references only a deleted row/column span;
/// this apply-time precheck runs the same rewrite so the batch fails loud
/// with the workbook untouched instead of at ⌘S.

function state(overrides: Record<string, unknown> = {}): LazyWorkbookState {
  return {
    file: {
      sessionId: 'session-1',
      sheets: [
        { id: 'sh1', name: 'Data', rowCount: 100, columnCount: 26, pivotRanges: [] },
        { id: 'sh2', name: 'Other', rowCount: 100, columnCount: 26, pivotRanges: [] },
      ],
      visuals: [],
    },
    editJournal: {
      cells: new Map(),
      structuralOps: new Map(),
      sheets: { added: new Set(), removed: new Set() },
    },
    ...overrides,
  } as unknown as LazyWorkbookState
}

const workbook = {
  getSheets: () => [
    { getSheetId: () => 'sh1', getSheetName: () => 'Data' },
    { getSheetId: () => 'sh2', getSheetName: () => 'Other' },
  ],
}

function stubFormulas(
  bySheet: Record<string, string[]>,
  sharedGroups: Record<string, WorkbookSharedFormulaGroup[]> = {},
): void {
  ;(globalThis as { window?: unknown }).window = {
    desktopApi: {
      readWorkbookFormulas: async ({ sheetId }: { sheetId: string }) => ({
        cells: (bySheet[sheetId] ?? []).map((formula, i) => ({
          row: 6 + i,
          column: 0,
          formula,
        })),
        sharedGroups: sharedGroups[sheetId] ?? [],
        truncated: false,
        indexingComplete: true,
      }),
    },
  }
}

const delD = { op: 'delete_cols', sheetId: 'sh1', column: 'D', count: 1 } as const

afterEach(() => {
  delete (globalThis as { window?: unknown }).window
})

describe('structuralDeleteFormulaError', () => {
  it('flags a formula referencing only the deleted column', async () => {
    stubFormulas({ sh1: ['=SUM($D$7:$D$9)'] })
    const error = await structuralDeleteFormulaError(state(), workbook, delD)
    expect(error).toContain('deleted columns')
    expect(error).toContain('$D$7:$D$9')
  })

  it('flags a shared-formula follower whose shifted refs sit in the deleted column', async () => {
    // master B1 =SUM(C7:C9) is fine; the follower in C1 reads SUM(D7:D9)
    stubFormulas(
      { sh1: [] },
      {
        sh1: [
          {
            si: 0,
            row: 0,
            column: 1,
            formula: '=SUM(C7:C9)',
            range: { startRow: 0, endRow: 0, startColumn: 1, endColumn: 2 },
          },
        ],
      },
    )
    const error = await structuralDeleteFormulaError(state(), workbook, delD)
    expect(error).toContain('D7:D9')
  })

  it('passes formulas that merely shift or shrink', async () => {
    stubFormulas({ sh1: ['=SUM(E7:E9)', '=SUM(C7:F9)'] })
    expect(await structuralDeleteFormulaError(state(), workbook, delD)).toBeNull()
  })

  it('flags deleted rows the same way', async () => {
    // the formula sits in A7, outside the deleted row 9
    stubFormulas({ sh1: ['=SUM($A$9:$C$9)'] })
    const error = await structuralDeleteFormulaError(state(), workbook, {
      op: 'delete_rows',
      sheetId: 'sh1',
      row: 9,
      count: 1,
    })
    expect(error).toContain('deleted rows')
  })

  it('ignores formulas that are deleted along with their own row or column', async () => {
    const delRow7 = { op: 'delete_rows', sheetId: 'sh1', row: 7, count: 1 } as const
    // A7 holds a per-row formula: deleting row 7 removes the formula too
    stubFormulas({ sh1: ['=B7*C7'] })
    expect(await structuralDeleteFormulaError(state(), workbook, delRow7)).toBeNull()
    // A8 refers only to row 7 and survives the deletion
    stubFormulas({ sh1: ['=B7*C7', '=B7*C7'] })
    expect(await structuralDeleteFormulaError(state(), workbook, delRow7)).toContain('B7')
    // the same cell on another sheet survives as well
    stubFormulas({ sh1: [], sh2: ['=Data!B7*Data!C7'] })
    expect(await structuralDeleteFormulaError(state(), workbook, delRow7)).not.toBeNull()
    // a column: A7 sums its own column, which is the one deleted
    stubFormulas({ sh1: ['=SUM(A1:A5)'] })
    const delA = { op: 'delete_cols', sheetId: 'sh1', column: 'A', count: 1 } as const
    expect(await structuralDeleteFormulaError(state(), workbook, delA)).toBeNull()
  })

  it('ignores shared-formula followers and journal formulas inside the deleted rows', async () => {
    // D2:D5 filled down from D2 =B2*C2; deleting rows 3-4 removes two followers
    stubFormulas(
      { sh1: [] },
      {
        sh1: [
          {
            si: 0,
            row: 1,
            column: 3,
            formula: '=B2*C2',
            range: { startRow: 1, endRow: 4, startColumn: 3, endColumn: 3 },
          },
        ],
      },
    )
    const journaled = state({
      editJournal: {
        cells: new Map([
          ['sh1', new Map([['2:4', { row: 2, column: 4, hasValue: true, formula: '=B3' }]])],
        ]),
        structuralOps: new Map(),
        sheets: { added: new Set(), removed: new Set() },
      },
    })
    expect(
      await structuralDeleteFormulaError(journaled, workbook, {
        op: 'delete_rows',
        sheetId: 'sh1',
        row: 3,
        count: 2,
      }),
    ).toBeNull()
  })

  it('only considers qualified references on other sheets', async () => {
    stubFormulas({ sh1: [], sh2: ['=SUM(D7:D9)'] })
    expect(await structuralDeleteFormulaError(state(), workbook, delD)).toBeNull()
    stubFormulas({ sh1: [], sh2: ['=SUM(Data!$D$7:$D$9)'] })
    expect(await structuralDeleteFormulaError(state(), workbook, delD)).not.toBeNull()
  })

  it('a style-only journal entry does not hide the file formula', async () => {
    stubFormulas({ sh1: ['=SUM($D$7:$D$9)'] })
    const styled = state({
      editJournal: {
        cells: new Map([
          [
            'sh1',
            new Map([['6:0', { row: 6, column: 0, hasValue: false, style: { bold: true } }]]),
          ],
        ]),
        structuralOps: new Map(),
        sheets: { added: new Set(), removed: new Set() },
      },
    })
    expect(await structuralDeleteFormulaError(styled, workbook, delD)).not.toBeNull()
    // a content overwrite at the same cell DOES supersede it
    const overwritten = state({
      editJournal: {
        cells: new Map([
          ['sh1', new Map([['6:0', { row: 6, column: 0, hasValue: true, value: 1 }]])],
        ]),
        structuralOps: new Map(),
        sheets: { added: new Set(), removed: new Set() },
      },
    })
    expect(await structuralDeleteFormulaError(overwritten, workbook, delD)).toBeNull()
  })

  it('checks session-written journal formulas too', async () => {
    stubFormulas({ sh1: [] })
    const journaled = state({
      editJournal: {
        cells: new Map([
          ['sh1', new Map([['1:1', { row: 1, column: 1, formula: '=SUM($D$7:$D$9)' }]])],
        ]),
        structuralOps: new Map(),
        sheets: { added: new Set(), removed: new Set() },
      },
    })
    expect(await structuralDeleteFormulaError(journaled, workbook, delD)).not.toBeNull()
  })

  it('falls back to the save-time guard when texts are unreliable', async () => {
    stubFormulas({ sh1: ['=SUM($D$7:$D$9)'] })
    const shifted = state({
      editJournal: {
        cells: new Map(),
        structuralOps: new Map([['sh1', [{ kind: 'insert-rows', index: 0, count: 1 }]]]),
        sheets: { added: new Set(), removed: new Set() },
      },
    })
    expect(await structuralDeleteFormulaError(shifted, workbook, delD)).toBeNull()
    ;(globalThis as { window?: { desktopApi: { readWorkbookFormulas: unknown } } }).window = {
      desktopApi: {
        readWorkbookFormulas: async () => ({
          cells: [],
          sharedGroups: [],
          truncated: true,
          indexingComplete: true,
        }),
      },
    }
    expect(await structuralDeleteFormulaError(state(), workbook, delD)).toBeNull()
  })
})

describe('structuralDeleteFormulaErrorSync (UI gate)', () => {
  function syncWorkbook(formulasBySheet: Record<string, string[][]>) {
    return {
      getSheets: () =>
        Object.keys(formulasBySheet).map((sheetId) => ({
          getSheetId: () => sheetId,
          getSheetName: () => (sheetId === 'sh1' ? 'Data' : 'Other'),
          getMaxRows: () => 20,
          getMaxColumns: () => 8,
          getRange: () => ({ getFormulas: () => formulasBySheet[sheetId] ?? [[]] }),
        })),
    }
  }

  it('scans the live model in full-load mode', () => {
    const wb = syncWorkbook({ sh1: [['', '=SUM($D$7:$D$9)']], sh2: [[]] })
    const full = state({ formulaMode: true })
    expect(structuralDeleteFormulaErrorSync(full, wb, delD)).toContain('deleted columns')
    const wbOk = syncWorkbook({ sh1: [['', '=SUM(E7:E9)']], sh2: [[]] })
    expect(structuralDeleteFormulaErrorSync(full, wbOk, delD)).toBeNull()
  })

  it('ignores formulas deleted along with their own row', () => {
    // D2:D4 hold =B2*C2, =B3*C3, =B4*C4; D5 totals them
    const perRow = Array.from({ length: 5 }, (_, row) => [
      '',
      '',
      '',
      row === 0 ? '' : row === 4 ? '=SUM(D2:D4)' : `=B${row + 1}*C${row + 1}`,
    ])
    const wb = syncWorkbook({ sh1: perRow, sh2: [[]] })
    const full = state({ formulaMode: true })
    const delRow3 = { op: 'delete_rows', sheetId: 'sh1', row: 3, count: 1 } as const
    expect(structuralDeleteFormulaErrorSync(full, wb, delRow3)).toBeNull()
    const streamed = state({
      formulaMode: false,
      formulaText: new Map([['sh1', new Map([['2:3', '=B3*C3']])]]),
      sharedFormulaGroups: new Map(),
    })
    expect(structuralDeleteFormulaErrorSync(streamed, wb, delRow3)).toBeNull()
    // a formula in another row that refers only to row 3 is still flagged
    perRow[3]![3] = '=B3*C3'
    expect(
      structuralDeleteFormulaErrorSync(full, syncWorkbook({ sh1: perRow }), delRow3),
    ).toContain('deleted rows')
  })

  it('keeps checking in full-load mode after session structural edits', () => {
    // the live model already reflects the shift — only the streamed
    // harvested-index path must fall back to the save-time guard
    const wb = syncWorkbook({ sh1: [['', '=SUM($D$7:$D$9)']], sh2: [[]] })
    const shifted = state({
      formulaMode: true,
      editJournal: {
        cells: new Map(),
        structuralOps: new Map([['sh1', [{ kind: 'insert-rows', index: 0, count: 1 }]]]),
        sheets: { added: new Set(), removed: new Set() },
      },
    })
    expect(structuralDeleteFormulaErrorSync(shifted, wb, delD)).toContain('deleted columns')
    const streamedShifted = state({
      formulaMode: false,
      formulaText: new Map([['sh1', new Map([['6:0', '=SUM($D$7:$D$9)']])]]),
      sharedFormulaGroups: new Map(),
      editJournal: {
        cells: new Map(),
        structuralOps: new Map([['sh1', [{ kind: 'insert-rows', index: 0, count: 1 }]]]),
        sheets: { added: new Set(), removed: new Set() },
      },
    })
    expect(structuralDeleteFormulaErrorSync(streamedShifted, wb, delD)).toBeNull()
  })

  it('uses the harvested formula index on streamed workbooks', () => {
    const wb = syncWorkbook({ sh1: [[]], sh2: [[]] })
    const streamed = state({
      formulaMode: false,
      formulaText: new Map([['sh1', new Map([['6:0', '=SUM($D$7:$D$9)']])]]),
      sharedFormulaGroups: new Map(),
    })
    expect(structuralDeleteFormulaErrorSync(streamed, wb, delD)).toContain('deleted columns')
    // a content overwrite at that cell supersedes the harvested text
    const overwritten = state({
      formulaMode: false,
      formulaText: new Map([['sh1', new Map([['6:0', '=SUM($D$7:$D$9)']])]]),
      sharedFormulaGroups: new Map(),
      editJournal: {
        cells: new Map([
          ['sh1', new Map([['6:0', { row: 6, column: 0, hasValue: true, value: 1 }]])],
        ]),
        structuralOps: new Map(),
        sheets: { added: new Set(), removed: new Set() },
      },
    })
    expect(structuralDeleteFormulaErrorSync(overwritten, wb, delD)).toBeNull()
  })
})
