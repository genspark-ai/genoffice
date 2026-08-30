import { afterEach, describe, expect, it } from 'vitest'

import {
  structuralDeleteFormulaError,
  structuralDeleteFormulaErrorSync,
} from '../src/renderer/plan-operations'
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

function stubFormulas(bySheet: Record<string, string[]>): void {
  ;(globalThis as { window?: unknown }).window = {
    desktopApi: {
      readWorkbookFormulas: async ({ sheetId }: { sheetId: string }) => ({
        cells: (bySheet[sheetId] ?? []).map((formula, i) => ({
          row: 6 + i,
          column: 0,
          formula,
        })),
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

  it('passes formulas that merely shift or shrink', async () => {
    stubFormulas({ sh1: ['=SUM(E7:E9)', '=SUM(C7:F9)'] })
    expect(await structuralDeleteFormulaError(state(), workbook, delD)).toBeNull()
  })

  it('flags deleted rows the same way', async () => {
    stubFormulas({ sh1: ['=SUM($A$7:$C$7)'] })
    const error = await structuralDeleteFormulaError(state(), workbook, {
      op: 'delete_rows',
      sheetId: 'sh1',
      row: 7,
      count: 1,
    })
    expect(error).toContain('deleted rows')
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
    })
    expect(structuralDeleteFormulaErrorSync(streamed, wb, delD)).toContain('deleted columns')
    // a content overwrite at that cell supersedes the harvested text
    const overwritten = state({
      formulaMode: false,
      formulaText: new Map([['sh1', new Map([['6:0', '=SUM($D$7:$D$9)']])]]),
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
