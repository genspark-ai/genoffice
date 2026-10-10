import { describe, expect, it } from 'vitest'

import { FULL_LOAD_MAX_CELLS, REORDER_JOURNAL_MAX_CELLS } from '../src/renderer/app-constants'
import { createEditJournal } from '../src/renderer/edit-journal'
import { SharedFormulaLookup } from '../src/renderer/shared-formula-index'
import {
  fullLoadGate,
  rangeHasStreamedFormulas,
  reorderCommandRanges,
  reorderGate,
  type ReorderGateInput,
} from '../src/renderer/reorder-gate'
import type { WorkbookSharedFormulaGroup } from '../src/shared/desktop-api'

function input(overrides: Partial<ReorderGateInput> = {}): ReorderGateInput {
  return {
    formulaMode: false,
    preloadComplete: true,
    preloadRunning: false,
    isAddedSheet: false,
    cellCounts: { gridCells: 60_000, storedCells: 60_000, valueCells: 60_000 },
    rangeCells: 60_000,
    rangeHasStreamedFormulas: () => false,
    ...overrides,
  }
}

describe('fullLoadGate', () => {
  const fits = { gridCells: 60_000, storedCells: 60_000, valueCells: 60_000 }
  const tooLarge = {
    gridCells: FULL_LOAD_MAX_CELLS + 1,
    storedCells: FULL_LOAD_MAX_CELLS + 1,
    valueCells: FULL_LOAD_MAX_CELLS + 1,
  }

  it('offers Load all on an idle streamed workbook that fits', () => {
    expect(fullLoadGate({ formulaMode: false, preloadRunning: false, cellCounts: fits })).toBe(
      'offerFullLoad',
    )
  })

  it('explains instead of offering above the full-load cap', () => {
    expect(fullLoadGate({ formulaMode: false, preloadRunning: false, cellCounts: tooLarge })).toBe(
      'workbookTooLarge',
    )
  })

  it('shows the in-progress notice while any preload runs, even one that Print started', () => {
    expect(fullLoadGate({ formulaMode: false, preloadRunning: true, cellCounts: fits })).toBe(
      'loading',
    )
    expect(fullLoadGate({ formulaMode: false, preloadRunning: true, cellCounts: tooLarge })).toBe(
      'loading',
    )
  })

  it('holds a formula-mode workbook at loading while it preloads itself at open', () => {
    expect(fullLoadGate({ formulaMode: true, preloadRunning: false, cellCounts: fits })).toBe(
      'loading',
    )
  })
})

describe('reorderGate', () => {
  it('allows a fully loaded value-mode workbook (>50k cells after Load all)', () => {
    expect(reorderGate(input())).toBe('allow')
  })

  it('allows a fully loaded formula-mode workbook', () => {
    expect(reorderGate(input({ formulaMode: true }))).toBe('allow')
  })

  it('never gates sheets added this session', () => {
    expect(
      reorderGate(
        input({ isAddedSheet: true, preloadComplete: false, rangeHasStreamedFormulas: () => true }),
      ),
    ).toBe('allow')
  })

  it('offers Load all on a streamed value-mode workbook that fits', () => {
    expect(reorderGate(input({ preloadComplete: false }))).toBe('offerFullLoad')
  })

  it('reports a workbook above the full-load cap as too large', () => {
    expect(
      reorderGate(
        input({
          preloadComplete: false,
          cellCounts: {
            gridCells: 0,
            storedCells: FULL_LOAD_MAX_CELLS + 1,
            valueCells: FULL_LOAD_MAX_CELLS + 1,
          },
        }),
      ),
    ).toBe('workbookTooLarge')
  })

  it('reports loading while the preload runs or formula mode auto-preloads', () => {
    expect(reorderGate(input({ preloadComplete: false, preloadRunning: true }))).toBe('loading')
    expect(reorderGate(input({ preloadComplete: false, formulaMode: true }))).toBe('loading')
  })

  it('blocks value-mode ranges holding sidecar-only formulas', () => {
    expect(reorderGate(input({ rangeHasStreamedFormulas: () => true }))).toBe('valueModeFormulas')
    expect(reorderGate(input({ formulaMode: true, rangeHasStreamedFormulas: () => true }))).toBe(
      'allow',
    )
  })

  it('refuses a range the journal snapshot would silently drop', () => {
    expect(reorderGate(input({ rangeCells: REORDER_JOURNAL_MAX_CELLS + 1 }))).toBe('rangeTooLarge')
    expect(reorderGate(input({ rangeCells: REORDER_JOURNAL_MAX_CELLS }))).toBe('allow')
  })
})

describe('reorderCommandRanges', () => {
  const extent = { rows: 100, columns: 10 }
  const range = { startRow: 0, endRow: 9, startColumn: 0, endColumn: 2 }

  it('reads sort `range` and move from/to', () => {
    expect(reorderCommandRanges('sheet.command.sort-range', { range }, undefined, extent)).toEqual([
      range,
    ])
    const fromRange = { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }
    const toRange = { startRow: 5, endRow: 6, startColumn: 0, endColumn: 0 }
    expect(
      reorderCommandRanges('sheet.command.move-range', { fromRange, toRange }, undefined, extent),
    ).toEqual([fromRange, toRange])
  })

  it('falls back to the selection, then to the whole sheet', () => {
    expect(reorderCommandRanges('sheet.command.sort-range', undefined, range, extent)).toEqual([
      range,
    ])
    expect(
      reorderCommandRanges(
        'sheet.command.sort-range',
        { range: { startRow: 0 } },
        undefined,
        extent,
      ),
    ).toEqual([{ startRow: 0, endRow: 99, startColumn: 0, endColumn: 9 }])
    expect(reorderCommandRanges('sheet.command.sort-range', undefined, undefined, null)).toEqual([])
  })

  it('extends shift commands to the sheet edge in the shift direction', () => {
    expect(
      reorderCommandRanges('sheet.command.insert-range-move-right', undefined, range, extent),
    ).toEqual([{ ...range, endColumn: 9 }])
    expect(
      reorderCommandRanges('sheet.command.delete-range-move-up', { range }, undefined, extent),
    ).toEqual([{ ...range, endRow: 99 }])
  })
})

describe('rangeHasStreamedFormulas', () => {
  function state(
    formulas: Record<string, string>,
    ops: unknown[] = [],
    shared?: WorkbookSharedFormulaGroup,
  ) {
    const editJournal = createEditJournal()
    if (ops.length > 0) editJournal.structuralOps.set('sh1', ops as never)
    return {
      editJournal,
      formulaText: new Map([['sh1', new Map(Object.entries(formulas))]]),
      formulaTextTruncated: new Set<string>(),
      sharedFormulaGroups: new Map(shared ? [['sh1', new SharedFormulaLookup([shared])]] : []),
    }
  }

  it('fails closed when the sheet formula index overflowed', () => {
    const s = state({})
    s.formulaTextTruncated.add('sh1')
    expect(
      rangeHasStreamedFormulas(s, 'sh1', [
        { startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 },
      ]),
    ).toBe(true)
  })

  it('finds a formula inside the range and ignores ones outside', () => {
    const s = state({ '5:1': '=A6*2' })
    expect(
      rangeHasStreamedFormulas(s, 'sh1', [
        { startRow: 0, endRow: 9, startColumn: 0, endColumn: 1 },
      ]),
    ).toBe(true)
    expect(
      rangeHasStreamedFormulas(s, 'sh1', [
        { startRow: 0, endRow: 4, startColumn: 0, endColumn: 1 },
      ]),
    ).toBe(false)
    expect(
      rangeHasStreamedFormulas(s, 'other', [
        { startRow: 0, endRow: 9, startColumn: 0, endColumn: 9 },
      ]),
    ).toBe(false)
  })

  it('sees shared-formula followers that the text index does not list', () => {
    const s = state({}, [], {
      si: 0,
      row: 0,
      column: 1,
      formula: '=A1*2',
      range: { startRow: 0, endRow: 99, startColumn: 1, endColumn: 1 },
    })
    expect(
      rangeHasStreamedFormulas(s, 'sh1', [
        { startRow: 50, endRow: 60, startColumn: 1, endColumn: 1 },
      ]),
    ).toBe(true)
    expect(
      rangeHasStreamedFormulas(s, 'sh1', [
        { startRow: 50, endRow: 60, startColumn: 2, endColumn: 2 },
      ]),
    ).toBe(false)
  })

  it('maps file coordinates through structural ops before testing', () => {
    // two rows inserted at the top: file row 5 shows at screen row 7
    const s = state({ '5:0': '=1' }, [{ kind: 'insert-rows', index: 0, count: 2 }])
    expect(
      rangeHasStreamedFormulas(s, 'sh1', [
        { startRow: 5, endRow: 6, startColumn: 0, endColumn: 0 },
      ]),
    ).toBe(false)
    expect(
      rangeHasStreamedFormulas(s, 'sh1', [
        { startRow: 7, endRow: 7, startColumn: 0, endColumn: 0 },
      ]),
    ).toBe(true)
  })
})
