import { describe, expect, it } from 'vitest'

import { externalLinkOverlayCell } from '../src/renderer/external-link-values'
import { recalculatedFormulaValues } from '../src/renderer/save-actions'
import { applyExternalLinkValues } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'
import type { ExternalLinkCell } from '../src/shared/desktop-api'

/// The overlay/state shape applyExternalLinkValues touches, no Univer needed:
/// the sheet exists by name, and an empty loaded-range map skips painting.
function emptyState(): LazyWorkbookState {
  return {
    editJournal: { cells: new Map(), sheets: { removed: new Set() } },
    recalc: { overlay: new Map(), incompleteRuns: new Set() },
    loadedRanges: new Map(),
  } as unknown as LazyWorkbookState
}

const runtime = {
  univerAPI: {
    getActiveWorkbook: () => ({
      getSheets: () => [{ getSheetName: () => 'Sheet1', getSheetId: () => 'sheet-1' }],
    }),
  },
} as never

function cell(overrides: Partial<ExternalLinkCell>): ExternalLinkCell {
  return {
    sheet: 'Sheet1',
    row: 0,
    column: 0,
    formatted: '',
    formula: '[1]Sheet1!A1',
    ...overrides,
  }
}

describe('externalLinkOverlayCell', () => {
  it('keeps a number numeric', () => {
    expect(externalLinkOverlayCell(cell({ kind: 'number', number: 20, formatted: '20' }))).toEqual({
      v: 20,
    })
  })

  it('keeps a boolean a boolean instead of the text TRUE', () => {
    expect(externalLinkOverlayCell(cell({ kind: 'boolean', formatted: 'TRUE' }))).toEqual({
      v: true,
    })
    expect(externalLinkOverlayCell(cell({ kind: 'boolean', formatted: 'FALSE' }))).toEqual({
      v: false,
    })
  })

  it('keeps an error an error instead of a text cache', () => {
    expect(externalLinkOverlayCell(cell({ kind: 'error', formatted: '#DIV/0!' }))).toEqual({
      v: '#DIV/0!',
      isError: true,
    })
  })

  it('keeps text text, empty string included', () => {
    expect(externalLinkOverlayCell(cell({ kind: 'text', formatted: 'Total!' }))).toEqual({
      v: 'Total!',
    })
    expect(externalLinkOverlayCell(cell({ kind: 'text', formatted: '' }))).toEqual({ v: '' })
  })

  it('has no value without a kind (the file never cached one)', () => {
    expect(externalLinkOverlayCell(cell({}))).toBeNull()
  })
})

describe('applying the report into a save request', () => {
  it('round-trips boolean and error cells into the request formula values', () => {
    const state = emptyState()
    applyExternalLinkValues(state, runtime, {
      cells: [
        cell({ row: 0, column: 0, kind: 'boolean', formatted: 'TRUE' }),
        cell({ row: 0, column: 1, kind: 'error', formatted: '#DIV/0!' }),
        cell({ row: 0, column: 2, kind: 'number', number: 20, formatted: '20' }),
        cell({ row: 0, column: 3, kind: 'text', formatted: 'Total!' }),
        // no kind: the file never cached a value, nothing to write
        cell({ row: 0, column: 4 }),
      ],
    })
    const values = recalculatedFormulaValues(state)
    expect(values).toEqual([
      { sheetId: 'sheet-1', row: 0, column: 0, value: true },
      { sheetId: 'sheet-1', row: 0, column: 1, value: { error: '#DIV/0!' } },
      { sheetId: 'sheet-1', row: 0, column: 2, value: 20 },
      { sheetId: 'sheet-1', row: 0, column: 3, value: 'Total!' },
    ])
  })

  it('counts the cells it applied and keeps existing overlay cells', () => {
    const state = emptyState()
    state.recalc.overlay.set('sheet-1', new Map([['9:9', { v: 1 }]]))
    const result = applyExternalLinkValues(state, runtime, {
      cells: [cell({ row: 0, column: 0, kind: 'boolean', formatted: 'TRUE' })],
    })
    expect(result).toEqual({ applied: 1, sheets: 1 })
    const overlay = state.recalc.overlay.get('sheet-1')!
    expect(overlay.get('9:9')).toEqual({ v: 1 })
    expect(overlay.get('0:0')).toEqual({ v: true })
  })
})
