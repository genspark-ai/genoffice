/**
 * A sort (reorder-range) moves whole cells in Univer, style ids included, but
 * the journal used to snapshot values only — the saved file kept every fill
 * at its pre-sort row while the screen showed them moved.
 */
import { describe, expect, it } from 'vitest'

import { createEditJournal } from '../src/renderer/edit-journal'
import { journalRangeSnapshot } from '../src/renderer/univer-sync'
import type { LazyWorkbookState, UniverRuntime } from '../src/renderer/univer-state'

const POOL: Record<string, Record<string, unknown>> = {
  red: { bg: { rgb: '#FF0000' } },
  blue: { bg: { rgb: '#0000FF' }, bl: 1 },
}

function fakeRuntime(cells: ({ v: unknown; s?: unknown } | null)[][]): UniverRuntime {
  const worksheet = {
    getRange: () => ({ getCellDatas: () => cells }),
    getSheet: () => ({ getCellMatrix: () => null }),
  }
  const workbook = {
    getSheetBySheetId: () => worksheet,
    getWorkbook: () => ({
      getStyles: () => ({
        getStyleByCell: (cell: { s?: unknown }) =>
          typeof cell.s === 'string' ? POOL[cell.s] : cell.s,
      }),
    }),
  }
  return {
    univerAPI: { getActiveWorkbook: () => workbook },
    univer: { __getInjector: () => ({ get: () => ({ moveFormulaRefOffset: (f: string) => f }) }) },
  } as unknown as UniverRuntime
}

function snapshot(cells: ({ v: unknown; s?: unknown } | null)[][], order?: Record<number, number>) {
  const journal = createEditJournal()
  const state = { editJournal: journal } as LazyWorkbookState
  journalRangeSnapshot(
    fakeRuntime(cells),
    state,
    'sheet-1',
    { startRow: 0, endRow: cells.length - 1, startColumn: 0, endColumn: 0 },
    order,
  )
  const entries = journal.cells.get('sheet-1')!
  return cells.map((_, row) => entries.get(`${row}:0`))
}

describe('journalRangeSnapshot', () => {
  it('moves resolved styles with sorted rows and resets the rows they left', () => {
    // pre-sort: a(red) b(-) c(blue); sorted descending → c b a
    const entries = snapshot([[{ v: 'c', s: 'blue' }], [{ v: 'b' }], [{ v: 'a', s: 'red' }]], {
      0: 2,
      1: 1,
      2: 0,
    })
    expect(entries[0]).toMatchObject({
      value: 'c',
      styleReset: true,
      style: { fillColor: '#0000FF', bold: true },
    })
    expect(entries[1]).toMatchObject({ value: 'b' })
    expect(entries[1]?.styleReset).toBeUndefined()
    expect(entries[1]?.style).toBeUndefined()
    expect(entries[2]).toMatchObject({
      value: 'a',
      styleReset: true,
      style: { fillColor: '#FF0000' },
    })
  })

  it('leaves uniformly styled rows alone so file formatting survives untouched', () => {
    const entries = snapshot([[{ v: 2, s: 'red' }], [{ v: 1, s: 'red' }]], { 0: 1, 1: 0 })
    for (const entry of entries) {
      expect(entry?.styleReset).toBeUndefined()
      expect(entry?.style).toBeUndefined()
    }
  })

  it('resets a target row when an unstyled cell lands on a styled one', () => {
    const entries = snapshot([[{ v: 'b' }], [{ v: 'a', s: 'red' }]], { 0: 1, 1: 0 })
    expect(entries[0]).toMatchObject({ value: 'b', styleReset: true })
    expect(entries[0]?.style).toBeUndefined()
  })

  it('keeps a style-only cell that moved instead of deleting it', () => {
    const entries = snapshot([[{ v: null, s: 'red' }], [{ v: 'x' }]], { 0: 1, 1: 0 })
    expect(entries[0]).toMatchObject({ styleReset: true, style: { fillColor: '#FF0000' } })
  })

  it('treats every cell of a moved range as replaced when no order is known', () => {
    const entries = snapshot([[{ v: 'a', s: 'red' }], [{ v: 'b' }]])
    expect(entries[0]).toMatchObject({
      value: 'a',
      styleReset: true,
      style: { fillColor: '#FF0000' },
    })
    expect(entries[1]).toMatchObject({ value: 'b', styleReset: true })
  })
})
