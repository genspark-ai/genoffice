/**
 * Shared-formula followers (`{si, v}` with no `f` — tiled paste / fill) must
 * journal as real formulas, materialized from the group master by offset —
 * otherwise the save writes them as plain values at best, and the follow-up
 * recalc mutation (`{v: null}`) wipes them entirely (alpha: copy C2:E2 →
 * tile-paste C3:E21 → save lost the block).
 */
import { describe, expect, it, vi } from 'vitest'
import { createEditJournal, recordSetRangeValues } from '../src/renderer/edit-journal'
import { makeSharedFormulaResolver } from '../src/renderer/shared-formula-journal'

type Cell = { f?: string; si?: string; v?: unknown }

function fakeMatrix(cells: Record<number, Record<number, Cell>>) {
  return {
    forValue(cb: (row: number, column: number, cell: unknown) => boolean | void) {
      for (const [row, columns] of Object.entries(cells)) {
        for (const [column, cell] of Object.entries(columns)) {
          if (cb(Number(row), Number(column), cell) === false) return this
        }
      }
      return this
    },
  }
}

/** offset marker instead of real ref-shifting — the args are what matters */
const fakeMove = (formula: string, dx: number, dy: number) => `${formula}@${dx},${dy}`

describe('makeSharedFormulaResolver', () => {
  const matrix = fakeMatrix({
    0: { 2: { f: '=A1&"-x"', si: 'grp1', v: '10-x' } },
    1: { 2: { si: 'grp1', v: '10-x' } },
  })

  it('materializes a follower by shifting the master formula', () => {
    const resolve = makeSharedFormulaResolver(() => matrix, fakeMove)
    expect(resolve(1, 2, 'grp1')).toBe('=A1&"-x"@0,1')
    expect(resolve(5, 4, 'grp1')).toBe('=A1&"-x"@2,5')
  })

  it('returns the master formula unshifted for the master cell itself', () => {
    const resolve = makeSharedFormulaResolver(() => matrix, fakeMove)
    expect(resolve(0, 2, 'grp1')).toBe('=A1&"-x"')
  })

  it('scans the matrix once per group', () => {
    const getMatrix = vi.fn(() => matrix)
    const resolve = makeSharedFormulaResolver(getMatrix, fakeMove)
    resolve(1, 2, 'grp1')
    resolve(2, 2, 'grp1')
    resolve(3, 2, 'grp1')
    expect(getMatrix).toHaveBeenCalledTimes(1)
  })

  it('returns null when the group has no master (and caches the miss)', () => {
    const getMatrix = vi.fn(() => matrix)
    const resolve = makeSharedFormulaResolver(getMatrix, fakeMove)
    expect(resolve(1, 2, 'ghost')).toBeNull()
    expect(resolve(2, 2, 'ghost')).toBeNull()
    expect(getMatrix).toHaveBeenCalledTimes(1)
  })
})

describe('recordSetRangeValues with a shared-formula resolver', () => {
  const resolver = makeSharedFormulaResolver(
    () =>
      fakeMatrix({
        0: { 2: { f: '=A1&"-x"', si: 'grp1' } },
      }),
    // realistic enough shift for A1-style relative refs in this test
    (formula, _dx, dy) => formula.replace('A1', `A${1 + dy}`),
  )

  it('journals an si-only follower as its materialized formula', () => {
    const journal = createEditJournal()
    recordSetRangeValues(
      journal,
      'sheet-1',
      { 1: { 2: { si: 'grp1', v: '10-x', t: 1 } } },
      resolver,
    )
    const entry = journal.cells.get('sheet-1')?.get('1:2')
    expect(entry?.formula).toBe('=A2&"-x"')
    expect(entry?.hasValue).toBe(true)
  })

  it('keeps the materialized formula when the recalc result arrives as {v}', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 1: { 2: { si: 'grp1', v: '10-x' } } }, resolver)
    // the formula engine follows up with a bare value mutation
    recordSetRangeValues(journal, 'sheet-1', { 1: { 2: { v: '10-x' } } }, resolver)
    const entry = journal.cells.get('sheet-1')?.get('1:2')
    expect(entry?.formula).toBe('=A2&"-x"')
  })

  it('prefers an inline formula over si resolution', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 0: { 2: { f: '=B1', si: 'grp1' } } }, resolver)
    expect(journal.cells.get('sheet-1')?.get('0:2')?.formula).toBe('=B1')
  })

  it('falls back to the plain value when the master is gone', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 1: { 2: { si: 'ghost', v: '10-x' } } }, resolver)
    const entry = journal.cells.get('sheet-1')?.get('1:2')
    expect(entry?.formula).toBeUndefined()
    expect(entry?.value).toBe('10-x')
  })

  it('behaves as before without a resolver', () => {
    const journal = createEditJournal()
    recordSetRangeValues(journal, 'sheet-1', { 1: { 2: { si: 'grp1', v: '10-x' } } })
    const entry = journal.cells.get('sheet-1')?.get('1:2')
    expect(entry?.formula).toBeUndefined()
    expect(entry?.value).toBe('10-x')
  })
})
