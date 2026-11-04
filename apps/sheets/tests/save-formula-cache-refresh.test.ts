import { describe, expect, it } from 'vitest'

import { liveFormulaCacheValues } from '../src/renderer/save-actions'
import type { LazyWorkbookState } from '../src/renderer/univer-state'
import type { CellState } from '@genoffice/xlsx-gateway/domain/workbook.types'

/// The state slice liveFormulaCacheValues reads, no Univer needed: the
/// journal's cells, the formula-text harvest and the file cache it pairs with.
function state(parts: {
  journal?: Map<string, Map<string, { formula?: string; value?: string | number | boolean | null }>>
  removed?: Set<string>
  formulas?: Map<string, Map<string, string>>
  cached?: Map<string, Map<string, string | number | boolean>>
  overlay?: Map<string, Map<string, unknown>>
  external?: Map<string, Map<string, unknown>>
}): LazyWorkbookState {
  return {
    editJournal: {
      cells: parts.journal ?? new Map(),
      sheets: { removed: parts.removed ?? new Set() },
    },
    recalc: {
      overlay: parts.overlay ?? new Map(),
      externalOverlay: parts.external ?? new Map(),
      incompleteRuns: new Set(),
    },
    formulaText: parts.formulas ?? new Map(),
    cachedFormulaValues: parts.cached ?? new Map(),
  } as unknown as LazyWorkbookState
}

/// A model answer for one address; `formula` omitted mimics a withheld cell,
/// which answers with a placeholder and no formula text.
function cellOf(value: string | number | boolean | null, formula?: string): CellState {
  return formula === undefined ? { value } : { value, formula, rawValue: value }
}

function reader(bySheet: Map<string, Record<string, CellState>>): ((
  addresses: string[],
  sheetId: string,
) => Record<string, CellState>) & {
  asked: string[]
} {
  const asked: string[] = []
  const read = (addresses: string[], sheetId: string) => {
    for (const address of addresses) asked.push(`${sheetId}!${address}`)
    return bySheet.get(sheetId) ?? {}
  }
  return Object.assign(read, { asked })
}

describe('liveFormulaCacheValues', () => {
  it('reads a journal-owned formula cell even though the file has no cache for it', () => {
    const journal = new Map([['sheet-1', new Map([['2:2', { formula: '=A2+1', value: null }]])]])
    const read = reader(new Map([['sheet-1', { '2:2': cellOf(6, '=A2+1') }]]))
    expect(liveFormulaCacheValues(state({ journal }), read)).toEqual([
      { sheetId: 'sheet-1', row: 2, column: 2, value: 6 },
    ])
  })

  it('refreshes a dependent whose cached value went stale after a precedent edit', () => {
    const read = reader(new Map([['sheet-1', { '1:1': cellOf(10, '=A2*2') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2*2']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 6]])]]),
        }),
        read,
      ),
    ).toEqual([{ sheetId: 'sheet-1', row: 1, column: 1, value: 10 }])
  })

  it('keeps the file cache when the live value agrees with it', () => {
    const read = reader(new Map([['sheet-1', { '1:1': cellOf(6, '=A2*2') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2*2']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 6]])]]),
        }),
        read,
      ),
    ).toEqual([])
  })

  it('tolerates a last-ulp difference from a foreign engine but not a real change', () => {
    const formulas = new Map([['sheet-1', new Map([['1:1', '=SUM(A2:A9)']])]])
    const cached = new Map([['sheet-1', new Map([['1:1', 6]])]])
    const near = reader(new Map([['sheet-1', { '1:1': cellOf(6 + 1e-12, '=SUM(A2:A9)') }]]))
    expect(liveFormulaCacheValues(state({ formulas, cached }), near)).toEqual([])
    const changed = reader(new Map([['sheet-1', { '1:1': cellOf(6.1, '=SUM(A2:A9)') }]]))
    expect(liveFormulaCacheValues(state({ formulas, cached }), changed)).toEqual([
      { sheetId: 'sheet-1', row: 1, column: 1, value: 6.1 },
    ])
  })

  it('skips a cell whose model answer carries no formula (withheld or mid-commit)', () => {
    const journal = new Map([['sheet-1', new Map([['0:0', { formula: '=X', value: null }]])]])
    const read = reader(new Map([['sheet-1', { '0:0': { value: '•••' } }]]))
    expect(liveFormulaCacheValues(state({ journal }), read)).toEqual([])
    const compare = reader(new Map([['sheet-1', { '0:0': { value: '•••' } }]]))
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([['sheet-1', new Map([['0:0', '=X']])]]),
          cached: new Map([['sheet-1', new Map([['0:0', 1]])]]),
        }),
        compare,
      ),
    ).toEqual([])
  })

  it('skips comparison cells the recalc or external-link overlays already contribute', () => {
    const read = reader(
      new Map([
        ['sheet-1', { '1:1': cellOf(10, '=A2*2') }],
        ['sheet-2', { '0:0': cellOf(20, '[1]Sheet1!A1') }],
      ]),
    )
    expect(
      liveFormulaCacheValues(
        state({
          formulas: new Map([
            ['sheet-1', new Map([['1:1', '=A2*2']])],
            ['sheet-2', new Map([['0:0', '[1]Sheet1!A1']])],
          ]),
          cached: new Map([
            ['sheet-1', new Map([['1:1', 6]])],
            ['sheet-2', new Map([['0:0', 6]])],
          ]),
          overlay: new Map([['sheet-1', new Map([['1:1', { v: 10 }]])]]),
          external: new Map([['sheet-2', new Map([['0:0', { v: 20 }]])]]),
        }),
        read,
      ),
    ).toEqual([])
    expect(read.asked).toEqual([])
  })

  it('a journal constant still owns its cell: no comparison value rides over it', () => {
    const journal = new Map([['sheet-1', new Map([['1:1', { value: 5 }]])]])
    const read = reader(new Map([['sheet-1', { '1:1': cellOf(10, '=A2*2') }]]))
    expect(
      liveFormulaCacheValues(
        state({
          journal,
          formulas: new Map([['sheet-1', new Map([['1:1', '=A2*2']])]]),
          cached: new Map([['sheet-1', new Map([['1:1', 6]])]]),
        }),
        read,
      ),
    ).toEqual([])
  })

  it('caches an error literal as one, skips the engine-internal failure', () => {
    const journal = new Map([
      ['sheet-1', new Map([['0:0', { formula: '=1/0', value: null }]])],
      ['sheet-2', new Map([['0:0', { formula: '=BROKEN()', value: null }]])],
    ])
    const read = reader(
      new Map([
        ['sheet-1', { '0:0': cellOf('#DIV/0!', '=1/0') }],
        ['sheet-2', { '0:0': cellOf('#ERROR!', '=BROKEN()') }],
      ]),
    )
    expect(liveFormulaCacheValues(state({ journal }), read)).toEqual([
      { sheetId: 'sheet-1', row: 0, column: 0, value: { error: '#DIV/0!' } },
    ])
  })

  it('skips comparison cells whose saved form the session never saw', () => {
    const read = reader(new Map([['sheet-1', { '3:3': cellOf(10, '=A2*2') }]]))
    expect(
      liveFormulaCacheValues(
        state({ formulas: new Map([['sheet-1', new Map([['3:3', '=A2*2']])]]) }),
        read,
      ),
    ).toEqual([])
  })

  it('skips a removed sheet entirely', () => {
    const journal = new Map([['sheet-1', new Map([['0:0', { formula: '=A2', value: null }]])]])
    const read = reader(new Map([['sheet-1', { '0:0': cellOf(1, '=A2') }]]))
    expect(liveFormulaCacheValues(state({ journal, removed: new Set(['sheet-1']) }), read)).toEqual(
      [],
    )
    expect(read.asked).toEqual([])
  })
})
