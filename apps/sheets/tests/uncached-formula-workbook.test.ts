import { afterEach, describe, expect, it } from 'vitest'

import { activateFormulaClosure, truncatedIndexRetiresEngine } from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'

const sheetMeta = (id: string, name: string) => ({
  id,
  name,
  rowCount: 10,
  columnCount: 5,
  tables: [],
  pivotTables: [],
  freeze: null,
  columnWidths: [],
})

function cacheModeState(): LazyWorkbookState {
  return {
    formulaMode: false,
    file: {
      sessionId: 'session-1',
      styles: [],
      sheets: [sheetMeta('sheet-1', 'Dashboard'), sheetMeta('sheet-2', 'Journal')],
    },
    editJournal: { structuralOps: new Map(), cells: new Map() },
    closure: { status: 'idle', pinned: new Map() },
    formulaText: new Map(),
    cachedFormulaValues: new Map(),
    recalc: {
      timer: null,
      generation: 0,
      failures: 0,
      engineOverBudget: false,
      formulaCells: new Map(),
      overlay: new Map(),
      follow: new Map(),
      running: false,
      lastRunAt: 0,
    },
  } as unknown as LazyWorkbookState
}

type FormulaCell = { row: number; column: number; formula: string; value?: string | number | null }

/// Sheet 1 has a short complete index; sheet 2 is the 100k+ formula sheet
/// whose index comes back truncated.
function stubDesktopApi(small: FormulaCell[], truncated: FormulaCell[]): void {
  ;(globalThis as { window?: unknown }).window = {
    desktopApi: {
      readWorkbookFormulas: async ({ sheetId }: { sheetId: string }) =>
        sheetId === 'sheet-1'
          ? { truncated: false, indexingComplete: true, cells: small }
          : { truncated: true, indexingComplete: false, cells: truncated },
    },
  }
}

const runtime = {
  univerAPI: { getActiveWorkbook: () => ({ getActiveSheet: () => undefined }) },
} as never

describe('truncatedIndexRetiresEngine', () => {
  it('retires the engine once any formula cell carries a cached value', () => {
    expect(
      truncatedIndexRetiresEngine({
        cachedFormulaValues: new Map([['sheet-2', new Map([['0:1', 12]])]]),
      }),
    ).toBe(true)
  })

  it('keeps the engine for a file without a single cached value', () => {
    expect(truncatedIndexRetiresEngine({ cachedFormulaValues: new Map() })).toBe(false)
    expect(
      truncatedIndexRetiresEngine({ cachedFormulaValues: new Map([['sheet-1', new Map()]]) }),
    ).toBe(false)
  })
})

describe('activateFormulaClosure over a truncated formula index', () => {
  let state: LazyWorkbookState | null = null

  afterEach(() => {
    if (state?.recalc.timer) clearTimeout(state.recalc.timer)
    state = null
    delete (globalThis as { window?: unknown }).window
  })

  it('leaves the recalc fallback enabled when no formula cell has a cached value', async () => {
    state = cacheModeState()
    stubDesktopApi(
      [{ row: 6, column: 2, formula: 'Journal!B3', value: null }],
      [
        { row: 0, column: 1, formula: 'A1*2', value: null },
        { row: 1, column: 1, formula: 'A2*2' },
      ],
    )

    await activateFormulaClosure(runtime, { current: state }, () => undefined)

    expect(state.closure.status).toBe('unavailable')
    expect(state.recalc.engineOverBudget).toBe(false)
    // the fallback recalc is queued right away, not at the first edit
    expect(state.recalc.timer).not.toBeNull()
  })

  it('retires the engine when the file carries cached values', async () => {
    state = cacheModeState()
    stubDesktopApi(
      [{ row: 6, column: 2, formula: 'Journal!B3', value: 700 }],
      [{ row: 0, column: 1, formula: 'A1*2', value: null }],
    )

    await activateFormulaClosure(runtime, { current: state }, () => undefined)

    expect(state.closure.status).toBe('unavailable')
    expect(state.recalc.engineOverBudget).toBe(true)
    expect(state.recalc.timer).toBeNull()
  })

  it('counts caches on the truncated sheet itself', async () => {
    state = cacheModeState()
    stubDesktopApi(
      [{ row: 6, column: 2, formula: 'Journal!B3', value: null }],
      [{ row: 0, column: 1, formula: 'A1*2', value: 4 }],
    )

    await activateFormulaClosure(runtime, { current: state }, () => undefined)

    expect(state.recalc.engineOverBudget).toBe(true)
  })

  it('a size-gated workbook stays over budget regardless of caches', async () => {
    state = cacheModeState()
    state.recalc.engineOverBudget = true
    stubDesktopApi([], [{ row: 0, column: 1, formula: 'A1*2', value: null }])

    await activateFormulaClosure(runtime, { current: state }, () => undefined)

    expect(state.recalc.engineOverBudget).toBe(true)
    expect(state.recalc.timer).toBeNull()
  })
})
