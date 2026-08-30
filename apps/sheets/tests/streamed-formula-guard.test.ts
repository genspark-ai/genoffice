import { describe, expect, it, vi } from 'vitest'

import type { WorkbookOperation } from '../src/domain/workbook-dsl'
import { cellKey } from '../src/renderer/formula-closure'
import {
  carryCopyFormulasPlan,
  collectStreamedFormulaPrecedents,
  proposeOperations,
  streamedPinBudgetError,
  type PlanContext,
  type StreamedRefSheet,
} from '../src/renderer/plan-operations'
import { CLOSURE_MAX_CELLS, type LazyWorkbookState } from '../src/renderer/univer-state'

/// Streaming-mode formula handling: written formulas stay LIVE — the apply
/// path loads & pins their referenced file cells into the engine — within a
/// session budget shared with closure mode. Batches whose references exceed
/// the budget fail closed with alternatives (they would otherwise evaluate
/// against partially loaded data and display silently wrong results).

const sheets: StreamedRefSheet[] = [
  { id: 'd1', name: 'Data', fileExtent: { rows: 20_000, columns: 6 } },
  { id: 'small1', name: 'Small', fileExtent: { rows: 500, columns: 4 } },
  { id: 'new1', name: 'ja' },
]

function collect(
  formula: string,
  hostSheetId: string,
  alreadyPinned?: ReadonlyMap<string, ReadonlyMap<string, unknown>>,
): Map<string, Set<number>> {
  const out = new Map<string, Set<number>>()
  collectStreamedFormulaPrecedents(formula, hostSheetId, sheets, out, alreadyPinned)
  return out
}

const totalOf = (needs: ReadonlyMap<string, ReadonlySet<number>>): number =>
  [...needs.values()].reduce((sum, cells) => sum + cells.size, 0)

describe('collectStreamedFormulaPrecedents', () => {
  it('collects the referenced file-sheet cells', () => {
    const needs = collect('=SUM(Data!A1:A100)', 'new1')
    expect(needs.get('d1')?.size).toBe(100)
    expect(needs.get('d1')?.has(cellKey(0, 0))).toBe(true)
    expect(needs.get('d1')?.has(cellKey(99, 0))).toBe(true)
  })

  it('clamps whole-column refs to the file extent', () => {
    expect(totalOf(collect('=SUM(Small!A:A)', 'new1'))).toBe(500)
    expect(totalOf(collect('=SUM(Small!A2:A99999)', 'new1'))).toBe(499)
  })

  it('resolves unqualified refs against the host sheet', () => {
    const needs = collect('=SUM(A1:B10)', 'd1')
    expect(needs.get('d1')?.size).toBe(20)
  })

  it('skips session-added sheets and already-pinned cells', () => {
    expect(totalOf(collect('=SUM(ja!A1:A50000)', 'd1'))).toBe(0)
    expect(totalOf(collect('=SUM(A1:A50000)', 'new1'))).toBe(0)
    const pinned = new Map([['d1', new Map([['0:0', { v: 1 }]])]])
    const needs = collect('=SUM(Data!A1:A10)', 'new1', pinned)
    expect(needs.get('d1')?.size).toBe(9)
    expect(needs.get('d1')?.has(cellKey(0, 0))).toBe(false)
  })

  it('dedupes overlapping references and resolves quoted qualifiers', () => {
    const spaced: StreamedRefSheet[] = [
      { id: 'd1', name: 'My Data', fileExtent: { rows: 20_000, columns: 6 } },
    ]
    const out = new Map<string, Set<number>>()
    collectStreamedFormulaPrecedents(
      "=SUM('My Data'!A1:A50)+AVERAGE('My Data'!A1:A50)",
      'd1',
      spaced,
      out,
    )
    expect(out.get('d1')?.size).toBe(50)
  })

  it('resolves case-variant sheet qualifiers (Excel refs are case-insensitive)', () => {
    // Skipping a case-variant spelling would bypass the pin AND the budget,
    // letting the formula evaluate against the streamed viewport.
    const cased = collect('=SUM(DATA!A1:A50)', 'new1')
    expect(cased.get('d1')?.size).toBe(50)
    const quotedCased = collect("=SUM('dAtA'!A1:A50)", 'new1')
    expect(quotedCased.get('d1')?.size).toBe(50)
  })

  it('stops materializing keys just past the session budget', () => {
    // 20000×6 = 120k referenced cells — collection must cap out instead of
    // building the full key set (the budget check rejects the batch anyway).
    const needs = collect('=FILTER(Data!A1:F20000,Data!D1:D20000="ja")', 'new1')
    expect(totalOf(needs)).toBeGreaterThan(CLOSURE_MAX_CELLS)
    expect(totalOf(needs)).toBeLessThanOrEqual(CLOSURE_MAX_CELLS + 2)
  })
})

function stateWithPinned(pinnedCount: number): LazyWorkbookState {
  const pinned = new Map<string, { v: number }>()
  for (let index = 0; index < pinnedCount; index += 1) pinned.set(`${index}:0`, { v: index })
  return {
    closure: { status: 'unavailable', pinned: new Map([['d1', pinned]]) },
  } as unknown as LazyWorkbookState
}

describe('streamedPinBudgetError', () => {
  const needsOf = (count: number): Map<string, Set<number>> => {
    const cells = new Set<number>()
    for (let index = 0; index < count; index += 1) cells.add(cellKey(index, 1))
    return new Map([['d1', cells]])
  }

  it('is silent within the budget and for formula-free batches', () => {
    expect(streamedPinBudgetError(stateWithPinned(0), new Map())).toBeNull()
    expect(streamedPinBudgetError(stateWithPinned(0), needsOf(20_000))).toBeNull()
  })

  it('rejects past the budget, counting already-pinned cells', () => {
    const error = streamedPinBudgetError(stateWithPinned(40_000), needsOf(20_000))
    expect(error).toContain('session budget')
    expect(error).toContain('aggregate_range')
    expect(error).toContain('filterColumn')
    expect(streamedPinBudgetError(stateWithPinned(40_000), needsOf(9_000))).toBeNull()
  })
})

describe('carryCopyFormulasPlan', () => {
  const state = {
    file: {
      sheets: [{ id: 'd1', name: 'Data', rowCount: 20_000, columnCount: 26, pivotRanges: [] }],
      visuals: [],
    },
    closure: { status: 'unavailable', pinned: new Map() },
  } as unknown as LazyWorkbookState
  const workbook = {
    getSheets: () => [{ getSheetId: () => 'd1', getSheetName: () => 'Data' }],
  }
  it('carries small sets, pinning the offset references', () => {
    // =SUM(A2:A100) offsets to =SUM(H2:H100). In-block references pin like
    // any other precedent: copied cells are plain journal cells, and
    // eviction would otherwise wipe the ones outside the current window.
    const plan = carryCopyFormulasPlan(
      state,
      workbook,
      'd1',
      'Data',
      new Map([['1:0', '=SUM(A2:A100)']]),
      0,
      7,
    )
    expect(plan.ok).toBe(true)
    if (plan.ok) {
      expect(plan.needs.get('d1')?.size).toBe(99)
      expect(plan.needs.get('d1')?.has(cellKey(1, 7))).toBe(true)
      expect(plan.needs.get('d1')?.has(cellKey(99, 7))).toBe(true)
    }
  })

  it('skips already-pinned cells', () => {
    const pinnedState = {
      ...state,
      closure: { status: 'unavailable', pinned: new Map([['d1', new Map([['1:7', { v: 1 }]])]]) },
    } as unknown as LazyWorkbookState
    const plan = carryCopyFormulasPlan(
      pinnedState,
      workbook,
      'd1',
      'Data',
      new Map([['1:0', '=SUM(A2:A100)']]),
      0,
      7,
    )
    expect(plan.ok).toBe(true)
    if (plan.ok) expect(plan.needs.get('d1')?.size).toBe(98)
  })

  it('refuses when the offset references exceed the session pin budget', () => {
    const plan = carryCopyFormulasPlan(
      state,
      workbook,
      'd1',
      'Data',
      new Map([['0:0', '=SUM($A$1:$Z$20000)']]),
      0,
      7,
    )
    expect(plan.ok).toBe(false)
    if (!plan.ok) expect(plan.reason).toContain('session budget')
  })

  it('refuses quadratic criteria formulas', () => {
    const plan = carryCopyFormulasPlan(
      state,
      workbook,
      'd1',
      'Data',
      new Map([['0:0', '=SUMPRODUCT(1/COUNTIF($A$2:$A$20000,$A$2:$A$20000))']]),
      0,
      7,
    )
    expect(plan.ok).toBe(false)
    if (!plan.ok) expect(plan.reason).toContain('too expensive')
  })
})

describe('proposeOperations: streaming-mode budget rejection (lazy workbook)', () => {
  function streamedState(): LazyWorkbookState {
    return {
      file: {
        sessionId: 'session-1',
        sheets: [{ id: 'd1', name: 'Data', rowCount: 20_000, columnCount: 6, pivotRanges: [] }],
        visuals: [],
      },
      editJournal: {
        cells: new Map(),
        structuralOps: new Map(),
        sheets: { added: new Set(['new1']), removed: new Set() },
        visualAdds: [],
        tableAdds: [],
      },
      loadedRanges: new Map(),
      formulaMode: false,
      flags: { preloadComplete: false },
      filterOrigins: new Map(),
      appliedDvSheets: new Set(),
      closure: { status: 'unavailable', pinned: new Map() },
    } as unknown as LazyWorkbookState
  }

  function lazyContext(state: LazyWorkbookState): PlanContext {
    const worksheets = new Map(
      [
        { id: 'd1', name: 'Data', rows: 20_000, columns: 6 },
        { id: 'new1', name: 'ja', rows: 7000, columns: 6 },
      ].map((sheet) => [
        sheet.id,
        {
          getSheetId: () => sheet.id,
          getSheetName: () => sheet.name,
          getMaxRows: () => sheet.rows,
          getMaxColumns: () => sheet.columns,
          getRange: () => ({ getValue: () => null, getRawValue: () => null }),
        },
      ]),
    )
    const workbook = {
      getActiveSheet: () => worksheets.get('d1'),
      getSheetBySheetId: (id: string) => worksheets.get(id) ?? null,
      getSheets: () => [...worksheets.values()],
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
    return proposeOperations(lazyContext(streamedState()), [operation], 'test')
  }

  it('rejects set_formula whose refs exceed the pin budget', () => {
    const outcome = propose({
      op: 'set_formula',
      sheetId: 'new1',
      address: 'A2',
      formula: '=FILTER(Data!A2:F20000,Data!D2:D20000="ja")',
    })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) {
      expect(outcome.error).toContain('session budget')
      expect(outcome.error).toContain('aggregate_range')
    }
  })

  it('rejects "="-strings smuggled in through set_range past the budget', () => {
    const outcome = propose({
      op: 'set_range',
      sheetId: 'new1',
      start: 'A2',
      values: [['=SUM(Data!A1:F20000)']],
    })
    expect(outcome.ok).toBe(false)
    if (!outcome.ok) expect(outcome.error).toContain('session budget')
  })
})
