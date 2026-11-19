import type { IRange } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  analyzeCfFormulaFold,
  buildRegisteredRanges,
  countCells,
  FOLD_MIN_COVERED_CELLS,
  notifyCfStreamWindow,
  REGISTER_HARD_CAP_CELLS,
  remapFoldedOffset,
  resetCfStreamWindows,
  WINDOW_MARGIN_ROWS,
  wrapCfFormulaService,
  type CfFormulaServiceLike,
} from '../src/renderer/cf-formula-fold'

const range = (startRow: number, endRow: number, startColumn: number, endColumn: number): IRange =>
  ({ startRow, endRow, startColumn, endColumn }) as IRange

describe('analyzeCfFormulaFold', () => {
  it('folds columns for absolute-column relative-row predicates', () => {
    expect(analyzeCfFormulaFold('=$L2="done"')).toEqual({ foldColumns: true, foldRows: false })
    expect(analyzeCfFormulaFold('=AND($A2>0,$B2<5)')).toEqual({
      foldColumns: true,
      foldRows: false,
    })
    expect(analyzeCfFormulaFold('=$A1<>""')).toEqual({ foldColumns: true, foldRows: false })
  })

  it('folds rows for absolute-row predicates and both axes for fully absolute ones', () => {
    expect(analyzeCfFormulaFold('=B$1="x"')).toEqual({ foldColumns: false, foldRows: true })
    expect(analyzeCfFormulaFold('=$B$1="x"')).toEqual({ foldColumns: true, foldRows: true })
  })

  it('keeps relative-column references per-cell', () => {
    expect(analyzeCfFormulaFold('=A2>0')).toEqual({ foldColumns: false, foldRows: false })
    expect(analyzeCfFormulaFold('=MOD(B2,2)=0')).toEqual({ foldColumns: false, foldRows: false })
  })

  it('treats quoted strings and sheet prefixes as opaque', () => {
    // The "A1" inside the string literal must not count as a reference.
    expect(analyzeCfFormulaFold('=$L2="A1"')).toEqual({ foldColumns: true, foldRows: false })
    expect(analyzeCfFormulaFold("='My Sheet'!$A2=1")).toEqual({
      foldColumns: true,
      foldRows: false,
    })
    // Sheet name that itself looks like a cell reference.
    expect(analyzeCfFormulaFold('=Q1!$A2=1')).toEqual({ foldColumns: true, foldRows: false })
    expect(analyzeCfFormulaFold('=Q1!A2=1')).toEqual({ foldColumns: false, foldRows: false })
  })

  it('blocks folding on position-dependent or opaque constructs', () => {
    expect(analyzeCfFormulaFold('=COLUMN()=2')).toEqual({ foldColumns: false, foldRows: true })
    expect(analyzeCfFormulaFold('=ROW()=$A$1')).toEqual({ foldColumns: true, foldRows: false })
    expect(analyzeCfFormulaFold('=INDIRECT("A"&ROW())=1')).toEqual({
      foldColumns: false,
      foldRows: false,
    })
    expect(analyzeCfFormulaFold('=CELL("col")=2')).toEqual({ foldColumns: false, foldRows: false })
    expect(analyzeCfFormulaFold('=RAND()>0.5')).toEqual({ foldColumns: false, foldRows: false })
    // Defined names may hide relative references.
    expect(analyzeCfFormulaFold('=$A2>Threshold')).toEqual({
      foldColumns: false,
      foldRows: false,
    })
    // Structured table references.
    expect(analyzeCfFormulaFold('=[@Amount]>0')).toEqual({ foldColumns: false, foldRows: false })
  })

  it('does not confuse COLUMNS/ROWS aggregate functions with COLUMN/ROW', () => {
    expect(analyzeCfFormulaFold('=COLUMNS($A:$B)=2')).toEqual({
      foldColumns: true,
      foldRows: true,
    })
    expect(analyzeCfFormulaFold('=ROWS($1:$2)=2')).toEqual({ foldColumns: true, foldRows: true })
  })

  it('handles span references', () => {
    expect(analyzeCfFormulaFold('=SUM($A:$A)>0')).toEqual({ foldColumns: true, foldRows: true })
    expect(analyzeCfFormulaFold('=SUM(A:A)>0')).toEqual({ foldColumns: false, foldRows: true })
    expect(analyzeCfFormulaFold('=SUM($1:$1)>0')).toEqual({ foldColumns: true, foldRows: true })
    expect(analyzeCfFormulaFold('=SUM(1:1)>0')).toEqual({ foldColumns: true, foldRows: false })
  })

  it('ignores TRUE/FALSE literals and numeric forms', () => {
    expect(analyzeCfFormulaFold('=IF($A2>1.5E2,TRUE,FALSE)')).toEqual({
      foldColumns: true,
      foldRows: false,
    })
  })
})

describe('buildRegisteredRanges', () => {
  const fold = { foldColumns: true, foldRows: false }

  it('collapses folded axes to the range edge', () => {
    const out = buildRegisteredRanges([range(0, 99_999, 0, 2)], fold, undefined)
    expect(out).toEqual([range(0, 99_999, 0, 0)])
  })

  it('clamps large registrations to the stream window and keeps the anchor cell', () => {
    const out = buildRegisteredRanges([range(0, 99_999, 0, 2)], fold, {
      startRow: 50_000,
      endRow: 50_200,
    })
    // Anchor 1x1 first so the engine's offset origin stays at the rule's
    // top-left, then the windowed rows.
    expect(out[0]).toEqual(range(0, 0, 0, 0))
    expect(out[1]).toEqual(range(50_000 - WINDOW_MARGIN_ROWS, 50_200 + WINDOW_MARGIN_ROWS, 0, 0))
  })

  it('does not duplicate the anchor when the window covers it', () => {
    const out = buildRegisteredRanges([range(0, 99_999, 0, 2)], fold, {
      startRow: 0,
      endRow: 200,
    })
    expect(out).toEqual([range(0, 200 + WINDOW_MARGIN_ROWS, 0, 0)])
  })

  it('caps unfoldable unwindowed registrations', () => {
    const out = buildRegisteredRanges(
      [range(0, 999_999, 0, 9)],
      { foldColumns: false, foldRows: false },
      undefined,
    )
    expect(countCells(out)).toBeLessThanOrEqual(REGISTER_HARD_CAP_CELLS)
    expect(out[0]?.startRow).toBe(0)
  })
})

describe('remapFoldedOffset', () => {
  it('maps every column of a row onto the registered first column', () => {
    const entry = {
      sortedRanges: [range(0, 99_999, 0, 2)],
      fold: { foldColumns: true, foldRows: false },
    }
    expect(remapFoldedOffset(entry, 5, 0)).toEqual({ row: 5, col: 0 })
    expect(remapFoldedOffset(entry, 5, 2)).toEqual({ row: 5, col: 0 })
    expect(remapFoldedOffset(entry, 0, 1)).toEqual({ row: 0, col: 0 })
  })

  it('maps secondary ranges onto their own first column', () => {
    const entry = {
      sortedRanges: [range(0, 9, 1, 3), range(20, 29, 5, 8)],
      fold: { foldColumns: true, foldRows: false },
    }
    // Cell (22, 7): inside the second range, folded to column 5; offsets are
    // relative to the first range's top-left (0-based row 0, column 1).
    expect(remapFoldedOffset(entry, 22, 6)).toEqual({ row: 22, col: 4 })
  })

  it('passes through offsets outside every range', () => {
    const entry = {
      sortedRanges: [range(0, 9, 0, 0)],
      fold: { foldColumns: true, foldRows: false },
    }
    expect(remapFoldedOffset(entry, 50, 4)).toEqual({ row: 50, col: 4 })
  })
})

interface RegisterCall {
  ranges: IRange[] | undefined
}

function mockService(): {
  service: CfFormulaServiceLike
  registers: RegisterCall[]
  deletedEngineFormulas: string[][]
} {
  const registers: RegisterCall[] = []
  const deletedEngineFormulas: string[][] = []
  const registeredIds = new Map<string, { formulaId: string }>()
  const service = {
    registerFormulaWithRange(
      unitId: string,
      _subUnitId: string,
      cfId: string,
      formulaText: string,
      ranges?: IRange[],
    ) {
      registers.push({ ranges })
      registeredIds.set(`${cfId}_${formulaText}`, { formulaId: `f${registers.length}` })
    },
    getFormulaResultWithCoords(
      _unitId: string,
      _subUnitId: string,
      _cfId: string,
      _formulaText: string,
      row = 0,
      col = 0,
    ) {
      return { status: 'success', result: `${row}:${col}` }
    },
    getFormulaMatrix() {
      return {
        status: 'success',
        result: { getValue: (row: number, col: number) => `${row}:${col}` },
      }
    },
    deleteCache(_unitId: string, _subUnitId: string, cfId: string, formulaText?: string) {
      if (formulaText !== undefined) registeredIds.delete(`${cfId}_${formulaText}`)
      return []
    },
    createCFormulaId: (cfId: string, formulaText: string) => `${cfId}_${formulaText}`,
    getSubUnitFormulaMap: () => ({
      getValue: (key: string) => registeredIds.get(key),
    }),
    _registerOtherFormulaService: {
      deleteFormula(_unitId: string, _subUnitId: string, formulaIds: string[]) {
        deletedEngineFormulas.push(formulaIds)
      },
    },
  }
  return { service: service as unknown as CfFormulaServiceLike, registers, deletedEngineFormulas }
}

describe('wrapCfFormulaService', () => {
  it('registers small rules untouched', () => {
    resetCfStreamWindows()
    const { service, registers } = mockService()
    const wrapped = wrapCfFormulaService(service)
    const small = [range(0, 9, 0, 9)]
    service.registerFormulaWithRange('u', 's', 'cf1', '=$A1=1', small)
    expect(countCells(small)).toBeLessThanOrEqual(FOLD_MIN_COVERED_CELLS)
    expect(registers[0]?.ranges).toBe(small)
    wrapped.dispose()
  })

  it('folds wide rules and remaps result lookups', () => {
    resetCfStreamWindows()
    const { service, registers } = mockService()
    const wrapped = wrapCfFormulaService(service)
    service.registerFormulaWithRange('u', 's', 'cf1', '=$A1="x"', [range(0, 9_999, 0, 9)])
    expect(registers[0]?.ranges).toEqual([range(0, 9_999, 0, 0)])
    // Column 7 of row 3 reads the registered column-0 result.
    expect(service.getFormulaResultWithCoords('u', 's', 'cf1', '=$A1="x"', 3, 7)).toEqual({
      status: 'success',
      result: '3:0',
    })
    const matrix = service.getFormulaMatrix('u', 's', 'cf1', '=$A1="x"')
    expect(matrix.result?.getValue(4, 9)).toBe('4:0')
    wrapped.dispose()
  })

  it('remaps a service-cached matrix without mutating it or nesting layers', () => {
    resetCfStreamWindows()
    const { service } = mockService()
    // The mock below returns the SAME matrix instance on every call — the
    // cache-reuse case: repeated wrapping must not stack remaps, and the
    // shared instance must keep its raw getValue.
    const shared = { getValue: (row: number, col: number) => `${row}:${col}` }
    ;(service as { getFormulaMatrix: unknown }).getFormulaMatrix = () => ({
      status: 'success',
      result: shared,
    })
    const wrapped = wrapCfFormulaService(service)
    service.registerFormulaWithRange('u', 's', 'cf1', '=$A1="x"', [range(0, 9_999, 0, 9)])
    const first = service.getFormulaMatrix('u', 's', 'cf1', '=$A1="x"')
    const second = service.getFormulaMatrix('u', 's', 'cf1', '=$A1="x"')
    expect(first.result?.getValue(4, 9)).toBe('4:0')
    expect(second.result?.getValue(4, 9)).toBe('4:0')
    expect(shared.getValue(4, 9)).toBe('4:9')
    wrapped.dispose()
  })

  it('ignores duplicate registrations like the original dedupe', () => {
    resetCfStreamWindows()
    const { service, registers } = mockService()
    const wrapped = wrapCfFormulaService(service)
    const ranges = [range(0, 9_999, 0, 9)]
    service.registerFormulaWithRange('u', 's', 'cf1', '=$A1="x"', ranges)
    service.registerFormulaWithRange('u', 's', 'cf1', '=$A1="x"', ranges)
    expect(registers).toHaveLength(1)
    wrapped.dispose()
  })

  it('windows huge unfoldable rules and re-registers when the stream window moves', async () => {
    resetCfStreamWindows()
    const { service, registers, deletedEngineFormulas } = mockService()
    const wrapped = wrapCfFormulaService(service)
    notifyCfStreamWindow('s', 0, 200)
    service.registerFormulaWithRange('u', 's', 'cf1', '=A1>0', [range(0, 99_999, 0, 9)])
    expect(registers).toHaveLength(1)
    const first = registers[0]?.ranges
    expect(first?.[0]).toEqual(range(0, 200 + WINDOW_MARGIN_ROWS, 0, 9))
    // Lookups inside the window pass through unchanged (no folding).
    expect(service.getFormulaResultWithCoords('u', 's', 'cf1', '=A1>0', 3, 7)).toEqual({
      status: 'success',
      result: '3:7',
    })
    // Scroll within the margin: no re-registration.
    notifyCfStreamWindow('s', 100, 400)
    await new Promise((resolve) => setTimeout(resolve, 5))
    expect(registers).toHaveLength(1)
    // Deep scroll: engine formula deleted and re-registered around the new window.
    notifyCfStreamWindow('s', 50_000, 50_200)
    await new Promise((resolve) => setTimeout(resolve, 5))
    expect(deletedEngineFormulas).toHaveLength(1)
    expect(registers).toHaveLength(2)
    const second = registers[1]?.ranges
    expect(second?.[0]).toEqual(range(0, 0, 0, 0))
    expect(second?.[1]).toEqual(
      range(50_000 - WINDOW_MARGIN_ROWS, 50_200 + WINDOW_MARGIN_ROWS, 0, 9),
    )
    wrapped.dispose()
  })

  it('drops tracking when a rule is deleted', () => {
    resetCfStreamWindows()
    const { service } = mockService()
    const wrapped = wrapCfFormulaService(service)
    service.registerFormulaWithRange('u', 's', 'cf1', '=$A1="x"', [range(0, 9_999, 0, 9)])
    service.deleteCache('u', 's', 'cf1')
    // Lookups fall through untouched once tracking is gone.
    expect(service.getFormulaResultWithCoords('u', 's', 'cf1', '=$A1="x"', 3, 7)).toEqual({
      status: 'success',
      result: '3:7',
    })
    wrapped.dispose()
  })
})
