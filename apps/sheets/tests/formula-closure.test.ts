import { describe, expect, it } from 'vitest'

import {
  cellKey,
  closureFetchRanges,
  computeFormulaClosure,
  containsUnresolvedNames,
  parseFormulaReferences,
  recalcReadRanges,
  shiftPinnedCells,
  type ClosureSheetInput,
} from '../src/renderer/formula-closure'
import { formulaKeepsCache } from '../src/renderer/univer-sync'

const sheet = (
  id: string,
  name: string,
  formulas: { row: number; column: number; formula: string }[],
  rowCount = 100,
  columnCount = 20,
): ClosureSheetInput => ({ id, name, rowCount, columnCount, formulas })

describe('parseFormulaReferences', () => {
  it('parses cells, ranges, absolutes, and qualified references', () => {
    const refs = parseFormulaReferences("SUM(B2:B10)+$C$1+'My Sheet'!A5+Data!D1:D3")
    expect(refs).toHaveLength(4)
    expect(refs[0]?.token).toEqual({ startRow: 1, endRow: 9, startColumn: 1, endColumn: 1 })
    expect(refs[1]?.token).toEqual({ startRow: 0, endRow: 0, startColumn: 2, endColumn: 2 })
    expect(refs[2]?.qualifier).toBe("'My Sheet'")
    expect(refs[3]?.qualifier).toBe('Data')
  })

  // The prod_040 incident shape: an Arabic sheet name, unquoted, qualifying
  // absolute references. An ASCII-only qualifier pattern dropped these
  // entirely, so the closure never loaded B30/B31 and NPV computed with a
  // zero rate.
  it('parses unquoted non-ASCII sheet qualifiers, absolute and relative', () => {
    const zakat = parseFormulaReferences('IF(B18>0,B18*الافتراضات!$B$30,0)')
    expect(zakat).toHaveLength(3)
    expect(zakat[2]?.qualifier).toBe('الافتراضات')
    expect(zakat[2]?.token).toEqual({ startRow: 29, endRow: 29, startColumn: 1, endColumn: 1 })
    const relative = parseFormulaReferences('الافتراضات!B16+集計!C2:C4')
    expect(relative.map((ref) => ref.qualifier)).toEqual(['الافتراضات', '集計'])
  })

  it('ignores references inside string literals and marks whole columns unbounded', () => {
    const refs = parseFormulaReferences('IF(A1>0,"see B2:B9",SUM(D:D))')
    expect(refs).toHaveLength(2)
    expect(refs[1]?.token).toEqual({
      startRow: null,
      endRow: null,
      startColumn: 3,
      endColumn: 3,
    })
  })
})

describe('containsUnresolvedNames', () => {
  it('accepts plain functions, booleans, and references', () => {
    expect(containsUnresolvedNames('SUM(A1:A5)+IF(TRUE,MAX(B1),0)')).toBe(false)
    expect(containsUnresolvedNames('COUNTIF(C:C,">1")')).toBe(false)
  })

  it('flags defined-name style identifiers', () => {
    expect(containsUnresolvedNames('SUM(MyRange)')).toBe(true)
    expect(containsUnresolvedNames('Total*2')).toBe(true)
  })

  it('flags non-ASCII defined names but not non-ASCII sheet qualifiers', () => {
    expect(containsUnresolvedNames('IF(B18>0,B18*الافتراضات!$B$30,0)')).toBe(false)
    expect(containsUnresolvedNames('NPV(الافتراضات!$B$31,C16:G16)+B18')).toBe(false)
    expect(containsUnresolvedNames('合計*2')).toBe(true)
  })
})

describe('formulaKeepsCache', () => {
  it('keeps the cached value for Google Sheets __xludf exports (prod_016)', () => {
    // Recalculating turns the float-repr fallback literal (46235.0) into a
    // string the numfmt layer skips; the cached <v> is the computed value.
    expect(formulaKeepsCache('IFERROR(__xludf.DUMMYFUNCTION("""COMPUTED_VALUE"""),46235.0)')).toBe(
      true,
    )
    expect(formulaKeepsCache('IFERROR(__xludf.DUMMYFUNCTION("X"),"✖")')).toBe(true)
  })

  it('recalculates ordinary formulas', () => {
    expect(formulaKeepsCache('SUM(A1:A5)')).toBe(false)
  })
})

describe('computeFormulaClosure', () => {
  it('collects formulas plus referenced precedents across sheets', () => {
    const result = computeFormulaClosure(
      [
        sheet('sheet-1', 'Data', [
          { row: 0, column: 3, formula: 'SUM(A1:A3)' },
          { row: 1, column: 3, formula: "'Other'!B2*2" },
        ]),
        sheet('sheet-2', 'Other', []),
      ],
      1000,
    )
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.formulaCount).toBe(2)
    // 2 formulas + A1:A3 + Other!B2
    expect(result.totalCells).toBe(6)
    expect(result.cellsBySheet.get('sheet-2')?.has(cellKey(1, 1))).toBe(true)
  })

  it('loads precedents referenced through unquoted non-ASCII sheet names', () => {
    const result = computeFormulaClosure(
      [
        sheet('sheet-2', 'الافتراضات', []),
        sheet('sheet-3', 'قائمة الدخل', [
          { row: 18, column: 1, formula: 'IF(B18>0,B18*الافتراضات!$B$30,0)' },
        ]),
      ],
      1000,
    )
    expect(result.ok).toBe(true)
    if (!result.ok) return
    expect(result.cellsBySheet.get('sheet-2')?.has(cellKey(29, 1))).toBe(true)
  })

  it('clamps whole-column references to the used range and bails over budget', () => {
    const small = computeFormulaClosure(
      [sheet('sheet-1', 'Data', [{ row: 0, column: 5, formula: 'SUM(A:A)' }], 50, 10)],
      1000,
    )
    expect(small.ok).toBe(true)
    if (small.ok) expect(small.totalCells).toBe(51)

    const over = computeFormulaClosure(
      [sheet('sheet-1', 'Data', [{ row: 0, column: 5, formula: 'SUM(A:C)' }], 5000, 10)],
      1000,
    )
    expect(over.ok).toBe(false)
  })

  it('fails on unknown sheets and defined names', () => {
    const unknown = computeFormulaClosure(
      [sheet('sheet-1', 'Data', [{ row: 0, column: 0, formula: 'Missing!A1' }])],
      1000,
    )
    expect(unknown.ok).toBe(false)
    const named = computeFormulaClosure(
      [sheet('sheet-1', 'Data', [{ row: 0, column: 0, formula: 'SUM(MyRange)' }])],
      1000,
    )
    expect(named.ok).toBe(false)
  })
})

describe('closureFetchRanges', () => {
  it('bands nearby rows and splits when over the read budget', () => {
    const cells = new Set<number>()
    for (let row = 0; row < 10; row += 1) cells.add(cellKey(row, 2))
    cells.add(cellKey(500, 0))
    const ranges = closureFetchRanges(cells)
    expect(ranges).toEqual([
      { startRow: 0, endRow: 9, startColumn: 2, endColumn: 2 },
      { startRow: 500, endRow: 500, startColumn: 0, endColumn: 0 },
    ])

    const wide = new Set<number>()
    for (let row = 0; row < 40; row += 1) {
      wide.add(cellKey(row, 0))
      wide.add(cellKey(row, 999))
    }
    for (const range of closureFetchRanges(wide, 10_000)) {
      const area = (range.endRow - range.startRow + 1) * (range.endColumn - range.startColumn + 1)
      expect(area).toBeLessThanOrEqual(10_000)
    }
  })
})

describe('shiftPinnedCells', () => {
  it('moves keys through inserts and drops removed positions', () => {
    const pinned = new Map([
      ['5:2', 'a'],
      ['10:2', 'b'],
      ['3:0', 'c'],
    ])
    const inserted = shiftPinnedCells(pinned, { kind: 'insert-rows', index: 4, count: 2 })
    expect([...inserted.keys()].sort()).toEqual(['12:2', '3:0', '7:2'])
    const removed = shiftPinnedCells(pinned, { kind: 'remove-rows', index: 5, count: 1 })
    expect([...removed.keys()].sort()).toEqual(['3:0', '9:2'])
    const columns = shiftPinnedCells(pinned, { kind: 'insert-cols', index: 1, count: 3 })
    expect([...columns.keys()].sort()).toEqual(['10:5', '3:0', '5:5'])
  })
})

describe('recalcReadRanges', () => {
  it('keeps bands within the budget, nearest the viewport first', () => {
    const keys = new Set<number>()
    // Two distant bands: rows 0-1 (2 cells) and rows 5000-5999 (1000 cells).
    keys.add(cellKey(0, 0))
    keys.add(cellKey(1, 0))
    for (let row = 5000; row < 6000; row += 1) keys.add(cellKey(row, 0))
    const nearTop = recalcReadRanges(keys, 0, 500)
    expect(nearTop).toEqual([{ startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 }])
    const nearBottom = recalcReadRanges(keys, 5500, 1200)
    expect(nearBottom[0]?.startRow).toBe(5000)
    expect(nearBottom).toHaveLength(2)
  })

  it('returns everything when the budget covers all bands', () => {
    const keys = new Set([cellKey(2, 1), cellKey(3, 2)])
    expect(recalcReadRanges(keys, 0, 20_000)).toEqual([
      { startRow: 2, endRow: 3, startColumn: 1, endColumn: 2 },
    ])
  })
})
