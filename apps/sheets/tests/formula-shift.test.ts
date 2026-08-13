import { describe, expect, it } from 'vitest'

import { shiftFormulaRefs } from '../src/domain/formula-shift'
import type { StructuralOperation } from '../src/domain/workbook-dsl'

const insertRows = (row: number, count = 1): StructuralOperation => ({
  op: 'insert_rows',
  sheetId: 's1',
  row,
  count,
})
const deleteRows = (row: number, count = 1): StructuralOperation => ({
  op: 'delete_rows',
  sheetId: 's1',
  row,
  count,
})
const insertCols = (column: string, count = 1): StructuralOperation => ({
  op: 'insert_cols',
  sheetId: 's1',
  column,
  count,
})
const deleteCols = (column: string, count = 1): StructuralOperation => ({
  op: 'delete_cols',
  sheetId: 's1',
  column,
  count,
})

function shift(formula: string, op: StructuralOperation, sameSheet = true, opSheetName = 'Sheet1') {
  return shiftFormulaRefs(formula, op, sameSheet, opSheetName)
}

describe('shiftFormulaRefs: row inserts', () => {
  it('shifts refs at or below the insertion point', () => {
    expect(shift('=SUM(B2:B10)+A1', insertRows(5, 2)).formula).toBe('=SUM(B2:B12)+A1')
  })

  it('shifts absolute refs too (structural edits ignore $)', () => {
    expect(shift('=$B$5+B$5+$B5', insertRows(3)).formula).toBe('=$B$6+B$6+$B6')
  })

  it('leaves refs above the insertion point alone', () => {
    const result = shift('=A1+A2', insertRows(5))
    expect(result.formula).toBe('=A1+A2')
    expect(result.changed).toBe(false)
  })
})

describe('shiftFormulaRefs: row deletes', () => {
  it('shifts refs below the deleted rows up', () => {
    expect(shift('=A10', deleteRows(2, 3)).formula).toBe('=A7')
  })

  it('turns a ref into #REF! when its row is deleted', () => {
    const result = shift('=A3*2', deleteRows(3))
    expect(result.formula).toBe('=#REF!*2')
    expect(result.hasRefError).toBe(true)
  })

  it('shrinks a range whose end is deleted', () => {
    expect(shift('=SUM(B2:B4)', deleteRows(3, 3)).formula).toBe('=SUM(B2:B2)')
  })

  it('shrinks a range whose start is deleted', () => {
    expect(shift('=SUM(B3:B10)', deleteRows(3, 2)).formula).toBe('=SUM(B3:B8)')
  })

  it('turns a fully deleted range into #REF!', () => {
    const result = shift('=SUM(B3:B4)', deleteRows(3, 2))
    expect(result.formula).toBe('=SUM(#REF!)')
    expect(result.hasRefError).toBe(true)
  })
})

describe('shiftFormulaRefs: column shifts', () => {
  it('shifts columns on insert and delete', () => {
    expect(shift('=SUM(C1:E1)', insertCols('B', 2)).formula).toBe('=SUM(E1:G1)')
    expect(shift('=D1', deleteCols('B')).formula).toBe('=C1')
    expect(shift('=B1', deleteCols('B')).hasRefError).toBe(true)
  })
})

describe('shiftFormulaRefs: what must NOT be rewritten', () => {
  it('does not mangle function names containing digits', () => {
    const result = shift('=LOG10(A1)', insertRows(5))
    expect(result.formula).toBe('=LOG10(A1)')
  })

  it('skips refs inside string literals', () => {
    expect(shift('="see B5"&B5', insertRows(3)).formula).toBe('="see B5"&B6')
  })

  it('ignores bare refs when the formula lives on a different sheet', () => {
    expect(shift('=B5', insertRows(3), false).changed).toBe(false)
  })
})

describe('shiftFormulaRefs: sheet prefixes', () => {
  it('rewrites prefixed refs that target the op sheet', () => {
    expect(shift("=Sheet1!B5+'Sheet1'!C5", insertRows(3), false).formula).toBe(
      "=Sheet1!B6+'Sheet1'!C6",
    )
  })

  it('leaves prefixed refs to other sheets alone', () => {
    expect(shift('=Other!B5', insertRows(3), true).formula).toBe('=Other!B5')
  })

  it('handles quoted sheet names with spaces', () => {
    expect(shift("='My Data'!B5", insertRows(3), false, 'My Data').formula).toBe("='My Data'!B6")
  })
})

describe('shiftFormulaRefs: add_sheet is a no-op', () => {
  it('never changes formulas', () => {
    const result = shift('=A1', { op: 'add_sheet', name: 'New' })
    expect(result.changed).toBe(false)
  })
})
