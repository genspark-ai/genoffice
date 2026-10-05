import { describe, expect, it } from 'vitest'
import {
  offsetFormulaRefs,
  shiftFormulaRefs,
  shiftIndex,
  shiftSpecForOp,
} from '../src/domain/formula-shift'
import type { StructuralOperation } from '../src/domain/workbook-dsl'
import { applyDefinedNamesState } from '../src/gateway/xlsx-defined-names'
import { spillsDynamicArray, withFutureFunctionMarkers } from '../src/gateway/future-functions'

const SHEET = 'Sheet1'

function insertRows(row: number, count = 1): StructuralOperation {
  return { op: 'insert_rows', sheetId: 's1', row, count }
}

function deleteRows(row: number, count = 1): StructuralOperation {
  return { op: 'delete_rows', sheetId: 's1', row, count }
}

function insertCols(column: string, count = 1): StructuralOperation {
  return { op: 'insert_cols', sheetId: 's1', column, count }
}

function deleteCols(column: string, count = 1): StructuralOperation {
  return { op: 'delete_cols', sheetId: 's1', column, count }
}

describe('shiftSpecForOp', () => {
  it('maps row ops to a row spec with a 0-based start', () => {
    expect(shiftSpecForOp(insertRows(2, 3))).toEqual({ axis: 'row', start: 1, delta: 3 })
    expect(shiftSpecForOp(deleteRows(2, 3))).toEqual({ axis: 'row', start: 1, delta: -3 })
  })

  it('maps column ops to a column spec with a 0-based start', () => {
    expect(shiftSpecForOp(insertCols('B', 2))).toEqual({ axis: 'column', start: 1, delta: 2 })
    expect(shiftSpecForOp(deleteCols('C', 1))).toEqual({ axis: 'column', start: 2, delta: -1 })
  })

  it('returns null for sheet-level ops', () => {
    expect(shiftSpecForOp({ op: 'add_sheet', name: 'New' })).toBeNull()
    expect(shiftSpecForOp({ op: 'delete_sheet', sheetId: 's1' })).toBeNull()
  })
})

describe('shiftIndex', () => {
  it('leaves indexes before the insertion alone and shifts later ones', () => {
    expect(shiftIndex(0, { axis: 'row', start: 1, delta: 1 })).toBe(0)
    expect(shiftIndex(1, { axis: 'row', start: 1, delta: 2 })).toBe(3)
  })

  it('maps deleted indexes to null and shifts later ones up', () => {
    const spec = { axis: 'row', start: 1, delta: -2 } as const
    expect(shiftIndex(0, spec)).toBe(0)
    expect(shiftIndex(1, spec)).toBeNull()
    expect(shiftIndex(2, spec)).toBeNull()
    expect(shiftIndex(3, spec)).toBe(1)
  })
})

describe('shiftFormulaRefs on insert rows', () => {
  it('shifts refs at or below the insertion down', () => {
    const result = shiftFormulaRefs('=A1+A3', insertRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=A1+A4')
    expect(result.changed).toBe(true)
    expect(result.hasRefError).toBe(false)
  })

  it('leaves refs above the insertion unchanged', () => {
    const result = shiftFormulaRefs('=A1', insertRows(5, 2), true, SHEET)
    expect(result.formula).toBe('=A1')
    expect(result.changed).toBe(false)
    expect(result.hasRefError).toBe(false)
  })

  it('shifts absolute refs too because dollars do not pin structural edits', () => {
    const result = shiftFormulaRefs('=$A$1+$B$2', insertRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=$A$1+$B$3')
    expect(result.changed).toBe(true)
  })

  it('shifts multi-row inserts by the full count', () => {
    const result = shiftFormulaRefs('=A2', insertRows(2, 2), true, SHEET)
    expect(result.formula).toBe('=A4')
  })

  it('shifts ranges endpoint by endpoint', () => {
    const result = shiftFormulaRefs('=SUM($A$1:$B$2)', insertRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=SUM($A$1:$B$3)')
  })
})

describe('shiftFormulaRefs on delete rows', () => {
  it('turns refs inside the deleted region into #REF!', () => {
    const result = shiftFormulaRefs('=A1+A2+A3', deleteRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=A1+#REF!+A2')
    expect(result.changed).toBe(true)
    expect(result.hasRefError).toBe(true)
  })

  it('shifts refs below the deletion up', () => {
    const result = shiftFormulaRefs('=A5', deleteRows(2, 2), true, SHEET)
    expect(result.formula).toBe('=A3')
    expect(result.hasRefError).toBe(false)
  })

  it('shrinks a partially deleted range instead of erroring', () => {
    const result = shiftFormulaRefs('=SUM(A1:A3)', deleteRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=SUM(A1:A2)')
    expect(result.hasRefError).toBe(false)
  })

  it('turns a fully deleted range into #REF!', () => {
    const result = shiftFormulaRefs('=SUM(A2:A3)', deleteRows(2, 2), true, SHEET)
    expect(result.formula).toBe('=SUM(#REF!)')
    expect(result.hasRefError).toBe(true)
  })

  it('does not spare absolute refs inside the deleted region', () => {
    const result = shiftFormulaRefs('=$A$2', deleteRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=#REF!')
    expect(result.hasRefError).toBe(true)
  })
})

describe('shiftFormulaRefs on columns', () => {
  it('shifts lowercase cell refs and whole-column spans on insert', () => {
    const result = shiftFormulaRefs('=sum(a1:b2)+sum(b:d)', insertCols('B', 1), true, SHEET)
    expect(result.formula).toBe('=sum(A1:C2)+sum(C:E)')
    expect(result.changed).toBe(true)
  })

  it('shrinks lowercase ranges when a column is deleted', () => {
    const result = shiftFormulaRefs('=sum(a1:b2)', deleteCols('B', 1), true, SHEET)
    expect(result.formula).toBe('=sum(A1:A2)')
    expect(result.hasRefError).toBe(false)
  })

  it('shifts refs right on insert', () => {
    const result = shiftFormulaRefs('=A1+B1+C1', insertCols('B', 1), true, SHEET)
    expect(result.formula).toBe('=A1+C1+D1')
  })

  it('shifts absolute column refs right on insert', () => {
    const result = shiftFormulaRefs('=$A$1+$B$1', insertCols('B', 1), true, SHEET)
    expect(result.formula).toBe('=$A$1+$C$1')
  })

  it('turns deleted column refs into #REF! and shifts later ones left', () => {
    const result = shiftFormulaRefs('=A1+B1+C1', deleteCols('B', 1), true, SHEET)
    expect(result.formula).toBe('=A1+#REF!+B1')
    expect(result.hasRefError).toBe(true)
  })

  it('shrinks a partially deleted column range', () => {
    const result = shiftFormulaRefs('=SUM(A1:C1)', deleteCols('B', 1), true, SHEET)
    expect(result.formula).toBe('=SUM(A1:B1)')
    expect(result.hasRefError).toBe(false)
  })

  it('shifts whole-column spans on insert', () => {
    const result = shiftFormulaRefs('=SUM(B:B)', insertCols('B', 1), true, SHEET)
    expect(result.formula).toBe('=SUM(C:C)')
  })

  it('shifts whole-row spans on insert', () => {
    const result = shiftFormulaRefs('=SUM(2:4)', insertRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=SUM(3:5)')
  })
})

describe('shiftFormulaRefs at the grid edge', () => {
  it('turns a ref the shift pushes past the last row or column into #REF!', () => {
    const lastRow = shiftFormulaRefs('=A1048576', insertRows(2, 5), true, SHEET)
    expect(lastRow.formula).toBe('=#REF!')
    expect(lastRow.hasRefError).toBe(true)

    const lastColumn = shiftFormulaRefs('=XFD1', insertCols('B', 1), true, SHEET)
    expect(lastColumn.formula).toBe('=#REF!')

    // Spans take the same bound, and a range that only overruns at its far end
    // is #REF! too — never shrunk into an inverted range.
    expect(shiftFormulaRefs('=SUM(XFD:XFD)', insertCols('B', 1), true, SHEET).formula).toBe(
      '=SUM(#REF!)',
    )
    expect(shiftFormulaRefs('=SUM(2:1048576)', insertRows(2, 1), true, SHEET).formula).toBe(
      '=SUM(#REF!)',
    )
    expect(shiftFormulaRefs('=SUM(A2:B1048576)', insertRows(2, 1), true, SHEET).formula).toBe(
      '=SUM(#REF!)',
    )
  })
})

describe('offsetFormulaRefs', () => {
  it('shifts lowercase cell refs and whole-column spans on fill', () => {
    expect(offsetFormulaRefs('=sum(a1:b2)+sum($b:d)', 0, 1)).toBe('=sum(B1:C2)+sum($B:E)')
  })
})

describe('shiftFormulaRefs cross-sheet', () => {
  it('leaves bare refs alone when the formula lives on another sheet', () => {
    const result = shiftFormulaRefs('=A3', insertRows(2, 1), false, SHEET)
    expect(result.formula).toBe('=A3')
    expect(result.changed).toBe(false)
  })

  it('rewrites a matching sheet prefix even when the formula lives elsewhere', () => {
    const result = shiftFormulaRefs('=Sheet1!A3', insertRows(2, 1), false, SHEET)
    expect(result.formula).toBe('=Sheet1!A4')
    expect(result.changed).toBe(true)
  })

  it('leaves a non-matching sheet prefix alone', () => {
    const result = shiftFormulaRefs('=Other!A3', insertRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=Other!A3')
    expect(result.changed).toBe(false)
  })

  it('rewrites quoted sheet names with spaces', () => {
    const result = shiftFormulaRefs("='My Sheet'!A3", insertRows(2, 1), false, 'My Sheet')
    expect(result.formula).toBe("='My Sheet'!A4")
  })

  it('rewrites quoted names with an escaped single quote', () => {
    const result = shiftFormulaRefs("='Bob''s'!A3", insertRows(2, 1), false, "Bob's")
    expect(result.formula).toBe("='Bob''s'!A4")
  })

  it('rewrites sheet-qualified ranges on the target sheet only', () => {
    const kept = shiftFormulaRefs('=SUM(Other!A1:A3)', deleteRows(2, 1), true, SHEET)
    expect(kept.formula).toBe('=SUM(Other!A1:A3)')
    const moved = shiftFormulaRefs('=SUM(Sheet1!A1:A3)', deleteRows(2, 1), false, SHEET)
    expect(moved.formula).toBe('=SUM(Sheet1!A1:A2)')
  })
})

describe('shiftFormulaRefs literals and names', () => {
  it('skips refs inside quoted string literals', () => {
    const result = shiftFormulaRefs('="A3"+A3', insertRows(2, 1), true, SHEET)
    expect(result.formula).toBe('="A3"+A4')
  })

  it('leaves function names such as LOG10 alone', () => {
    const result = shiftFormulaRefs('=LOG10(A1)+A3', insertRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=LOG10(A1)+A4')
  })

  it('returns the input unchanged for sheet-level ops', () => {
    const result = shiftFormulaRefs('=A1', { op: 'add_sheet', name: 'New' }, true, SHEET)
    expect(result).toEqual({ formula: '=A1', changed: false, hasRefError: false })
  })
})

describe('shiftFormulaRefs structured references', () => {
  // A structured reference names a COLUMN, not a cell: Table1[Q1] is the Q1 column
  // and must not move when rows are inserted, while a real ref in the same formula
  // still shifts.
  it('leaves a bracketed column name alone but still shifts the real refs', () => {
    const result = shiftFormulaRefs('=SUM(Table1[Q1],A3)', insertRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=SUM(Table1[Q1],A4)')
  })

  it('leaves the this-row form alone', () => {
    const result = shiftFormulaRefs('=SUM(Table1[@Q1],A3)', insertRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=SUM(Table1[@Q1],A4)')
  })

  it('leaves a padded column name alone', () => {
    const result = shiftFormulaRefs('=SUM(Table1[ Q1 ],A3)', insertRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=SUM(Table1[ Q1 ],A4)')
  })

  it('leaves a nested header reference alone', () => {
    const result = shiftFormulaRefs(
      '=SUM(Table1[[#Headers],[Q1]],A3)',
      insertRows(2, 1),
      true,
      SHEET,
    )
    expect(result.formula).toBe('=SUM(Table1[[#Headers],[Q1]],A4)')
  })

  it('handles several references and a column insert', () => {
    const result = shiftFormulaRefs('=SUM(Table1[Q1],B3)', insertRows(2, 1), true, SHEET)
    expect(result.formula).toBe('=SUM(Table1[Q1],B4)')

    const colInsert = shiftFormulaRefs('=SUM(Table1[Q1],B3)', insertCols('B', 1), true, SHEET)
    expect(colInsert.formula).toBe('=SUM(Table1[Q1],C3)')
  })
})

describe('future function markers', () => {
  it('marks calls after a quoted sheet qualifier', () => {
    expect(withFutureFunctionMarkers('\'Data Sheet\'!MINIFS(A1:A3,A1:A3,">0")')).toBe(
      '\'Data Sheet\'!_xlfn.MINIFS(A1:A3,A1:A3,">0")',
    )
  })

  it('leaves callable defined names containing the marker unchanged', () => {
    const formula = 'Budget_xlfn.Total(A1)+Other_xlfn.FILTER(A1)'
    expect(withFutureFunctionMarkers(formula)).toBe(formula)
    expect(spillsDynamicArray('Budget_xlfn.FILTER(A1)')).toBe(false)
  })

  it('does not treat a marked sheet name as a spill function', () => {
    expect(spillsDynamicArray("'Data_xlfn.FILTER'!A1+FILTER(A1:A2)")).toBe(true)
    expect(withFutureFunctionMarkers("'Data_xlfn.FILTER'!FILTER(A1:A2)")).toBe(
      "'Data_xlfn.FILTER'!_xlfn._xlws.FILTER(A1:A2)",
    )
  })

  it('protects marker-bearing callable names in defined-name save XML', () => {
    const workbook =
      '<workbook><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>'
    const saved = applyDefinedNamesState(workbook, {
      names: [
        { name: 'Budget_xlfn.Total', formula: 'Budget_xlfn.Total(A1)' },
        { name: 'FutureTotal', formula: `'Data_xlfn.Total'!MINIFS(A1:A3,A1:A3,">0")` },
      ],
      preserveNames: [],
    })

    expect(saved).toContain(
      '<definedName name="Budget_xlfn.Total">Budget_xlfn.Total(A1)</definedName>',
    )
    expect(saved).toContain(
      `<definedName name="FutureTotal">'Data_xlfn.Total'!_xlfn.MINIFS(A1:A3,A1:A3,"&gt;0")</definedName>`,
    )
  })
})
