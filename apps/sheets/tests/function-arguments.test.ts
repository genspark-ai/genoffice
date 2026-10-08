import { describe, expect, it } from 'vitest'

import {
  buildFunctionCall,
  findFunctionCallAtCaret,
  finishedFormula,
  formatPreviewValue,
  parameterRows,
  parseSyntaxParams,
  pickedRangeRef,
  quoteSheetName,
  spliceFunctionCall,
  splitTopLevelArgs,
} from '../src/renderer/function-arguments'

describe('splitTopLevelArgs', () => {
  it('splits at top-level commas only', () => {
    expect(splitTopLevelArgs('A1:A3, SUM(B1, B2), "a,b", {1,2;3,4}, \'My, Sheet\'!C1')).toEqual([
      'A1:A3',
      'SUM(B1, B2)',
      '"a,b"',
      '{1,2;3,4}',
      "'My, Sheet'!C1",
    ])
  })

  it('keeps empty positional slots and treats blank text as no arguments', () => {
    expect(splitTopLevelArgs('A1,,3')).toEqual(['A1', '', '3'])
    expect(splitTopLevelArgs('')).toEqual([])
    expect(splitTopLevelArgs('  ')).toEqual([])
  })

  it('honours doubled quotes inside strings', () => {
    expect(splitTopLevelArgs('"say ""hi"", ok", 2')).toEqual(['"say ""hi"", ok"', '2'])
  })
})

describe('findFunctionCallAtCaret', () => {
  const formula = '=IF(SUM(A1:A3)>10, "big", LEN("x)y"))'

  it('finds the innermost call around the caret', () => {
    const caret = formula.indexOf('A1') + 1
    expect(findFunctionCallAtCaret(formula, caret)).toMatchObject({
      name: 'SUM',
      start: 4,
      end: formula.indexOf(')') + 1,
      args: ['A1:A3'],
    })
  })

  it('attributes a caret right after a closing paren to that call', () => {
    expect(findFunctionCallAtCaret('=SUM(A1:A3)', 11)?.name).toBe('SUM')
    expect(findFunctionCallAtCaret(formula, formula.length)?.name).toBe('IF')
  })

  it('ignores parens inside string literals', () => {
    const caret = formula.indexOf('"x)y"') + 2
    expect(findFunctionCallAtCaret(formula, caret)?.name).toBe('LEN')
  })

  it('handles an unclosed call while typing', () => {
    expect(findFunctionCallAtCaret('=SUM(A1:A3', 10)).toMatchObject({
      name: 'SUM',
      start: 1,
      end: 10,
      args: ['A1:A3'],
    })
  })

  it('returns null for a bare expression or caret outside every call', () => {
    expect(findFunctionCallAtCaret('=A1+A2', 3)).toBeNull()
    expect(findFunctionCallAtCaret('=1+SUM(A1)', 1)).toBeNull()
    expect(findFunctionCallAtCaret('=(1+2)*3', 3)).toBeNull()
  })
})

describe('buildFunctionCall / spliceFunctionCall / finishedFormula', () => {
  it('drops trailing empties but keeps inner positional gaps', () => {
    expect(buildFunctionCall('IF', ['A1>1', '', '"no"', '', ''])).toBe('IF(A1>1, , "no")')
    expect(buildFunctionCall('TODAY', [])).toBe('TODAY()')
  })

  it('replaces exactly the call under the caret', () => {
    const formula = '=1+SUM(A1:A3)*2'
    const span = findFunctionCallAtCaret(formula, 8)!
    expect(spliceFunctionCall(formula, span, 'SUM(B1:B9, C1)')).toBe('=1+SUM(B1:B9, C1)*2')
  })

  it('becomes the whole formula of a plain cell, else lands at the caret', () => {
    expect(finishedFormula({ formula: '', caret: 0 }, null, 'SUM', ['A1'])).toBe('=SUM(A1)')
    expect(finishedFormula({ formula: '=1+', caret: 3 }, null, 'SUM', ['A1'])).toBe('=1+SUM(A1)')
    const formula = '=SUM(A1:A3)'
    const span = findFunctionCallAtCaret(formula, formula.length)
    expect(finishedFormula({ formula, caret: formula.length }, span, 'SUM', ['A1:A5'])).toBe(
      '=SUM(A1:A5)',
    )
  })
})

describe('parseSyntaxParams', () => {
  it('reads required, optional and repeating parameters from the syntax line', () => {
    expect(parseSyntaxParams('SUM(number1, [number2], …)')).toEqual([
      { name: 'number1', detail: '', require: true, repeat: false },
      { name: 'number2', detail: '', require: false, repeat: true },
    ])
    expect(
      parseSyntaxParams('RATE(nper, pmt, pv, [fv], [type], [guess])').map((p) => p.require),
    ).toEqual([true, true, true, false, false, false])
    expect(parseSyntaxParams('TODAY()')).toEqual([])
  })
})

describe('parameterRows', () => {
  const sum = parseSyntaxParams('SUM(number1, [number2], …)')

  it('shows one spare repeating row and grows as arguments fill', () => {
    expect(parameterRows(sum, 0).map((row) => row.name)).toEqual(['number1', 'number2'])
    expect(parameterRows(sum, 2).map((row) => row.name)).toEqual(['number1', 'number2', 'number3'])
    expect(parameterRows(sum, 2).map((row) => row.index)).toEqual([0, 1, 2])
  })

  it('bolds only required parameters, never a repeat continuation', () => {
    const rows = parameterRows(sum, 3)
    expect(rows.map((row) => row.require)).toEqual([true, false, false, false])
    const ifRows = parameterRows(
      parseSyntaxParams('IF(logical_test, value_if_true, [value_if_false])'),
      0,
    )
    expect(ifRows.map((row) => row.require)).toEqual([true, true, false])
  })

  it('repeats a trailing group of parameters together', () => {
    const sumifs = [
      { name: 'sum_range', detail: '', require: true, repeat: false },
      { name: 'criteria_range1', detail: '', require: true, repeat: false },
      { name: 'criteria1', detail: '', require: true, repeat: false },
      { name: 'criteria_range2', detail: '', require: false, repeat: true },
      { name: 'criteria2', detail: '', require: false, repeat: true },
    ]
    expect(parameterRows(sumifs, 3).map((row) => row.name)).toEqual([
      'sum_range',
      'criteria_range1',
      'criteria1',
      'criteria_range2',
      'criteria2',
    ])
    expect(parameterRows(sumifs, 4).map((row) => row.name)).toEqual([
      'sum_range',
      'criteria_range1',
      'criteria1',
      'criteria_range2',
      'criteria2',
    ])
    const grown = parameterRows(sumifs, 5)
    expect(grown.map((row) => row.name).slice(5)).toEqual(['criteria_range3', 'criteria3'])
    expect(grown.map((row) => row.index)).toEqual([0, 1, 2, 3, 4, 5, 6])
  })

  it('numbers continuations of a name without a digit', () => {
    const rows = parameterRows([{ name: 'value', detail: '', require: true, repeat: true }], 2)
    expect(rows.map((row) => row.name)).toEqual(['value', 'value2', 'value3'])
  })
})

describe('formatPreviewValue', () => {
  it('renders scalars Excel-style', () => {
    expect(formatPreviewValue(42)).toBe('42')
    expect(formatPreviewValue(0.1 + 0.2)).toBe('0.3')
    expect(formatPreviewValue('text')).toBe('"text"')
    expect(formatPreviewValue('#NAME?')).toBe('#NAME?')
    expect(formatPreviewValue(true)).toBe('TRUE')
    expect(formatPreviewValue(null)).toBe('')
    expect(formatPreviewValue(undefined)).toBe('')
  })

  it('renders arrays as {row;row} and elides long ones', () => {
    expect(formatPreviewValue([[1], [2], [3]])).toBe('{1;2;3}')
    expect(
      formatPreviewValue([
        [1, 'a'],
        [true, null],
      ]),
    ).toBe('{1,"a";TRUE,}')
    expect(formatPreviewValue([[7]])).toBe('7')
    expect(formatPreviewValue([[1, 2, 3, 4, 5, 6, 7, 8]])).toBe('{1,2,3,4,5,6;...}')
  })
})

describe('pickedRangeRef', () => {
  it('writes relative A1 text, qualified only across sheets', () => {
    const range = { startRow: 0, endRow: 2, startColumn: 0, endColumn: 0 }
    expect(pickedRangeRef(range, 'Sheet1', 'Sheet1')).toBe('A1:A3')
    expect(pickedRangeRef({ ...range, endRow: 0 }, 'Sheet1', 'Sheet1')).toBe('A1')
    expect(pickedRangeRef(range, 'Data 2024', 'Sheet1')).toBe("'Data 2024'!A1:A3")
    expect(pickedRangeRef(range, 'Sheet2', 'Sheet1')).toBe('Sheet2!A1:A3')
  })

  it('writes whole rows and columns', () => {
    expect(
      pickedRangeRef(
        { startRow: 0, endRow: 999, startColumn: 1, endColumn: 2, rangeType: 2 },
        's',
        's',
      ),
    ).toBe('B:C')
    expect(
      pickedRangeRef(
        { startRow: 4, endRow: 4, startColumn: 0, endColumn: 99, rangeType: 1 },
        's',
        's',
      ),
    ).toBe('5:5')
  })

  it('quotes sheet names that are not bare identifiers', () => {
    expect(quoteSheetName('Sheet1')).toBe('Sheet1')
    expect(quoteSheetName("Bob's")).toBe("'Bob''s'")
    expect(quoteSheetName('2024')).toBe("'2024'")
    expect(quoteSheetName('AB12')).toBe("'AB12'")
  })
})
