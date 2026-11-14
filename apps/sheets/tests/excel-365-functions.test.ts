import {
  ArrayValueObject,
  ErrorType,
  NullValueObject,
  NumberValueObject,
  StringValueObject,
} from '@univerjs/engine-formula'
import type { BaseValueObject } from '@univerjs/engine-formula'
import { describe, expect, it } from 'vitest'

import {
  EXCEL_365_DESCRIPTIONS,
  PercentOf,
  RegexExtract,
  RegexReplace,
  RegexTest,
} from '../src/renderer/excel-365-functions'

const text = (value: string) => StringValueObject.create(value)
const num = (value: number) => NumberValueObject.create(value)

function column(values: BaseValueObject[]): ArrayValueObject {
  return ArrayValueObject.create({
    calculateValueList: values.map((value) => [value]),
    rowCount: values.length,
    columnCount: 1,
    unitId: '',
    sheetId: '',
    row: 0,
    column: 0,
  })
}

function rowValues(result: BaseValueObject): unknown[] {
  return (result as ArrayValueObject).getArrayValue()[0]!.map((cell) => cell?.getValue())
}

describe('REGEXTEST', () => {
  const fn = new RegexTest('REGEXTEST')

  it('is true when any part of the text matches', () => {
    expect(fn.calculate(text('Order 1234'), text('[0-9]+')).getValue()).toBe(true)
    expect(fn.calculate(text('no digits'), text('[0-9]')).getValue()).toBe(false)
  })

  it('is case-sensitive by default and insensitive with 1 or a leading (?i)', () => {
    expect(fn.calculate(text('Apple'), text('apple')).getValue()).toBe(false)
    expect(fn.calculate(text('Apple'), text('apple'), num(1)).getValue()).toBe(true)
    expect(fn.calculate(text('Apple'), text('(?i)apple')).getValue()).toBe(true)
  })

  it('applies elementwise over a text range', () => {
    const result = fn.calculate(column([text('a1'), text('b'), text('c3')]), text('[0-9]'))
    const values = (result as ArrayValueObject).getArrayValue().map((row) => row[0]!.getValue())
    expect(values).toEqual([true, false, true])
  })

  it('returns #VALUE! for an invalid pattern and passes errors through', () => {
    expect(fn.calculate(text('x'), text('(')).getValue()).toBe(ErrorType.VALUE)
    expect(fn.calculate(text('x'), text('x'), text('yes')).getValue()).toBe(ErrorType.VALUE)
  })
})

describe('REGEXEXTRACT', () => {
  const fn = new RegexExtract('REGEXEXTRACT')
  const phones = text('Call (555) 123-4567 or (555) 987-6543')

  it('returns the first match by default, not the capture groups', () => {
    expect(fn.calculate(text('Ana Silva'), text('([A-Z])[a-z]+')).getValue()).toBe('Ana')
  })

  it('returns every match with return_mode 1', () => {
    expect(rowValues(fn.calculate(phones, text('\\(\\d{3}\\) [\\d-]+'), num(1)))).toEqual([
      '(555) 123-4567',
      '(555) 987-6543',
    ])
  })

  it('returns the capture groups of the first match with return_mode 2', () => {
    const result = fn.calculate(text('2026-11-13'), text('(\\d{4})-(\\d{2})-(\\d{2})'), num(2))
    expect(rowValues(result)).toEqual(['2026', '11', '13'])
  })

  it('returns #N/A without a match and #VALUE! for an unknown mode', () => {
    expect(fn.calculate(text('abc'), text('[0-9]')).getValue()).toBe(ErrorType.NA)
    expect(fn.calculate(text('abc'), text('a'), num(3)).getValue()).toBe(ErrorType.VALUE)
  })

  it('honours case_sensitivity', () => {
    expect(fn.calculate(text('ABC'), text('b'), num(0), num(1)).getValue()).toBe('B')
  })
})

describe('REGEXREPLACE', () => {
  const fn = new RegexReplace('REGEXREPLACE')

  it('replaces every match by default and expands $n groups', () => {
    expect(fn.calculate(text('a1b22c333'), text('[0-9]+'), text('#')).getValue()).toBe('a#b#c#')
    expect(
      fn.calculate(text('SilvaAna'), text('([A-Z][a-z]+)([A-Z][a-z]+)'), text('$2 $1')).getValue(),
    ).toBe('Ana Silva')
  })

  it('replaces only the nth match, counting from the end when negative', () => {
    expect(fn.calculate(text('1-2-3'), text('[0-9]'), text('x'), num(2)).getValue()).toBe('1-x-3')
    expect(fn.calculate(text('1-2-3'), text('[0-9]'), text('x'), num(-1)).getValue()).toBe('1-2-x')
    expect(fn.calculate(text('1-2-3'), text('[0-9]'), text('x'), num(9)).getValue()).toBe('1-2-3')
  })

  it('honours case_sensitivity', () => {
    expect(fn.calculate(text('Aa'), text('a'), text('-'), num(0), num(1)).getValue()).toBe('--')
  })
})

describe('PERCENTOF', () => {
  const fn = new PercentOf('PERCENTOF')

  it('divides the subset sum by the total sum, ignoring text and blanks', () => {
    const all = column([num(10), num(30), text('n/a'), NullValueObject.create(), num(60)])
    expect(fn.calculate(column([num(10), num(30)]), all).getValue()).toBeCloseTo(0.4)
  })

  it('returns #DIV/0! when the total is zero', () => {
    expect(fn.calculate(num(0), column([num(0)])).getValue()).toBe(ErrorType.DIV_BY_ZERO)
  })
})

describe('EXCEL_365_DESCRIPTIONS', () => {
  it('describes every executor with Excel argument names', () => {
    const regexReplace = EXCEL_365_DESCRIPTIONS.find((info) => info.functionName === 'REGEXREPLACE')
    expect(regexReplace?.functionParameter.map((p) => p.name)).toEqual([
      'text',
      'pattern',
      'replacement',
      'occurrence',
      'case_sensitivity',
    ])
    expect(EXCEL_365_DESCRIPTIONS.map((info) => info.functionName)).toEqual([
      'REGEXTEST',
      'REGEXEXTRACT',
      'REGEXREPLACE',
      'PERCENTOF',
    ])
  })
})
