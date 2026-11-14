import { describe, expect, it } from 'vitest'
import {
  neutralizeCsvFormulaField,
  startsLikeSpreadsheetFormula,
} from '../src/domain/csv-formula-guard'

describe('startsLikeSpreadsheetFormula', () => {
  it('flags the spreadsheet formula prefixes', () => {
    expect(startsLikeSpreadsheetFormula("=cmd|' /c calc'!A0")).toBe(true)
    expect(startsLikeSpreadsheetFormula('@SUM(A1)')).toBe(true)
    expect(startsLikeSpreadsheetFormula('\tcmd')).toBe(true)
    expect(startsLikeSpreadsheetFormula('=1+2')).toBe(true)
  })

  it('flags + and - unless they open a plain number', () => {
    expect(startsLikeSpreadsheetFormula('-cmd')).toBe(true)
    expect(startsLikeSpreadsheetFormula('+SUM(A1)')).toBe(true)
    expect(startsLikeSpreadsheetFormula('-(A1)')).toBe(true)
    expect(startsLikeSpreadsheetFormula('-5')).toBe(false)
    expect(startsLikeSpreadsheetFormula('-1.5e3')).toBe(false)
    expect(startsLikeSpreadsheetFormula('+.5')).toBe(false)
    expect(startsLikeSpreadsheetFormula('+42')).toBe(false)
  })

  it('skips leading whitespace like a spreadsheet does', () => {
    expect(startsLikeSpreadsheetFormula(' =cmd()')).toBe(true)
    expect(startsLikeSpreadsheetFormula('  -cmd')).toBe(true)
    expect(startsLikeSpreadsheetFormula(' -5')).toBe(false)
  })

  it('leaves ordinary text and numbers alone', () => {
    expect(startsLikeSpreadsheetFormula('hello')).toBe(false)
    expect(startsLikeSpreadsheetFormula('1+2')).toBe(false)
    expect(startsLikeSpreadsheetFormula('A1=B2')).toBe(false)
    expect(startsLikeSpreadsheetFormula('')).toBe(false)
  })
})

describe('neutralizeCsvFormulaField', () => {
  it('prepends the text marker to formula-shaped fields', () => {
    expect(neutralizeCsvFormulaField('=cmd')).toBe("'=cmd")
    expect(neutralizeCsvFormulaField('-cmd')).toBe("'-cmd")
  })

  it('passes numbers and text through untouched', () => {
    expect(neutralizeCsvFormulaField('-5')).toBe('-5')
    expect(neutralizeCsvFormulaField('hello')).toBe('hello')
  })
})
