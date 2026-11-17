import { describe, expect, it } from 'vitest'
import {
  neutralizeCsvFormulaField,
  startsLikeSpreadsheetFormula,
} from '../src/domain/csv-formula-guard'

describe('startsLikeSpreadsheetFormula', () => {
  it('flags the shapes the next app actually evaluates', () => {
    expect(startsLikeSpreadsheetFormula("=cmd|' /c calc'!A0")).toBe(true)
    expect(startsLikeSpreadsheetFormula('@SUM(A1)')).toBe(true)
    expect(startsLikeSpreadsheetFormula('+SUM(A1)')).toBe(true)
    expect(startsLikeSpreadsheetFormula('-SUM(A1)')).toBe(true)
    expect(startsLikeSpreadsheetFormula('\tcmd')).toBe(true)
    expect(startsLikeSpreadsheetFormula('=1+2')).toBe(true)
    // the DDE form: program|topic!cell behind a Lotus-prefix character
    expect(startsLikeSpreadsheetFormula("+cmd|' /C calc'!A0")).toBe(true)
    expect(startsLikeSpreadsheetFormula('@server|book!A1')).toBe(true)
  })

  it('leaves the shapes that do not evaluate as data (the review false positives)', () => {
    // an accounting-format zero / a bare dash / a bullet-ish line
    expect(startsLikeSpreadsheetFormula('-')).toBe(false)
    expect(startsLikeSpreadsheetFormula('-$1,234.50')).toBe(false)
    expect(startsLikeSpreadsheetFormula('-€3')).toBe(false)
    expect(startsLikeSpreadsheetFormula('- item')).toBe(false)
    // a social handle: @ without a function call or a DDE pipe
    expect(startsLikeSpreadsheetFormula('@handle')).toBe(false)
    // a plus/minus sign that opens nothing executable
    expect(startsLikeSpreadsheetFormula('-(A1)')).toBe(false)
    expect(startsLikeSpreadsheetFormula('+word')).toBe(false)
  })

  it('skips leading whitespace like a spreadsheet does', () => {
    expect(startsLikeSpreadsheetFormula(' =cmd()')).toBe(true)
    expect(startsLikeSpreadsheetFormula('  -SUM(A1)')).toBe(true)
    expect(startsLikeSpreadsheetFormula(' -5')).toBe(false)
    expect(startsLikeSpreadsheetFormula(' - item')).toBe(false)
  })

  it('leaves ordinary text and numbers alone', () => {
    expect(startsLikeSpreadsheetFormula('hello')).toBe(false)
    expect(startsLikeSpreadsheetFormula('1+2')).toBe(false)
    expect(startsLikeSpreadsheetFormula('A1=B2')).toBe(false)
    expect(startsLikeSpreadsheetFormula('-5')).toBe(false)
    expect(startsLikeSpreadsheetFormula('-1.5e3')).toBe(false)
    expect(startsLikeSpreadsheetFormula('+.5')).toBe(false)
    expect(startsLikeSpreadsheetFormula('+42')).toBe(false)
    expect(startsLikeSpreadsheetFormula('')).toBe(false)
  })
})

describe('neutralizeCsvFormulaField', () => {
  it('prepends the text marker to formula-shaped fields', () => {
    expect(neutralizeCsvFormulaField('=cmd')).toBe("'=cmd")
    expect(neutralizeCsvFormulaField('+SUM(A1)')).toBe("'+SUM(A1)")
    expect(neutralizeCsvFormulaField("+cmd|' /C calc'!A0")).toBe("'+cmd|' /C calc'!A0")
  })

  it('passes ordinary data through untouched', () => {
    expect(neutralizeCsvFormulaField('-5')).toBe('-5')
    expect(neutralizeCsvFormulaField('-$1,234.50')).toBe('-$1,234.50')
    expect(neutralizeCsvFormulaField('@handle')).toBe('@handle')
    expect(neutralizeCsvFormulaField('hello')).toBe('hello')
  })
})
