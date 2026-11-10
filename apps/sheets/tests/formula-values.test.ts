import { beforeEach, describe, expect, it } from 'vitest'
import type { CellState } from '@genoffice/xlsx-gateway/domain/workbook.types'
import {
  clearVerifiedFormulaValues,
  formulaTargetsFromOps,
  rememberFormulaValue,
  verifiedFormulaValues,
} from '../src/renderer/formula-values'

const reader =
  (cells: Record<string, Record<string, CellState>>) => (addresses: string[], sheetId: string) => {
    const sheet = cells[sheetId]
    if (!sheet) throw new Error(`no sheet ${sheetId}`)
    return Object.fromEntries(addresses.filter((a) => a in sheet).map((a) => [a, sheet[a]!]))
  }

describe('formulaTargetsFromOps', () => {
  it('keeps only set_formula ops, once per cell', () => {
    const ops = [
      { op: 'set_cell', sheetId: 's1', address: 'A1', value: 5 },
      { op: 'set_formula', sheetId: 's1', address: 'B1', formula: '=A1*2' },
      { op: 'set_formula', sheetId: 's1', address: 'B1', formula: '=A1*3' },
      { op: 'set_formula', sheetId: 's1', address: 'C1', formula: 'A1' },
    ]
    expect(formulaTargetsFromOps(ops)).toEqual([{ sheetId: 's1', address: 'B1' }])
  })
})

describe('verifiedFormulaValues', () => {
  beforeEach(() => clearVerifiedFormulaValues())

  it('writes the live value, not the one observed when the batch settled', () => {
    rememberFormulaValue('s1', 'B1', { formula: '=A1*2' })
    const read = reader({ s1: { B1: { value: 14, formula: '=A1*2' } } })
    expect(verifiedFormulaValues(read)).toEqual([{ sheetId: 's1', row: 0, column: 1, value: 14 }])
  })

  it('drops a cell whose formula changed, is unsettled, or whose sheet is gone', () => {
    rememberFormulaValue('s1', 'B1', { formula: '=A1*2' })
    rememberFormulaValue('s1', 'B2', { formula: '=A1' })
    rememberFormulaValue('gone', 'A1', { formula: '=1' })
    const read = reader({
      s1: { B1: { value: 7, formula: '=A1*3' }, B2: { value: null, formula: '=A1' } },
    })
    expect(verifiedFormulaValues(read)).toEqual([])
  })

  it('prefers the raw model value, maps Excel errors and skips the engine failure literal', () => {
    rememberFormulaValue('s1', 'A1', { formula: '=TODAY()' })
    rememberFormulaValue('s1', 'A2', { formula: '=1/0' })
    rememberFormulaValue('s1', 'A3', { formula: '=NOSUCHFN()' })
    const read = reader({
      s1: {
        A1: { value: '2026-09-17', formula: '=TODAY()', rawValue: 46282 },
        A2: { value: '#DIV/0!', formula: '=1/0' },
        A3: { value: '#ERROR!', formula: '=NOSUCHFN()' },
      },
    })
    expect(verifiedFormulaValues(read)).toEqual([
      { sheetId: 's1', row: 0, column: 0, value: 46282 },
      { sheetId: 's1', row: 1, column: 0, value: { error: '#DIV/0!' } },
    ])
  })
})
