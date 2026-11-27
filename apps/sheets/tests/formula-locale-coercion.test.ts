/**
 * #2883: text numbers coerce by the session's separators, like Excel. The
 * reporter's workbook stores "15,68"-shaped text cells (pt-BR data), and
 * "=C6/5" read the comma as a thousands group — 1568/5 = 313.6 instead of
 * 3.136. The harness computes the same formulas on two sessions: the en-US
 * engine semantics (the value the engine produces today) and the pt-language
 * coercion.
 */
import { LocaleType, LogLevel, Univer, UniverInstanceType } from '@univerjs/core'
import { FUniver } from '@univerjs/core/facade'
import { UniverFormulaEnginePlugin } from '@univerjs/engine-formula'
import { UniverSheetsPlugin } from '@univerjs/sheets'
import '@univerjs/sheets/facade'
import { UniverSheetsFormulaPlugin } from '@univerjs/sheets-formula'
import '@univerjs/sheets-formula/facade'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'

import {
  installFormulaLocaleCoercion,
  localeUsesCommaDecimal,
  parseCommaDecimal,
} from '../src/renderer/formula-locale-coercion'

interface Session {
  read(cell: string): unknown
  dispose(): void
}

const settle = () => new Promise((resolve) => setTimeout(resolve, 1500))

async function session(lang: 'en' | 'pt'): Promise<Session> {
  const univer = new Univer({ locale: LocaleType.EN_US, locales: {}, logLevel: LogLevel.ERROR })
  univer.registerPlugin(UniverFormulaEnginePlugin)
  univer.registerPlugin(UniverSheetsPlugin)
  univer.registerPlugin(UniverSheetsFormulaPlugin)
  univer.createUnit(UniverInstanceType.UNIVER_SHEET, {
    id: 'wb',
    sheetOrder: ['s1'],
    sheets: { s1: { id: 's1', name: 'S1', rowCount: 20, columnCount: 8, cellData: {} } },
  })
  const coercion = installFormulaLocaleCoercion(() => lang)
  await settle()
  const api = FUniver.newAPI(univer)
  const sheet = api.getActiveWorkbook()!.getActiveSheet()
  // the reporter's data: measurements as comma-decimal TEXT cells
  sheet.getRange('C6').setValue('15,68')
  sheet.getRange('B3').setValue('5,47')
  sheet.getRange('A1').setValue('1.234,56')
  sheet.getRange('A2').setValue('15.68') // dot shape: the engine's own path, both sessions
  sheet.getRange('F6').setFormula('=C6/5')
  sheet.getRange('B4').setFormula('=B3*2')
  sheet.getRange('A3').setFormula('=A1*2')
  sheet.getRange('A4').setFormula('=A2/2')
  await settle()
  const read = (cell: string): unknown => sheet.getRange(cell).getCellData()?.v
  return {
    read,
    dispose: () => {
      coercion.dispose()
      univer.dispose()
    },
  }
}

describe('text numbers coerce by the session separators (#2883)', () => {
  let en: Session
  let pt: Session

  beforeAll(async () => {
    en = await session('en')
    pt = await session('pt')
  }, 45_000)

  afterAll(() => {
    en.dispose()
    pt.dispose()
  })

  it('a comma-decimal session computes; the en session keeps the engine shape', () => {
    // pt: "15,68" is 15.68 → 3.136; en engine: not a number → #VALUE!,
    // the shape the reporter saw
    expect(pt.read('F6')).toBeCloseTo(3.136, 6)
    expect(en.read('F6')).toBe('#VALUE!')
    expect(pt.read('B4')).toBeCloseTo(10.94, 6)
    // the dot shape stays the engine's own value in both sessions
    expect(pt.read('A4')).toBeCloseTo(7.84, 6)
    expect(en.read('A4')).toBeCloseTo(7.84, 6)
    // the pt wholesale shape "1.234,56" is 1234.56 in pt; the en engine
    // refuses it and *2 over #VALUE! stays #VALUE!
    expect(pt.read('A3')).toBeCloseTo(2469.12, 6)
    expect(en.read('A3')).toBe('#VALUE!')
  })

  it('the helper round-trips the strict shapes', () => {
    expect(localeUsesCommaDecimal('pt')).toBe(true)
    expect(localeUsesCommaDecimal('en')).toBe(false)
    expect(parseCommaDecimal('15,68')).toBe(15.68)
    expect(parseCommaDecimal('1.234,56')).toBe(1234.56)
    expect(parseCommaDecimal('5,47')).toBe(5.47)
  })
})
