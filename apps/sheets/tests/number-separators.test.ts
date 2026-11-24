import { describe, expect, it } from 'vitest'

import {
  CANONICAL_SEPARATORS,
  canonicalNumberToLocal,
  formulaToCanonical,
  formulaToLocal,
  localNumberToCanonical,
  separatorsForLocale,
} from '../src/renderer/number-separators'

const ptBR = separatorsForLocale('pt-BR')
const frFR = separatorsForLocale('fr-FR')

describe('separatorsForLocale', () => {
  it('reads the decimal comma and the group of a region', () => {
    expect(ptBR).toEqual({ decimal: ',', group: '.' })
    expect(separatorsForLocale('de-DE')).toEqual({ decimal: ',', group: '.' })
    expect(frFR).toEqual({ decimal: ',', group: ' ' })
  })

  it('keeps decimal-point regions and unknown tags canonical', () => {
    for (const tag of ['en-US', 'zh-CN', 'ja-JP', 'es-MX', 'de-CH', '', 'not a tag', undefined])
      expect(separatorsForLocale(tag)).toEqual(CANONICAL_SEPARATORS)
  })
})

describe('localNumberToCanonical', () => {
  it.each([
    ['5,47', '5.47'],
    ['10,04', '10.04'],
    ['-3,2', '-3.2'],
    [',5', '.5'],
    ['1.234,56', '1,234.56'],
    ['1.234', '1,234'],
    ['1.234.567', '1,234,567'],
    ['1,234', '1.234'],
    ['1,5%', '1.5%'],
    ['12,5e3', '12.5e3'],
    ['(1.234,5)', '(1,234.5)'],
    ['€ 5,47', '€ 5.47'],
    ['5,47 €', '5.47 €'],
  ])('%s → %s', (typed, canonical) => {
    expect(localNumberToCanonical(typed, ptBR)).toBe(canonical)
  })

  it.each([
    '5',
    '5.47',
    '1.5',
    '1,2,3',
    '31/12/2026',
    '12:30',
    '5,47 kg',
    'R$ 5,47',
    'abc',
    '',
    ',',
    '1.23,4',
  ])('leaves %j alone', (typed) => {
    expect(localNumberToCanonical(typed, ptBR)).toBeNull()
  })

  it('takes a space or a dot as the group of a space-grouping region', () => {
    expect(localNumberToCanonical('1 234,5', frFR)).toBe('1,234.5')
    expect(localNumberToCanonical('1 234,5', frFR)).toBe('1,234.5')
    expect(localNumberToCanonical('1.234,5', frFR)).toBe('1,234.5')
  })

  it('does nothing in a decimal-point region', () => {
    expect(localNumberToCanonical('5,47', CANONICAL_SEPARATORS)).toBeNull()
  })
})

describe('canonicalNumberToLocal', () => {
  it('writes the decimal comma into numeric editor text only', () => {
    expect(canonicalNumberToLocal(5.47, ptBR)).toBe('5,47')
    expect(canonicalNumberToLocal('150.5%', ptBR)).toBe('150,5%')
    expect(canonicalNumberToLocal(1.5e-7, ptBR)).toBe('1,5e-7')
    expect(canonicalNumberToLocal(12, ptBR)).toBe('12')
    expect(canonicalNumberToLocal('v1.2', ptBR)).toBe('v1.2')
    expect(canonicalNumberToLocal(5.47, CANONICAL_SEPARATORS)).toBe('5.47')
  })
})

describe('formula notation', () => {
  const pairs: Array<[canonical: string, local: string]> = [
    ['=SUM(A1,B1)*1.5', '=SUM(A1;B1)*1,5'],
    ['=IF(A1>1.5,"a,b;c",0.25)', '=IF(A1>1,5;"a,b;c";0,25)'],
    ['=SUM(1,5)', '=SUM(1;5)'],
    ['={1,2;3.5,4}', '={1\\2;3,5\\4}'],
    ["='My, sheet 1.5'!A1+C6/5", "='My, sheet 1.5'!A1+C6/5"],
    ['=SUM(Table1[[#This Row],[Q1.5]],2.5)', '=SUM(Table1[[#This Row],[Q1.5]];2,5)'],
    ['=STDEV.S(A1:A10,1.5E+3)', '=STDEV.S(A1:A10;1,5E+3)'],
    ['=LOG10(A1.5)+SUM(1:5,$1:$2)', '=LOG10(A1.5)+SUM(1:5;$1:$2)'],
    ['="unterminated, 1.5', '="unterminated, 1.5'],
  ]

  it.each(pairs)('%s ⇄ %s', (canonical, local) => {
    expect(formulaToLocal(canonical, ptBR)).toBe(local)
    expect(formulaToCanonical(local, ptBR)).toBe(canonical)
  })

  it('reads a decimal comma between digits and keeps other commas as separators', () => {
    expect(formulaToCanonical('=SUM(1,5;2)', ptBR)).toBe('=SUM(1.5,2)')
    expect(formulaToCanonical('=SUM(A1,B1)', ptBR)).toBe('=SUM(A1,B1)')
    expect(formulaToCanonical('=SUM(A1,5)', ptBR)).toBe('=SUM(A1,5)')
    expect(formulaToCanonical('=A1*1.5', ptBR)).toBe('=A1*1.5')
  })

  it('leaves non-formulas and decimal-point regions alone', () => {
    expect(formulaToLocal('1.5,2', ptBR)).toBe('1.5,2')
    expect(formulaToLocal('=SUM(A1,B1)*1.5', CANONICAL_SEPARATORS)).toBe('=SUM(A1,B1)*1.5')
    expect(formulaToCanonical('=SUM(A1;B1)', CANONICAL_SEPARATORS)).toBe('=SUM(A1;B1)')
  })
})
