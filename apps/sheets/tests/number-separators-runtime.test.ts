import { CellValueType } from '@univerjs/core'
import { afterEach, describe, expect, it } from 'vitest'

import {
  CANONICAL_SEPARATORS,
  getNumberSeparators,
  separatorsForLocale,
  setNumberSeparators,
  withCanonicalNumbers,
} from '../src/renderer/number-separators'
import { cellForEditor, cellFromEditor } from '../src/renderer/number-separators-runtime'
import { formatGeneral } from '../src/renderer/numfmt-fix'
import { numfmtPreview } from '../src/renderer/numfmt-dialog'

afterEach(() => setNumberSeparators(CANONICAL_SEPARATORS))

describe('decimal-comma region', () => {
  const enter = () => setNumberSeparators(separatorsForLocale('pt-BR'))

  it('opens numbers and formulas in the region notation', () => {
    enter()
    expect(cellForEditor({ v: 5.47, t: CellValueType.NUMBER })).toEqual({
      v: '5,47',
      t: CellValueType.NUMBER,
    })
    expect(cellForEditor({ v: '150.5%', t: CellValueType.NUMBER })?.v).toBe('150,5%')
    expect(cellForEditor({ f: '=SUM(A1,B1)*1.5', v: 3 })?.f).toBe('=SUM(A1;B1)*1,5')
    const text = { v: '5.47', t: CellValueType.STRING }
    expect(cellForEditor(text)).toBe(text)
  })

  it('commits typed numbers and formulas in canonical notation', () => {
    enter()
    expect(cellFromEditor({ v: '5,47' })?.v).toBe('5.47')
    expect(cellFromEditor({ v: '1.234,5' })?.v).toBe('1,234.5')
    expect(cellFromEditor({ f: '=SUM(A1;B1)*1,5', v: null })?.f).toBe('=SUM(A1,B1)*1.5')
    for (const kept of [
      { v: '5,47', t: CellValueType.STRING },
      { v: '5,47', t: CellValueType.FORCE_STRING },
      { v: 'a, b' },
      { v: '31/12/2026' },
    ])
      expect(cellFromEditor(kept)).toBe(kept)
  })

  it('formats display text with the region separators', () => {
    enter()
    expect(formatGeneral(5.47, 11)).toBe('5,47')
    expect(formatGeneral(645577.79, 8)).toBe('645577,8')
    expect(formatGeneral(123456789.5, 7)).toBe('1,2E+08')
    expect(numfmtPreview('#,##0.00', 1234.5)).toBe('1.234,50')
    expect(numfmtPreview('mmm d, yyyy', 46000)).toBe('Dec 9, 2025')
  })

  it('reads canonical inside withCanonicalNumbers', () => {
    enter()
    withCanonicalNumbers(() => {
      expect(getNumberSeparators()).toEqual(CANONICAL_SEPARATORS)
      expect(formatGeneral(5.47, 11)).toBe('5.47')
      expect(numfmtPreview('#,##0.00', 1234.5)).toBe('1,234.50')
    })
    expect(formatGeneral(5.47, 11)).toBe('5,47')
  })
})

describe('decimal-point region', () => {
  it('changes nothing', () => {
    const number = { v: 5.47, t: CellValueType.NUMBER }
    const typed = { v: '5,47' }
    const formula = { f: '=SUM(A1,B1)*1.5' }
    expect(cellForEditor(number)).toBe(number)
    expect(cellForEditor(formula)).toBe(formula)
    expect(cellFromEditor(typed)).toBe(typed)
    expect(cellFromEditor(formula)).toBe(formula)
    expect(formatGeneral(5.47, 11)).toBe('5.47')
    expect(numfmtPreview('#,##0.00', 1234.5)).toBe('1,234.50')
  })
})
