import {
  CellValueType,
  extractPureTextFromCell,
  type ICellData,
  WrapStrategy,
} from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  installFormulaNewlineDisplay,
  isFormulaDriven,
  multilineFormulaDisplayCell,
} from '../src/renderer/formula-newline-display'
import type { UniverRuntime } from '../src/renderer/univer-state'
import { cachedFormulaCellData } from '../src/renderer/univer-sync'

// A `t="str"` cache whose text carries a literal line break.
const FORMULA = '="\u3000A"&A17&"B\u3002\n\u3000C"'
const CACHE = '\u3000Ax B\u3002\n\u3000C'
const WRAP = { tb: WrapStrategy.WRAP }

describe('isFormulaDriven', () => {
  it('accepts a master formula and a shared-formula follower', () => {
    expect(isFormulaDriven({ f: '=A1', v: 'x' })).toBe(true)
    expect(isFormulaDriven({ si: '0', v: 'x' })).toBe(true)
  })

  it('rejects plain values and empty markers', () => {
    expect(isFormulaDriven({ v: 'x' })).toBe(false)
    expect(isFormulaDriven({ f: '', si: '', v: 'x' })).toBe(false)
    expect(isFormulaDriven(null)).toBe(false)
  })
})

describe('multilineFormulaDisplayCell', () => {
  it('the cached value alone loses the break at load: Univer strips CR/LF from plain v', () => {
    const cell = cachedFormulaCellData(FORMULA, CACHE)
    expect(cell).toEqual({ f: FORMULA, v: CACHE, t: CellValueType.STRING })
    expect(extractPureTextFromCell(cell)).toBe('\u3000Ax B\u3002\u3000C')
  })

  it('a wrapped formula result gets a paragraph document with the break', () => {
    const cell = cachedFormulaCellData(FORMULA, CACHE)
    const shown = multilineFormulaDisplayCell(cell, true, { tb: WrapStrategy.WRAP, fs: 11 })
    expect(shown?.p?.body?.dataStream).toBe('\u3000Ax B\u3002\r\u3000C\r\n')
    expect(shown?.p?.body?.paragraphs?.map((p) => p.startIndex)).toEqual([6, 9])
    expect(shown?.p?.body?.textRuns?.[0]?.ts).toEqual({ fs: 11 })
    expect(shown?.v).toBe(CACHE)
    expect(extractPureTextFromCell(shown)).toBe('\u3000Ax B\u3002\r\u3000C')
  })

  it('a shared-formula follower (si without f) gets the same document', () => {
    const follower = { si: 'shared-0', v: 'a\nb', t: CellValueType.STRING }
    const shown = multilineFormulaDisplayCell(follower, isFormulaDriven(follower), WRAP)
    expect(shown?.p?.body?.dataStream).toBe('a\rb\r\n')
    expect(shown?.si).toBe('shared-0')
  })

  it('CRLF results render as one break', () => {
    const shown = multilineFormulaDisplayCell(
      { f: '=A1&CHAR(13)&CHAR(10)&B1', v: 'a\r\nb', t: CellValueType.STRING },
      true,
      WRAP,
    )
    expect(shown?.p?.body?.dataStream).toBe('a\rb\r\n')
  })

  it('leaves non-wrapping, single-line, rich and non-formula cells alone', () => {
    const multi = { f: '=A1', v: 'a\nb', t: CellValueType.STRING }
    expect(multilineFormulaDisplayCell(multi, true, { tb: WrapStrategy.OVERFLOW })).toBe(multi)
    expect(multilineFormulaDisplayCell(multi, true, undefined)).toBe(multi)
    expect(multilineFormulaDisplayCell(multi, false, WRAP)).toBe(multi)
    const single = { f: '=A1', v: 'ab', t: CellValueType.STRING }
    expect(multilineFormulaDisplayCell(single, true, WRAP)).toBe(single)
    const rich = { ...multi, p: { id: 'd' } as ICellData['p'] }
    expect(multilineFormulaDisplayCell(rich, true, WRAP)).toBe(rich)
    expect(multilineFormulaDisplayCell(null, true, WRAP)).toBeNull()
  })
})

describe('installFormulaNewlineDisplay', () => {
  type Handler = (cell: unknown, location: unknown, next: (cell: unknown) => unknown) => unknown

  function installedHandler(): Handler {
    let captured: Handler | undefined
    const runtime = {
      univer: {
        __getInjector: () => ({
          get: () => ({
            intercept: (_point: unknown, interceptor: { handler: Handler }) => {
              captured = interceptor.handler
              return { dispose() {} }
            },
          }),
        }),
      },
    } as unknown as UniverRuntime
    installFormulaNewlineDisplay(runtime)
    if (!captured) throw new Error('interceptor not registered')
    return captured
  }

  function locationFor(rawData: ICellData): object {
    return {
      row: 3,
      col: 3,
      rawData,
      worksheet: {
        // Runs the CELL_CONTENT chain again: a handler calling it recurses.
        getComposedCellStyle(): never {
          throw new Error('re-entered getCell')
        },
        getComposedCellStyleByCellData: () => WRAP,
      },
    }
  }

  it('composes the style from the cell in hand instead of re-entering getCell', () => {
    const handler = installedHandler()
    const next = (cell: unknown): unknown => cell
    const error = { f: '=1/0', v: '#DIV/0!', t: CellValueType.STRING }
    expect(handler(error, locationFor(error), next)).toBe(error)
    const follower = { si: 'shared-0', v: 'a\nb', t: CellValueType.STRING }
    const shown = handler(follower, locationFor(follower), next) as ICellData
    expect(shown.p?.body?.dataStream).toBe('a\rb\r\n')
  })
})
