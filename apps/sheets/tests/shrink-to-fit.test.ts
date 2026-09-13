import type { ICellData, IStyleData } from '@univerjs/core'
import { FontCache } from '@univerjs/engine-render'
import { beforeAll, describe, expect, it } from 'vitest'

import { SHRINK_TO_FIT_KEY } from '../src/renderer/numfmt-fix'
import {
  patchWorksheetRangeInner,
  shrinkToFitApplies,
  shrinkToFitFontSize,
} from '../src/renderer/univer-sync'
import type { WorkbookCellStyle, WorkbookRangeResult } from '../src/shared/desktop-api'

// 10px per character at the base size; scales linearly with font size.
const measure = (line: string): number => line.length * 10

describe('shrinkToFitFontSize', () => {
  it('returns null when the text already fits', () => {
    expect(shrinkToFitFontSize('1980', 12, 50, measure)).toBeNull()
  })

  it('scales the font down proportionally and floors to an integer', () => {
    // 5 chars * 10px = 50px into 42px: 12 * 42/50 = 10.08 -> 10
    expect(shrinkToFitFontSize('1980X', 12, 42, measure)).toBe(10)
  })

  it('uses the widest line of a multiline cell', () => {
    expect(shrinkToFitFontSize('ab\r\nabcdef', 12, 30, measure)).toBe(6)
  })

  it('clamps at 1pt and rejects a zero-width column', () => {
    expect(shrinkToFitFontSize('abcdefghij', 12, 1, measure)).toBe(1)
    expect(shrinkToFitFontSize('abc', 12, 0, measure)).toBeNull()
  })
})

describe('shrinkToFitApplies', () => {
  const base: WorkbookCellStyle = {
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    wrapText: false,
    diagonalUp: false,
    diagonalDown: false,
  }
  const shrink: WorkbookCellStyle = { ...base, shrinkToFit: true }

  it('applies to plain shrinkToFit text', () => {
    expect(shrinkToFitApplies(shrink, 'MOBILIZATION', false)).toBe(true)
  })

  it('yields to wrapText when both flags are set (Excel greys shrink out)', () => {
    expect(shrinkToFitApplies({ ...shrink, wrapText: true }, 'MOBILIZATION', false)).toBe(false)
  })

  it('skips empty text, pending formulas and unstyled cells', () => {
    expect(shrinkToFitApplies(shrink, '', false)).toBe(false)
    expect(shrinkToFitApplies(shrink, '=A1', true)).toBe(false)
    expect(shrinkToFitApplies(undefined, 'x', false)).toBe(false)
    expect(shrinkToFitApplies({ ...base, wrapText: true }, 'x', false)).toBe(false)
  })
})

// Fake 2D context: every glyph is half an em wide.
function installHalfEmMeasure(): void {
  const ctx = {
    font: '',
    textBaseline: 'alphabetic',
    measureText(text: string) {
      const match = /([\d.]+)pt/.exec(this.font)
      const px = match ? (Number(match[1]) * 96) / 72 : 0
      return {
        width: text.length * px * 0.5,
        fontBoundingBoxAscent: px * 0.75,
        fontBoundingBoxDescent: px * 0.25,
      }
    },
  }
  ;(FontCache as unknown as { _context: unknown })._context = ctx
}

function patchedCells(
  cells: WorkbookRangeResult['cells'],
  styles: readonly WorkbookCellStyle[],
  columnWidth: number,
): ICellData[][] {
  let captured: ICellData[][] = []
  const worksheet = {
    getRange: () => ({
      setValues: (matrix: ICellData[][]) => {
        captured = matrix
      },
    }),
    getSheet: () => ({ getMergedCell: () => null, getColumnWidth: () => columnWidth }),
  }
  patchWorksheetRangeInner(
    worksheet as never,
    undefined,
    { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
    cells,
    styles,
    [],
    [],
    null,
    false,
  )
  return captured
}

describe('shrinkToFit on numeric cells', () => {
  const base: WorkbookCellStyle = {
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    wrapText: false,
    diagonalUp: false,
    diagonalDown: false,
    fontSize: 10,
    numberFormat: '#,##0,\\ ',
  }
  const shrink: WorkbookCellStyle = { ...base, shrinkToFit: true }

  beforeAll(installHalfEmMeasure)

  it('shrinks the formatted number instead of leaving it to the #### rule', () => {
    // "105,303 " = 8 glyphs * 6.67px = 53px into a 30px column (25px budget):
    // 10pt * 25 / 53 -> 4pt.
    const cell = patchedCells(
      [{ row: 0, column: 0, value: 105303159, styleIndex: 0 }],
      [shrink],
      30,
    )[0]?.[0]
    expect((cell?.s as IStyleData).fs).toBe(4)
    expect(cell?.custom).toEqual({ [SHRINK_TO_FIT_KEY]: true })
  })

  it('leaves a number without shrinkToFit at full size for the #### rule', () => {
    const cell = patchedCells(
      [{ row: 0, column: 0, value: 105303159, styleIndex: 0 }],
      [base],
      30,
    )[0]?.[0]
    expect((cell?.s as IStyleData).fs).toBe(10)
    expect(cell?.custom).toBeUndefined()
  })

  it('marks a fitting number too (Excel never hashes a shrinkToFit cell) but not text', () => {
    const row = patchedCells(
      [
        { row: 0, column: 0, value: 105303159, styleIndex: 0 },
        { row: 0, column: 1, value: 'MOBILIZATION', styleIndex: 0 },
      ],
      [shrink],
      120,
    )[0]
    expect((row?.[0]?.s as IStyleData).fs).toBe(10)
    expect(row?.[0]?.custom).toEqual({ [SHRINK_TO_FIT_KEY]: true })
    expect((row?.[1]?.s as IStyleData).fs).toBe(10)
    expect(row?.[1]?.custom).toBeUndefined()
  })
})
