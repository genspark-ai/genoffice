/**
 * Excel `horizontal="centerContinuous"` (Center Across Selection): the anchor
 * text centers across the run of trailing blank cells that share the format
 * (prod ref: a wrapText+centerContinuous title spans A:D on one line instead
 * of folding inside A). The loader must mark the run end on the anchor cell
 * and force CENTER + OVERFLOW so the render patch can widen the box.
 */
import { HorizontalAlign, WrapStrategy, type ICellData, type IStyleData } from '@univerjs/core'
import { FontCache } from '@univerjs/engine-render'
import { beforeAll, describe, expect, it } from 'vitest'

import { CENTER_ACROSS_END_KEY } from '../src/renderer/center-continuous'
import { patchWorksheetRangeInner } from '../src/renderer/univer-sync'
import type { WorkbookCellStyle, WorkbookRangeResult } from '../src/shared/desktop-api'

const centerContinuous: WorkbookCellStyle = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: true,
  diagonalUp: false,
  diagonalDown: false,
  horizontalAlignment: 'centerContinuous',
}

const plain: WorkbookCellStyle = { ...centerContinuous, horizontalAlignment: undefined }

const COLUMN_WIDTH = 20

function patchedMatrix(
  cells: WorkbookRangeResult['cells'],
  styles: readonly WorkbookCellStyle[] = [centerContinuous, plain],
): ICellData[][] {
  let captured: ICellData[][] = []
  const worksheet = {
    getRange: () => ({
      setValues: (matrix: ICellData[][]) => {
        captured = matrix
      },
    }),
    getSheet: () => ({ getMergedCell: () => null, getColumnWidth: () => COLUMN_WIDTH }),
  }
  patchWorksheetRangeInner(
    worksheet as never,
    undefined,
    { startRow: 0, endRow: 1, startColumn: 0, endColumn: 5 },
    cells,
    styles,
    [],
    [],
    null,
    false,
  )
  return captured
}

describe('centerContinuous run marking', () => {
  it('marks the anchor with the run end over trailing blank same-format cells', () => {
    const matrix = patchedMatrix([
      { row: 0, column: 0, value: 'Title', styleIndex: 0 },
      { row: 0, column: 1, value: null, styleIndex: 0 },
      { row: 0, column: 2, value: null, styleIndex: 0 },
      { row: 0, column: 3, value: 'stops the run', styleIndex: 1 },
    ])
    expect(matrix[0]?.[0]?.custom).toEqual({ [CENTER_ACROSS_END_KEY]: 2 })
    expect(matrix[0]?.[1]?.custom).toBeUndefined()
  })

  it('ends the run at the next content cell, which anchors its own run', () => {
    const matrix = patchedMatrix([
      { row: 0, column: 0, value: 'A', styleIndex: 0 },
      { row: 0, column: 1, value: 'B', styleIndex: 0 },
      { row: 0, column: 2, value: null, styleIndex: 0 },
    ])
    expect(matrix[0]?.[0]?.custom).toBeUndefined()
    expect(matrix[0]?.[1]?.custom).toEqual({ [CENTER_ACROSS_END_KEY]: 2 })
  })

  it('leaves single-cell runs unmarked and never wraps centerContinuous text', () => {
    const matrix = patchedMatrix([{ row: 0, column: 0, value: 'Alone', styleIndex: 0 }])
    const style = matrix[0]?.[0]?.s as IStyleData
    expect(matrix[0]?.[0]?.custom).toBeUndefined()
    expect(style.ht).toBe(HorizontalAlign.CENTER)
    // wrapText is set in the xf but Excel keeps the title on one line.
    expect(style.tb).toBe(WrapStrategy.OVERFLOW)
  })

  it('treats formula cells as content that ends a run', () => {
    const matrix = patchedMatrix([
      { row: 0, column: 0, value: 'Title', styleIndex: 0 },
      { row: 0, column: 1, value: null, styleIndex: 0 },
      { row: 0, column: 2, value: null, formula: 'A1&""', styleIndex: 0 },
    ])
    expect(matrix[0]?.[0]?.custom).toEqual({ [CENTER_ACROSS_END_KEY]: 1 })
  })
})

describe('shrinkToFit on a centerContinuous anchor', () => {
  const shrinkAnchor: WorkbookCellStyle = {
    ...centerContinuous,
    wrapText: false,
    shrinkToFit: true,
    fontSize: 9,
  }

  beforeAll(() => {
    // Every glyph is half an em wide: 'ABCDEF' at 9pt (12px) measures 36px.
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
  })

  it('measures the shrink budget across the run, not the anchor column', () => {
    // Run A:D = 80px (75px budget) holds the 36px label at 9pt.
    const matrix = patchedMatrix(
      [
        { row: 0, column: 0, value: 'ABCDEF', styleIndex: 0 },
        { row: 0, column: 1, value: null, styleIndex: 0 },
        { row: 0, column: 2, value: null, styleIndex: 0 },
        { row: 0, column: 3, value: null, styleIndex: 0 },
      ],
      [shrinkAnchor],
    )
    expect((matrix[0]?.[0]?.s as IStyleData).fs).toBe(9)
    expect(matrix[0]?.[0]?.custom).toEqual({ [CENTER_ACROSS_END_KEY]: 3 })
  })

  it('still shrinks a single-cell run against its own column', () => {
    // 20px column (15px budget): 9pt * 15 / 36 -> 3pt.
    const matrix = patchedMatrix(
      [{ row: 0, column: 0, value: 'ABCDEF', styleIndex: 0 }],
      [shrinkAnchor],
    )
    expect((matrix[0]?.[0]?.s as IStyleData).fs).toBe(3)
  })
})
