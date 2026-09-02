/**
 * Excel `horizontal="centerContinuous"` (Center Across Selection): the anchor
 * text centers across the run of trailing blank cells that share the format
 * (prod ref: a wrapText+centerContinuous title spans A:D on one line instead
 * of folding inside A). The loader must mark the run end on the anchor cell
 * and force CENTER + OVERFLOW so the render patch can widen the box.
 */
import { HorizontalAlign, WrapStrategy, type ICellData, type IStyleData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

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

function patchedMatrix(cells: WorkbookRangeResult['cells']): ICellData[][] {
  let captured: ICellData[][] = []
  const worksheet = {
    getRange: () => ({
      setValues: (matrix: ICellData[][]) => {
        captured = matrix
      },
    }),
  }
  patchWorksheetRangeInner(
    worksheet as never,
    undefined,
    { startRow: 0, endRow: 1, startColumn: 0, endColumn: 5 },
    cells,
    [centerContinuous, plain],
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
