import type { ICellData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import { applyPivotStyling } from '../src/renderer/univer-sync'

/// Ground truth from POI sample 54436: <location ref="A8:B11"
/// firstHeaderRow="1" firstDataRow="1"/> where A8 = "Row Labels" (header),
/// A9/A10 = data, A11 = Grand Total. firstDataRow is the offset of the FIRST
/// DATA row inside the ref, so header rows are offsets 0..firstDataRow-1.
interface PivotSpec {
  outputRef: string
  headerFill?: string
  headerFontColor?: string
  wholeTableFill?: string
  stripeFill?: string
  styled?: boolean
  firstDataRow?: number
  rowGrandTotals?: boolean
}

function paintCells(pivot: PivotSpec): ICellData[][] {
  const range = { startRow: 0, startColumn: 0, endRow: 19, endColumn: 4 }
  const matrix: ICellData[][] = Array.from({ length: 20 }, () =>
    Array.from({ length: 5 }, () => ({})),
  )
  applyPivotStyling(matrix, range, [
    { path: 'xl/pivotTables/pivotTable1.xml', cachePath: null, ...pivot },
  ])
  return matrix
}

function paint(pivot: PivotSpec): boolean[] {
  return paintCells(pivot).map((row) => Boolean(row[0]?.s))
}

function fills(pivot: PivotSpec): (string | undefined)[] {
  return paintCells(pivot).map((row) => ((row[0]?.s ?? {}) as { bg?: { rgb?: string } }).bg?.rgb)
}

describe('applyPivotStyling', () => {
  it('paints only the header offset(s) below firstDataRow and the grand-total row', () => {
    // Sample 54436 shape: A8:B11, firstDataRow=1 → header A8, data A9/A10, total A11.
    const painted = paint({
      outputRef: 'A8:B11',
      headerFill: '#DCE6F1',
      firstDataRow: 1,
      rowGrandTotals: true,
    })
    expect(painted[7]).toBe(true) // A8 "Row Labels" header
    expect(painted[8]).toBe(false) // A9 first data row stays unpainted
    expect(painted[9]).toBe(false) // A10 data
    expect(painted[10]).toBe(true) // A11 Grand Total
  })

  it('paints two header rows when firstDataRow is 2', () => {
    const painted = paint({
      outputRef: 'A1:C6',
      headerFill: '#DCE6F1',
      firstDataRow: 2,
      rowGrandTotals: true,
    })
    expect(painted[0]).toBe(true)
    expect(painted[1]).toBe(true)
    expect(painted[2]).toBe(false) // first data row
    expect(painted[5]).toBe(true) // grand total
  })

  it('leaves the last row unpainted when rowGrandTotals is false', () => {
    const painted = paint({
      outputRef: 'A8:B11',
      headerFill: '#DCE6F1',
      firstDataRow: 1,
      rowGrandTotals: false,
    })
    expect(painted[7]).toBe(true)
    expect(painted[10]).toBe(false)
  })

  it('paints nothing without a pivot style', () => {
    const painted = paint({ outputRef: 'A8:B11', firstDataRow: 1, rowGrandTotals: true })
    expect(painted.every((row) => !row)).toBe(true)
  })

  // Light 1-7 with stripes off: no fills survive, but a named style still
  // bolds the header and grand-total bands.
  it('bolds band rows of a styled pivot even when every fill is absent', () => {
    const cells = paintCells({
      outputRef: 'A8:B11',
      styled: true,
      firstDataRow: 1,
      rowGrandTotals: true,
    })
    const bold = (row: number) => ((cells[row]?.[0]?.s ?? {}) as { bl?: number }).bl
    expect(bold(7)).toBe(1) // A8 header
    expect(cells[8]?.[0]?.s).toBeUndefined() // data rows untouched
    expect(cells[9]?.[0]?.s).toBeUndefined()
    expect(bold(10)).toBe(1) // A11 grand total
    const rowFills = fills({
      outputRef: 'A8:B11',
      styled: true,
      firstDataRow: 1,
      rowGrandTotals: true,
    })
    expect(rowFills.every((fill) => fill === undefined)).toBe(true)
  })

  // Ground truth from oxmlsdk_RelationalPivotB6 (PivotStyleLight2,
  // showRowStripes): unfilled header, stripes from the first data row.
  it('stripes alternate data rows and leave a fill-less header unfilled but bold', () => {
    const cells = paintCells({
      outputRef: 'A8:B14',
      stripeFill: '#DCE6F1',
      firstDataRow: 1,
      rowGrandTotals: true,
    })
    const rowFills = fills({
      outputRef: 'A8:B14',
      stripeFill: '#DCE6F1',
      firstDataRow: 1,
      rowGrandTotals: true,
    })
    expect(rowFills[7]).toBeUndefined() // A8 header: no fill for Light 1-7
    expect((cells[7]?.[0]?.s as { bl?: number }).bl).toBe(1) // ...but bold
    expect(rowFills[8]).toBe('#DCE6F1') // A9 first data row starts the stripe
    expect(rowFills[9]).toBeUndefined()
    expect(rowFills[10]).toBe('#DCE6F1')
    expect(rowFills[11]).toBeUndefined()
    expect(rowFills[12]).toBe('#DCE6F1')
    expect(rowFills[13]).toBeUndefined() // A14 grand total: band, not stripe
  })

  // Ground truth from aspose_sample1 (PivotStyleMedium9): solid accent header
  // with white bold text.
  it('paints the header font color with the header fill', () => {
    const cells = paintCells({
      outputRef: 'A8:B11',
      headerFill: '#4F81BD',
      headerFontColor: '#FFFFFF',
      firstDataRow: 1,
      rowGrandTotals: false,
    })
    const header = cells[7]?.[0]?.s as { bg?: { rgb?: string }; cl?: { rgb?: string }; bl?: number }
    expect(header.bg?.rgb).toBe('#4F81BD')
    expect(header.cl?.rgb).toBe('#FFFFFF')
    expect(header.bl).toBe(1)
  })

  // Ground truth from aspose_sample1 (Medium9, G6:H7 with grand totals): the
  // solid header does not bleed onto the grand-total row, which Excel leaves
  // plain black-on-white bold.
  it('keeps the solid header treatment off the grand-total row', () => {
    const cells = paintCells({
      outputRef: 'A8:B11',
      headerFill: '#4F81BD',
      headerFontColor: '#FFFFFF',
      firstDataRow: 1,
      rowGrandTotals: true,
    })
    type CellStyle = { bg?: { rgb?: string }; cl?: { rgb?: string }; bl?: number }
    const header = cells[7]?.[0]?.s as CellStyle
    expect(header.bg?.rgb).toBe('#4F81BD')
    expect(header.cl?.rgb).toBe('#FFFFFF')
    const total = cells[10]?.[0]?.s as CellStyle
    expect(total.bg).toBeUndefined()
    expect(total.cl).toBeUndefined()
    expect(total.bl).toBe(1)
  })

  // Dark1 fills its body, so its calibrated grand-total band keeps the
  // header fill and white text.
  it('extends the header treatment to the grand total when the body is filled', () => {
    const cells = paintCells({
      outputRef: 'A8:B11',
      headerFill: '#808080',
      headerFontColor: '#FFFFFF',
      wholeTableFill: '#A6A6A6',
      stripeFill: '#BFBFBF',
      firstDataRow: 1,
      rowGrandTotals: true,
    })
    const total = cells[10]?.[0]?.s as { bg?: { rgb?: string }; cl?: { rgb?: string }; bl?: number }
    expect(total.bg?.rgb).toBe('#808080')
    expect(total.cl?.rgb).toBe('#FFFFFF')
    expect(total.bl).toBe(1)
  })
})
