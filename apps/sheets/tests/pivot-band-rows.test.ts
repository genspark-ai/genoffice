import type { ICellData } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import type { WorkbookFile } from '../src/shared/desktop-api'
import { applyPivotStyling } from '../src/renderer/univer-sync'

/// Ground truth from POI sample 54436: <location ref="A8:B11"
/// firstHeaderRow="1" firstDataRow="1"/> where A8 = "Row Labels" (header),
/// A9/A10 = data, A11 = Grand Total. firstDataRow is the offset of the FIRST
/// DATA row inside the ref, so header rows are offsets 0..firstDataRow-1.
type PivotSpec = Omit<WorkbookFile['sheets'][number]['pivotTables'][number], 'path' | 'cachePath'>

type CellStyle = { bg?: { rgb?: string }; cl?: { rgb?: string }; bl?: number }

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
  return paintCells(pivot).map((row) => ((row[0]?.s ?? {}) as CellStyle).bg?.rgb)
}

/// `A1`-style lookup into the painted matrix (range origin A1).
function at(cells: ICellData[][], ref: string): CellStyle {
  const column = ref.charCodeAt(0) - 65
  const row = Number(ref.slice(1)) - 1
  return (cells[row]?.[column]?.s ?? {}) as CellStyle
}

/// Calibration workbook shape (genoffice-sample/sheets/calib/gen_pivot_calib.py):
/// A1:D11, two header rows, row-label column A, grand-total column D, rows
/// 3/7 level-1 subheadings, 4/5/8/9 data, 6/10 subtotals, 11 grand total.
const CALIB_LAYOUT = {
  outputRef: 'A1:D11',
  firstDataRow: 2,
  firstDataCol: 1,
  rowGrandTotals: true,
  rowKinds: 'sddtsddtg',
  styled: true,
} satisfies PivotSpec

/// PivotStyleDark23 under the Office 2007 theme, as the sidecar emits it
/// with showRowStripes and showColStripes on.
const DARK23: PivotSpec = {
  ...CALIB_LAYOUT,
  headerFill: '#376092',
  headerBold: false,
  firstHeaderCellFontColor: '#FFFFFF',
  firstHeaderCellBold: true,
  wholeTableFill: '#4F81BD',
  wholeTableFontColor: '#DCE6F2',
  secondRowStripeFill: '#95B3D7',
  secondColumnStripeFill: '#95B3D7',
  firstColumnFill: '#376092',
  firstColumnBold: false,
  subheadingFontColor: '#FFFFFF',
  subheadingBold: true,
  subheading2FontColor: '#FFFFFF',
  subheading2Bold: true,
  subtotalBold: false,
  totalRowFontColor: '#FFFFFF',
  totalRowBold: true,
}

/// PivotStyleMedium23 (Office 2007 theme), stripes on.
const MEDIUM23: PivotSpec = {
  ...CALIB_LAYOUT,
  headerBold: true,
  wholeTableFill: '#DCE6F2',
  wholeTableFontColor: '#376092',
  secondRowStripeFill: '#B9CDE5',
  secondColumnStripeFill: '#B9CDE5',
  firstColumnFill: '#B9CDE5',
  firstColumnBold: true,
  subheadingFontColor: '#000000',
  subheadingBold: true,
  subtotalFontColor: '#000000',
  subtotalBold: true,
  totalRowBold: true,
}

describe('applyPivotStyling', () => {
  it('paints only the header offset(s) below firstDataRow and the grand-total row', () => {
    // Sample 54436 shape: A8:B11, firstDataRow=1 → header A8, data A9/A10, total A11.
    const painted = paint({
      outputRef: 'A8:B11',
      headerFill: '#DCE6F1',
      totalRowFill: '#DCE6F1',
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
      totalRowFill: '#DCE6F1',
      firstDataRow: 2,
      rowGrandTotals: true,
    })
    expect(painted[0]).toBe(true)
    expect(painted[1]).toBe(true)
    expect(painted[2]).toBe(false) // first data row
    expect(painted[5]).toBe(true) // grand total
  })

  it('leaves the last row as plain data when rowGrandTotals is false', () => {
    const rowFills = fills({
      outputRef: 'A8:B11',
      headerFill: '#DCE6F1',
      totalRowFill: '#DCE6F1',
      firstDataRow: 1,
      rowGrandTotals: false,
    })
    expect(rowFills[7]).toBe('#DCE6F1')
    expect(rowFills[10]).toBeUndefined()
  })

  it('paints nothing without a pivot style', () => {
    const painted = paint({ outputRef: 'A8:B11', firstDataRow: 1, rowGrandTotals: true })
    expect(painted.every((row) => !row)).toBe(true)
  })

  // Light 1-7 with stripes off: no fills survive, but a named style still
  // bolds the header and grand-total bands.
  it('bolds band rows of a styled pivot even when every fill is absent', () => {
    const spec: PivotSpec = {
      outputRef: 'A8:B11',
      styled: true,
      firstDataRow: 1,
      rowGrandTotals: true,
    }
    const cells = paintCells(spec)
    expect(at(cells, 'A8').bl).toBe(1) // header
    expect(cells[8]?.[0]?.s).toBeUndefined() // data rows untouched
    expect(cells[9]?.[0]?.s).toBeUndefined()
    expect(at(cells, 'A11').bl).toBe(1) // grand total
    expect(fills(spec).every((fill) => fill === undefined)).toBe(true)
  })

  // Ground truth from oxmlsdk_RelationalPivotB6 (PivotStyleLight2,
  // showRowStripes): unfilled header, first stripe from the first data row.
  it('stripes alternate data rows and leave a fill-less header unfilled but bold', () => {
    const spec: PivotSpec = {
      outputRef: 'A8:B14',
      stripeFill: '#DCE6F1',
      totalRowFill: '#FFFFFF',
      firstDataRow: 1,
      rowGrandTotals: true,
    }
    const cells = paintCells(spec)
    const rowFills = fills(spec)
    expect(rowFills[7]).toBeUndefined() // A8 header: no fill for Light 1-7
    expect(at(cells, 'A8').bl).toBe(1) // ...but bold
    expect(rowFills[8]).toBe('#DCE6F1') // A9 first data row starts the stripe
    expect(rowFills[9]).toBeUndefined()
    expect(rowFills[10]).toBe('#DCE6F1')
    expect(rowFills[11]).toBeUndefined()
    expect(rowFills[12]).toBe('#DCE6F1')
    expect(rowFills[13]).toBe('#FFFFFF') // A14 grand total: explicit white covers the stripe
  })

  // Ground truth from aspose_sample1 (PivotStyleMedium9): solid accent header
  // with white bold text.
  it('paints the header font color with the header fill', () => {
    const cells = paintCells({
      outputRef: 'A8:B11',
      headerFill: '#4F81BD',
      headerFontColor: '#FFFFFF',
      headerBold: true,
      firstDataRow: 1,
      rowGrandTotals: false,
    })
    const header = at(cells, 'A8')
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
      headerBold: true,
      totalRowBold: true,
      firstDataRow: 1,
      rowGrandTotals: true,
    })
    const header = at(cells, 'A8')
    expect(header.bg?.rgb).toBe('#4F81BD')
    expect(header.cl?.rgb).toBe('#FFFFFF')
    const total = at(cells, 'A11')
    expect(total.bg).toBeUndefined()
    expect(total.cl).toBeUndefined()
    expect(total.bl).toBe(1)
  })

  // PivotStyleDark1 (LO pivot_dark1 + calibration): dk1 tint 0.5 header and
  // grand total with white bold text over a 0.75 body; the SECOND stripe
  // (0.65) paints the odd offsets — opposite to the table-style Dark stripes.
  it('paints the Dark1 body with second-stripe rows and a filled grand total', () => {
    const cells = paintCells({
      outputRef: 'A8:B12',
      headerFill: '#808080',
      headerFontColor: '#FFFFFF',
      headerBold: true,
      wholeTableFill: '#BFBFBF',
      secondRowStripeFill: '#A6A6A6',
      totalRowFill: '#808080',
      totalRowFontColor: '#FFFFFF',
      totalRowBold: true,
      firstDataRow: 1,
      rowGrandTotals: true,
    })
    expect(at(cells, 'A9').bg?.rgb).toBe('#BFBFBF') // offset 0: body
    expect(at(cells, 'A10').bg?.rgb).toBe('#A6A6A6') // offset 1: second stripe
    expect(at(cells, 'A11').bg?.rgb).toBe('#BFBFBF')
    const total = at(cells, 'A12')
    expect(total.bg?.rgb).toBe('#808080')
    expect(total.cl?.rgb).toBe('#FFFFFF')
    expect(total.bl).toBe(1)
  })

  // Excel truth for PivotStyleDark23 (pivot-style-truths.json, rowStripes /
  // colStripes variants).
  it('layers the Dark23 bands like Excel', () => {
    const cells = paintCells(DARK23)
    // First header cell: white bold over the shaded header.
    expect(at(cells, 'A1')).toMatchObject({ bg: { rgb: '#376092' }, cl: { rgb: '#FFFFFF' }, bl: 1 })
    // The rest of the header keeps the tint-0.8 body text, not bold.
    expect(at(cells, 'B1')).toMatchObject({ bg: { rgb: '#376092' }, cl: { rgb: '#DCE6F2' } })
    expect(at(cells, 'B1').bl).toBeUndefined()
    // Column stripes skip the header rows.
    expect(at(cells, 'C2').bg?.rgb).toBe('#376092')
    // Level-1 subheading: row-label column keeps its shade, values show the
    // solid body; both white bold.
    expect(at(cells, 'A3')).toMatchObject({ bg: { rgb: '#376092' }, cl: { rgb: '#FFFFFF' }, bl: 1 })
    expect(at(cells, 'B3')).toMatchObject({ bg: { rgb: '#4F81BD' }, cl: { rgb: '#FFFFFF' }, bl: 1 })
    // Offset 1 (odd): second row stripe over the data columns, the row-label
    // column covers it.
    expect(at(cells, 'A4').bg?.rgb).toBe('#376092')
    expect(at(cells, 'A4').bl).toBeUndefined()
    expect(at(cells, 'B4').bg?.rgb).toBe('#95B3D7')
    expect(at(cells, 'D4').bg?.rgb).toBe('#95B3D7')
    // Offset 2 (even): body, with the second column stripe on column C only.
    expect(at(cells, 'B5').bg?.rgb).toBe('#4F81BD')
    expect(at(cells, 'C5').bg?.rgb).toBe('#95B3D7')
    expect(at(cells, 'D5').bg?.rgb).toBe('#4F81BD')
    // Subtotal (offset 3): no band fill, so the stripe shows; not bold.
    expect(at(cells, 'B6').bg?.rgb).toBe('#95B3D7')
    expect(at(cells, 'B6').bl).toBeUndefined()
    // Grand total: no fill of its own (body + column stripe show through),
    // white bold text; the row-label cell keeps the first-column shade.
    expect(at(cells, 'A11')).toMatchObject({
      bg: { rgb: '#376092' },
      cl: { rgb: '#FFFFFF' },
      bl: 1,
    })
    expect(at(cells, 'B11')).toMatchObject({
      bg: { rgb: '#4F81BD' },
      cl: { rgb: '#FFFFFF' },
      bl: 1,
    })
    expect(at(cells, 'C11').bg?.rgb).toBe('#95B3D7')
  })

  // Excel truth for PivotStyleMedium23: unfilled header over a tinted body.
  it('layers the Medium23 bands like Excel', () => {
    const cells = paintCells(MEDIUM23)
    // Header shows the whole-table fill and its shaded text, bold; no column
    // stripe on header cells; the row-label header cell takes the column fill.
    expect(at(cells, 'B1')).toMatchObject({ bg: { rgb: '#DCE6F2' }, cl: { rgb: '#376092' }, bl: 1 })
    expect(at(cells, 'C2').bg?.rgb).toBe('#DCE6F2')
    expect(at(cells, 'A2')).toMatchObject({ bg: { rgb: '#B9CDE5' }, cl: { rgb: '#376092' }, bl: 1 })
    // Subheading row: black bold everywhere, row-label column fill stays.
    expect(at(cells, 'A3')).toMatchObject({ bg: { rgb: '#B9CDE5' }, cl: { rgb: '#000000' }, bl: 1 })
    expect(at(cells, 'B3')).toMatchObject({ bg: { rgb: '#DCE6F2' }, cl: { rgb: '#000000' }, bl: 1 })
    // Data rows: second stripe on odd offsets, second column stripe on C.
    expect(at(cells, 'B4').bg?.rgb).toBe('#B9CDE5')
    expect(at(cells, 'B5').bg?.rgb).toBe('#DCE6F2')
    expect(at(cells, 'C5').bg?.rgb).toBe('#B9CDE5')
    expect(at(cells, 'B5').bl).toBeUndefined()
    expect(at(cells, 'A5').bl).toBe(1) // row-label column is bold
    // Grand total: no fill, shaded bold text.
    expect(at(cells, 'B11')).toMatchObject({
      bg: { rgb: '#DCE6F2' },
      cl: { rgb: '#376092' },
      bl: 1,
    })
  })

  // Excel truth (top3 layout, three row fields): a deeper subheading has no
  // bold of its own — its label is bold only through the row-label column,
  // its values stay regular.
  it('lets an unflagged subheading band inherit bold from the row-label column', () => {
    const cells = paintCells({ ...MEDIUM23, outputRef: 'A1:D17', rowKinds: 'sSddSddsSddSddg' })
    expect(at(cells, 'A4').bl).toBe(1)
    expect(at(cells, 'B4').bl).toBeUndefined()
    expect(at(cells, 'B4').bg?.rgb).toBe('#B9CDE5') // offset 1: second stripe
    expect(at(cells, 'A3').bl).toBe(1) // level-1 subheading is bold everywhere
    expect(at(cells, 'B3').bl).toBe(1)
  })

  it('treats rows as data with a trailing grand total when rowKinds is absent', () => {
    const cells = paintCells({
      ...DARK23,
      outputRef: 'A1:D6',
      rowKinds: undefined,
    })
    expect(at(cells, 'B3').bl).toBeUndefined() // would be a bold subheading with rowKinds
    expect(at(cells, 'B4').bg?.rgb).toBe('#95B3D7')
    expect(at(cells, 'B6')).toMatchObject({ cl: { rgb: '#FFFFFF' }, bl: 1 })
  })

  it('keeps explicit cell fills and non-black font colors, still applying bold', () => {
    const range = { startRow: 0, startColumn: 0, endRow: 10, endColumn: 3 }
    const matrix: ICellData[][] = Array.from({ length: 11 }, () =>
      Array.from({ length: 4 }, () => ({})),
    )
    matrix[0]![1] = { s: { bg: { rgb: '#FF0000' }, cl: { rgb: '#00FF00' } } }
    applyPivotStyling(matrix, range, [
      { path: 'xl/pivotTables/pivotTable1.xml', cachePath: null, ...DARK23 },
    ])
    expect(at(matrix, 'B1')).toEqual({ bg: { rgb: '#FF0000' }, cl: { rgb: '#00FF00' } })
    expect(at(matrix, 'C1').bg?.rgb).toBe('#376092')
    expect(at(matrix, 'A1').bl).toBe(1)
  })
})
