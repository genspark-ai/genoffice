import { BorderStyleTypes, composeStyles } from '@univerjs/core'
import { describe, expect, it } from 'vitest'

import {
  recordRowStyleKeys,
  sheetRowColStyleKeys,
  toUniverStyle,
  withRowColOverrides,
} from '../src/renderer/univer-sync'
import type { LazyWorkbookState } from '../src/renderer/univer-state'
import type { WorkbookCellStyle } from '../src/shared/desktop-api'

const baseStyle: WorkbookCellStyle = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: false,
  shrinkToFit: false,
  diagonalUp: false,
  diagonalDown: false,
}

// <row s= customFormat> default: bold red 12pt (prod_001 xf16 shape).
const rowStyle: WorkbookCellStyle = {
  ...baseStyle,
  bold: true,
  fontSize: 12,
  fontColor: '#FF0000',
  fontFamily: 'Malgun Gothic',
  verticalAlignment: 'center',
}

// Cell xf: bold 10pt with NO font color (automatic — must render default).
const cellStyle: WorkbookCellStyle = {
  ...baseStyle,
  bold: true,
  fontSize: 10,
  fontFamily: 'Malgun Gothic',
  fillColor: '#DDEBF7',
}

function fakeState(styles: WorkbookCellStyle[], columnStyleIndex?: number): LazyWorkbookState {
  return {
    file: {
      styles,
      sheets: [
        {
          id: 'sheet-1',
          columnWidths:
            columnStyleIndex === undefined
              ? []
              : [{ startColumn: 0, endColumn: 0, styleIndex: columnStyleIndex }],
        },
      ],
    },
    rowColStyleKeys: new Map(),
  } as unknown as LazyWorkbookState
}

describe('row/column style bleed overrides', () => {
  it('fills exactly the row-style keys the cell xf leaves unset', () => {
    const bleed = new Set(Object.keys(toUniverStyle(rowStyle)))
    const s = withRowColOverrides(toUniverStyle(cellStyle), bleed, baseStyle)
    expect(s.cl).toEqual({ rgb: '#000000' }) // automatic color must block the row red
    expect(s.vt).toBe(0)
    expect(s.fs).toBe(10) // present keys stay untouched
    expect(s.bg).toEqual({ rgb: '#DDEBF7' })
  })

  it('composes to the cell values, never the row defaults, for styled cells', () => {
    const bleed = new Set(Object.keys(toUniverStyle(rowStyle)))
    const composed = composeStyles(
      toUniverStyle(rowStyle),
      withRowColOverrides(toUniverStyle(cellStyle), bleed, baseStyle),
    )
    expect(composed.cl).toEqual({ rgb: '#000000' })
    expect(composed.fs).toBe(10)
    // Unstyled cells (no override baked in) still inherit the row default.
    expect(composeStyles(toUniverStyle(rowStyle), {}).cl).toEqual({ rgb: '#FF0000' })
  })

  it('resolves missing font family/size from the workbook Normal style', () => {
    const bleed = new Set(['ff', 'fs'])
    const s = withRowColOverrides({}, bleed, {
      ...baseStyle,
      fontFamily: 'Batang',
      fontSize: 9,
    })
    expect(s.ff).toBe('Batang')
    expect(s.fs).toBe(9)
  })

  it('blocks a column-default border with a NONE edge the renderer skips', () => {
    // <col style=> carries a light-gray thin border; the cells own
    // an xf with borderId 1 (no border) and Excel paints no lines there.
    const columnStyle: WorkbookCellStyle = {
      ...baseStyle,
      borderTop: { style: 'thin', color: '#D5D3D1' },
      borderBottom: { style: 'thin', color: '#D5D3D1' },
      borderLeft: { style: 'thin', color: '#D5D3D1' },
      borderRight: { style: 'thin', color: '#D5D3D1' },
    } as WorkbookCellStyle
    const bleed = new Set(Object.keys(toUniverStyle(columnStyle)))
    const composed = composeStyles(
      toUniverStyle(columnStyle),
      withRowColOverrides(toUniverStyle(cellStyle), bleed, baseStyle),
    )
    expect(composed.bd).toEqual({ t: { s: BorderStyleTypes.NONE, cl: { rgb: '#000000' } } })
    // A cell with its own border keeps it untouched.
    const bordered = withRowColOverrides(
      { bd: { b: { s: BorderStyleTypes.THIN, cl: { rgb: '#000000' } } } },
      bleed,
      baseStyle,
    )
    expect(bordered.bd).toEqual({ b: { s: BorderStyleTypes.THIN, cl: { rgb: '#000000' } } })
  })

  it('leaves keys without a safe stand-in alone', () => {
    const s = withRowColOverrides({}, new Set(['pd']), baseStyle)
    expect(s).toEqual({})
  })

  it('is a no-op without any row/column styles', () => {
    const s = toUniverStyle(cellStyle)
    expect(withRowColOverrides({ ...s }, undefined, baseStyle)).toEqual(s)
    expect(withRowColOverrides({ ...s }, new Set(), baseStyle)).toEqual(s)
  })

  it('records row style keys and seeds column style keys per sheet', () => {
    const state = fakeState([cellStyle, rowStyle], 1)
    const seeded = sheetRowColStyleKeys(state, 'sheet-1')
    expect(seeded.has('cl')).toBe(true) // from the column default
    recordRowStyleKeys(state, 'sheet-1', [
      { row: 5, hidden: false, styleIndex: 1 },
    ] as unknown as Parameters<typeof recordRowStyleKeys>[2])
    expect(sheetRowColStyleKeys(state, 'sheet-1').has('vt')).toBe(true)
    // Rows without a style contribute nothing.
    const bare = fakeState([cellStyle])
    recordRowStyleKeys(bare, 'sheet-1', [{ row: 1, hidden: false }] as unknown as Parameters<
      typeof recordRowStyleKeys
    >[2])
    expect(sheetRowColStyleKeys(bare, 'sheet-1').size).toBe(0)
  })
})
