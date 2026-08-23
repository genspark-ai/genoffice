import { describe, expect, it } from 'vitest'

import { toUniverStyle } from '../src/renderer/univer-sync'
import type { WorkbookCellStyle } from '../src/shared/desktop-api'

const base: WorkbookCellStyle = {
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  wrapText: false,
  diagonalUp: false,
  diagonalDown: false,
}

describe('toUniverStyle fill', () => {
  it('emits an empty-rgb bg sentinel for a fill-less xf so a column-style fill cannot bleed through', () => {
    // Univer composes styles by key: a MISSING bg would inherit the
    // <col style=> fill, while Excel treats the cell xf as a full
    // replacement (prod_001: a light-blue column fill tinted every
    // explicitly-styled cell in the column). bg: null would be stripped
    // by SetRangeValues' removeNull, hence the empty-rgb sentinel.
    expect(toUniverStyle(base).bg).toEqual({ rgb: '' })
  })

  it('keeps explicit fills', () => {
    expect(toUniverStyle({ ...base, fillColor: '#FF0000' }).bg).toEqual({ rgb: '#FF0000' })
  })
})
