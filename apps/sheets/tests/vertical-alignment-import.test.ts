import { VerticalAlign } from '@univerjs/core'
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

describe('toUniverStyle verticalAlignment', () => {
  it('maps the three native values', () => {
    expect(toUniverStyle({ ...base, verticalAlignment: 'top' }).vt).toBe(VerticalAlign.TOP)
    expect(toUniverStyle({ ...base, verticalAlignment: 'center' }).vt).toBe(VerticalAlign.MIDDLE)
    expect(toUniverStyle({ ...base, verticalAlignment: 'bottom' }).vt).toBe(VerticalAlign.BOTTOM)
  })

  it('centres justify and distributed like Excel does for a single line', () => {
    expect(toUniverStyle({ ...base, verticalAlignment: 'justify' }).vt).toBe(VerticalAlign.MIDDLE)
    expect(toUniverStyle({ ...base, verticalAlignment: 'distributed' }).vt).toBe(
      VerticalAlign.MIDDLE,
    )
  })

  it('leaves vt unset when the xf has no vertical alignment', () => {
    expect(toUniverStyle(base)).not.toHaveProperty('vt')
    expect(toUniverStyle({ ...base, verticalAlignment: 'bogus' })).not.toHaveProperty('vt')
  })
})
