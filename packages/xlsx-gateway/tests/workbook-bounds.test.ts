import { describe, expect, it } from 'vitest'
import {
  isGridCellAddress,
  MAX_GRID_COLUMNS,
  MAX_GRID_ROWS,
  parseSharedStringsXml,
} from '../src/gateway/xlsx-gateway'

describe('workbook open bounds', () => {
  it('accepts grid-edge addresses and rejects out-of-grid ones', () => {
    expect(MAX_GRID_ROWS).toBe(1_048_576)
    expect(MAX_GRID_COLUMNS).toBe(16_384)
    expect(isGridCellAddress('A1')).toBe(true)
    expect(isGridCellAddress('XFD1048576')).toBe(true)
    expect(isGridCellAddress('XFE1')).toBe(false)
    expect(isGridCellAddress('A1048577')).toBe(false)
    expect(isGridCellAddress('A0')).toBe(false)
    expect(isGridCellAddress('a1')).toBe(false)
    expect(isGridCellAddress('')).toBe(false)
  })

  it('parses shared strings', () => {
    const xml =
      `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2">` +
      `<si><t>hello</t></si><si><t>a</t><t>b</t></si></sst>`
    expect(parseSharedStringsXml(xml)).toEqual(['hello', 'ab'])
    expect(parseSharedStringsXml('<sst/>')).toEqual([])
  })
})
