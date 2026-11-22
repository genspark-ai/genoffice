/**
 * Word's Columns dialog: the number of columns stays between one and the most
 * the dialog offers; the Left and Right presets are two columns with the
 * narrow one on that side; the page lays out the widths the section was given.
 */
import { describe, expect, it } from 'vitest'
import type { SectionInfo, SectionSettings } from '@genoffice/docx-engine'
import {
  MAX_COLUMNS,
  clampColumns,
  equalColumnWidth,
  fitColumns,
  unequalPresetWidths,
} from '../src/renderer/components/ColumnsDialog'
import { columnRuleXs, hasColumnRule } from '../src/renderer/editor/column-rules'
import { sectionColGeom } from '../src/renderer/pagination-sections'

const LETTER = {
  pageWidth: 12240,
  pageHeight: 15840,
  marginLeft: 1440,
  marginRight: 1440,
  columns: 2,
  colSpace: 720,
} as SectionSettings

const sectionOf = (settings: Partial<SectionSettings>, sectPrXml = ''): SectionInfo =>
  ({ settings: { ...LETTER, ...settings }, sectPrXml }) as SectionInfo

describe('columns dialog', () => {
  it('keeps the number of columns in range and whole', () => {
    expect(clampColumns('2')).toBe(2)
    expect(clampColumns('0')).toBe(1)
    expect(clampColumns('-3')).toBe(1)
    expect(clampColumns('99')).toBe(MAX_COLUMNS)
    expect(clampColumns('2.4')).toBe(2)
    expect(clampColumns('abc')).toBe(1)
  })

  it('sizes the Left and Right presets: a third of the text width less the gap, and the rest', () => {
    expect(unequalPresetWidths('left', LETTER)).toEqual([2640, 6000])
    expect(unequalPresetWidths('right', LETTER)).toEqual([6000, 2640])
    expect(unequalPresetWidths('left', { ...LETTER, colSpace: 360 })).toEqual([2880, 6120])
    expect(unequalPresetWidths('left', { ...LETTER, pageWidth: 5000 })).toBeNull()
  })
})

describe('columns of unequal width', () => {
  it('start from the width equal columns have', () => {
    expect(equalColumnWidth(LETTER, 2, 720)).toBe(4320)
    expect(equalColumnWidth(LETTER, 3, 360)).toBe(2880)
  })

  it('let the last column take up what a typed width or gap leaves', () => {
    expect(fitColumns(LETTER, [4320, 4320], [720], { width: 0, value: 3000 })).toEqual({
      widths: [3000, 5640],
      spaces: [720],
    })
    expect(fitColumns(LETTER, [3000, 5640], [720], { space: 0, value: 360 })).toEqual({
      widths: [3000, 6000],
      spaces: [360],
    })
    // the last column typed: the first one gives way
    expect(fitColumns(LETTER, [2880, 2880, 2880], [360, 360], { width: 2, value: 4000 })).toEqual({
      widths: [1760, 2880, 4000],
      spaces: [360, 360],
    })
  })

  it('stop a typed value where the other column would get too narrow', () => {
    expect(fitColumns(LETTER, [4320, 4320], [720], { width: 0, value: 20000 }).widths).toEqual([
      7920, 720,
    ])
    expect(fitColumns(LETTER, [4320, 4320], [720], { width: 0, value: 10 }).widths).toEqual([
      720, 7920,
    ])
    expect(fitColumns(LETTER, [4320, 4320], [720], { space: 0, value: 20000 })).toEqual({
      widths: [4320, 720],
      spaces: [4320],
    })
  })
})

describe('column geometry of a section edited in the session', () => {
  const FILE_UNEQUAL =
    '<w:sectPr><w:cols w:num="2" w:space="720" w:equalWidth="0"><w:col w:w="3000" w:space="360"/><w:col w:w="6000"/></w:cols></w:sectPr>'

  it('lays out widths the file does not hold yet', () => {
    const geom = sectionColGeom(sectionOf({ colWidths: [2640, 6000], colEqualWidth: false }))
    expect(geom.equalWidth).toBe(false)
    expect(geom.widths).toEqual([176, 400])
    expect(geom.gaps).toEqual([48])
  })

  it('keeps reading the file for widths it already holds, its own gaps included', () => {
    const geom = sectionColGeom(
      sectionOf({ colWidths: [3000, 6000], colEqualWidth: false }, FILE_UNEQUAL),
    )
    expect(geom.widths).toEqual([200, 400])
    expect(geom.gaps).toEqual([24])
  })

  it('lays out a gap changed in the session beside widths the file holds', () => {
    const geom = sectionColGeom(
      sectionOf({ colWidths: [3000, 6000], colSpaces: [360], colEqualWidth: false }, FILE_UNEQUAL),
    )
    expect(geom.gaps).toEqual([24])
    const moved = sectionColGeom(
      sectionOf({ colWidths: [3000, 6000], colSpaces: [240], colEqualWidth: false }, FILE_UNEQUAL),
    )
    expect(moved.widths).toEqual([200, 400])
    expect(moved.gaps).toEqual([16])
  })

  it('evens out a section whose file still lists widths', () => {
    const geom = sectionColGeom(sectionOf({ colEqualWidth: true }, FILE_UNEQUAL))
    expect(geom.equalWidth).toBe(true)
    expect(geom.widths).toEqual([288, 288])
  })

  it('draws the line in the middle of each gap, only where there are columns', () => {
    const ruled = sectionOf({ colSep: true })
    expect(hasColumnRule(ruled)).toBe(true)
    expect(columnRuleXs(ruled)).toEqual([312])
    expect(columnRuleXs(sectionOf({ colSep: true, columns: 3 }))).toEqual([200, 424])
    expect(
      columnRuleXs(sectionOf({ colSep: true, colWidths: [2640, 6000], colEqualWidth: false })),
    ).toEqual([200])
    expect(hasColumnRule(sectionOf({ colSep: true, columns: 1 }))).toBe(false)
    expect(hasColumnRule(sectionOf({}))).toBe(false)
  })
})
