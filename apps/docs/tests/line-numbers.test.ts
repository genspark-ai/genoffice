// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import type { LineNumbering, SectionInfo } from '@genoffice/docx-engine'
import {
  columnAt,
  columnLefts,
  groupLineRects,
  numberLines,
} from '../src/renderer/editor/line-numbers'

const line = (page: number, top: number, section = 0, col = 0) => ({ page, col, top, section })

describe('numberLines', () => {
  it('counts every line in reading order and labels multiples of countBy', () => {
    const ln: LineNumbering = { countBy: 2, start: 1, restart: 'continuous' }
    const out = numberLines(
      [line(0, 40), line(0, 20), line(1, 0), line(0, 0, 0, 1), line(0, 60)],
      () => ln,
    )
    expect(out.map((o) => [o.line.page, o.line.col, o.line.top, o.label])).toEqual([
      [0, 0, 40, '2'],
      [0, 1, 0, '4'],
    ])
  })

  it('restarts per page or per section from w:start; continuous never restarts', () => {
    const lines = [line(0, 0, 0), line(0, 10, 0), line(1, 0, 0), line(1, 10, 1), line(2, 0, 1)]
    const labels = (restart: LineNumbering['restart']) =>
      numberLines(lines, () => ({ countBy: 1, start: 5, restart })).map((o) => o.label)
    expect(labels('continuous')).toEqual(['5', '6', '7', '8', '9'])
    expect(labels('newPage')).toEqual(['5', '6', '5', '6', '5'])
    expect(labels('newSection')).toEqual(['5', '6', '7', '5', '6'])
  })

  it('skips lines of sections without numbering without breaking the count', () => {
    const lines = [line(0, 0, 0), line(0, 10, 1), line(0, 20, 0)]
    const out = numberLines(lines, (s) =>
      s === 0 ? { countBy: 1, start: 1, restart: 'continuous' } : undefined,
    )
    expect(out.map((o) => [o.line.top, o.label])).toEqual([
      [0, '1'],
      [20, '2'],
    ])
  })
})

describe('groupLineRects', () => {
  it('merges fragments sharing a line box and keeps distinct lines apart', () => {
    const lines = groupLineRects([
      { top: 20, bottom: 40, left: 100, right: 160 },
      { top: 0, bottom: 20, left: 10, right: 50 },
      { top: 1, bottom: 19, left: 50, right: 90 },
      { top: 0, bottom: 0, left: 0, right: 0 },
    ])
    expect(lines).toEqual([
      { top: 0, bottom: 20, left: 10, right: 90 },
      { top: 20, bottom: 40, left: 100, right: 160 },
    ])
  })
})

describe('columnLefts', () => {
  const section = (sectPrXml: string): SectionInfo =>
    ({
      settings: {
        pageWidth: 12240,
        marginLeft: 1440,
        marginRight: 1440,
        columns: 2,
        colSpace: 720,
      },
      sectPrXml,
    }) as unknown as SectionInfo

  it('lists reading-order column lefts; a w:bidi section fills from the right', () => {
    expect(columnLefts(section('<w:sectPr/>'))).toEqual([0, 336])
    expect(columnLefts(section('<w:sectPr><w:bidi/></w:sectPr>'))).toEqual([336, 0])
    const unequal =
      '<w:cols w:num="2" w:equalWidth="0"><w:col w:w="3000" w:space="720"/><w:col w:w="5640"/></w:cols>'
    expect(columnLefts(section(`<w:sectPr>${unequal}</w:sectPr>`))).toEqual([0, 248])
    expect(columnLefts(section(`<w:sectPr>${unequal}<w:bidi/></w:sectPr>`))).toEqual([424, 0])
  })
})

describe('columnAt', () => {
  it('picks the column under x for plain and mirrored (w:bidi) lefts', () => {
    expect([columnAt([0, 336], -5), columnAt([0, 336], 10), columnAt([0, 336], 340)]).toEqual([
      0, 0, 1,
    ])
    expect([columnAt([336, 0], -5), columnAt([336, 0], 10), columnAt([336, 0], 340)]).toEqual([
      1, 1, 0,
    ])
    expect(columnAt([0], 500)).toBe(0)
  })
})
