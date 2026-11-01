import { describe, expect, it } from 'vitest'
import type { PmNode } from '../src/renderer/editor/convert'
import {
  AVERY_5160,
  cellContentOf,
  customGridLayout,
  emptyParagraph,
  envelopeLayout,
  labelCellSize,
  labelSheetDoc,
  layoutSection,
} from '../src/renderer/mail-merge/layout'
import type { SectionSettings } from '@genoffice/docx-engine'

const paragraph = (text: string): PmNode => ({
  type: 'docParagraph',
  content: [{ type: 'text', text }],
})

const cellText = (cell: PmNode): string | undefined => {
  const para = cell.content?.[0]
  const run = para?.content?.[0]
  return run && 'text' in run ? run.text : undefined
}

describe('label geometry', () => {
  it('the Avery 5160 stock sums to its paper', () => {
    const grid = AVERY_5160.grid!
    const cell = labelCellSize(AVERY_5160)
    // the derived cells split the usable area exactly: margins + grid + gaps = paper
    expect(
      AVERY_5160.marginLeft +
        grid.columns * cell.width +
        (grid.columns - 1) * grid.gapX +
        AVERY_5160.marginRight,
    ).toBe(AVERY_5160.pageWidth)
    expect(
      AVERY_5160.marginTop +
        grid.rows * cell.height +
        (grid.rows - 1) * grid.gapY +
        AVERY_5160.marginBottom,
    ).toBe(AVERY_5160.pageHeight)
  })

  it('a custom grid splits the usable area between the cells', () => {
    const layout = customGridLayout('a4', 2, 7)
    const cell = labelCellSize(layout)
    expect(cell.width * 2).toBeLessThanOrEqual(
      layout.pageWidth - layout.marginLeft - layout.marginRight,
    )
    expect(cell.height * 7).toBeLessThanOrEqual(
      layout.pageHeight - layout.marginTop - layout.marginBottom,
    )
  })
})

describe('layoutSection', () => {
  const base: SectionSettings = {
    pageWidth: 11906,
    pageHeight: 16838,
    orientation: 'portrait',
    marginTop: 1440,
    marginRight: 1440,
    marginBottom: 1440,
    marginLeft: 1440,
    pageBorder: false,
    columns: 1,
  }

  it('carries the layout paper and margins, keeping the rest of the section', () => {
    const layout = envelopeLayout('size10')
    const section = layoutSection(layout, base)
    expect(section.pageWidth).toBe(layout.pageWidth)
    expect(section.pageHeight).toBe(layout.pageHeight)
    expect(section.marginTop).toBe(layout.marginTop)
    expect(section.pageBorder).toBe(false)
  })

  it('a wider-than-tall paper is landscape (the envelope sizes)', () => {
    expect(layoutSection(envelopeLayout('size10'), base).orientation).toBe('landscape')
    expect(layoutSection(envelopeLayout('dl'), base).orientation).toBe('landscape')
  })
})

describe('labelSheetDoc', () => {
  const layout = AVERY_5160

  it('fills the grid row-major and pads the last row with empty cells', () => {
    const cells = Array.from({ length: 25 }, (_, i) => [paragraph(`r${i + 1}`)])
    const doc = labelSheetDoc(cells, layout)
    const table = doc.content![0]!
    const rows = table.content!
    // 25 records on a 3-wide grid: 8 full rows and one with 1 record + 2 empty
    expect(rows).toHaveLength(9)
    const first = rows[0]!.content!
    expect(first).toHaveLength(3)
    expect(cellText(first[0]!)).toBe('r1')
    expect(cellText(first[2]!)).toBe('r3')
    const last = rows[8]!.content!
    expect(cellText(last[0]!)).toBe('r25')
    expect(last[1]!.content).toEqual([emptyParagraph()])
    expect(last[2]!.content).toEqual([emptyParagraph()])
  })

  it('the rows carry the stock cell height and the first row carries the widths', () => {
    const cells = [[paragraph('a')]]
    const doc = labelSheetDoc(cells, layout)
    const table = doc.content![0]!
    expect(table.attrs?.docxIndex).toBeNull()
    const row = table.content![0]!
    const cell = labelCellSize(layout)
    expect(row.attrs?.heightTwips).toBe(cell.height)
    expect(row.attrs?.heightRule).toBe('exact')
    expect(row.content![0]!.attrs?.colwidth).toEqual([Math.round(cell.width / 15)])
  })

  it('a single record still gets a full first row', () => {
    const doc = labelSheetDoc([[paragraph('only')]], layout)
    expect(doc.content![0]!.content![0]!.content).toHaveLength(3)
  })
})

describe('cellContentOf', () => {
  it('keeps paragraphs and list items, turns headings into paragraphs', () => {
    const doc: PmNode = {
      type: 'doc',
      content: [
        paragraph('name'),
        { type: 'docHeading', attrs: { level: 1 }, content: [{ type: 'text', text: 'title' }] },
        { type: 'docTable', attrs: {}, content: [] },
      ],
    }
    const cells = cellContentOf(doc)
    expect(cells).toHaveLength(2)
    expect(cells[0]!.type).toBe('docParagraph')
    expect(cells[1]!.type).toBe('docParagraph')
    expect((cells[1]!.content![0] as { text?: string }).text).toBe('title')
  })

  it('an empty template degrades to one empty paragraph', () => {
    expect(cellContentOf({ type: 'doc', content: [] })).toEqual([emptyParagraph()])
  })
})
