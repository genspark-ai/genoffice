import { describe, expect, it } from 'vitest'
import type { PmNode } from '../src/renderer/editor/convert'
import {
  AVERY_5160,
  cellContentOf,
  clampGridCount,
  customGridLayout,
  emptyParagraph,
  envelopeLayout,
  labelCellSize,
  labelSheetDoc,
  layoutSection,
  MAX_CUSTOM_GRID,
  unsupportedLabelBlocks,
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

  it('clamps a typed custom grid count instead of emitting pixel-sized cells', () => {
    expect(customGridLayout('letter', 500, 0).grid).toMatchObject({
      columns: MAX_CUSTOM_GRID,
      rows: 1,
    })
    expect(customGridLayout('letter', Number.NaN, 7).grid).toMatchObject({ columns: 1, rows: 7 })
    expect(clampGridCount(-3)).toBe(1)
    expect(clampGridCount(2.9)).toBe(2)
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

  it('fills the grid row-major, with gutter cells, and pads the last row', () => {
    const cells = Array.from({ length: 25 }, (_, i) => [paragraph(`r${i + 1}`)])
    const doc = labelSheetDoc(cells, layout)
    const table = doc.content![0]!
    const rows = table.content!
    // 25 records on a 3-wide grid: 8 full rows and one with 1 record + 2 empty
    expect(rows).toHaveLength(9)
    // label, gutter, label, gutter, label
    expect(rows[0]!.content).toHaveLength(5)
    const first = rows[0]!.content!
    expect(cellText(first[0]!)).toBe('r1')
    expect(cellText(first[1]!)).toBeUndefined()
    expect(cellText(first[2]!)).toBe('r2')
    expect(cellText(first[4]!)).toBe('r3')
    expect(first[1]!.attrs?.colwidth).toEqual([12])
    const last = rows[8]!.content!
    expect(cellText(last[0]!)).toBe('r25')
    expect(cellText(last[2]!)).toBeUndefined()
    expect(last[2]!.content).toEqual([emptyParagraph()])
    expect(last[4]!.content).toEqual([emptyParagraph()])
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
    expect(row.content![1]!.attrs?.colwidth).toEqual([Math.round(layout.grid!.gapX / 15)])
  })

  it('is borderless and puts a paragraph after the table', () => {
    const doc = labelSheetDoc([[paragraph('a')]], layout)
    expect(doc.content).toHaveLength(2)
    expect(doc.content![1]).toEqual(emptyParagraph())
    const table = doc.content![0]!
    const cells = table.content!.flatMap((row) => row.content ?? [])
    expect(cells.length).toBeGreaterThan(0)
    for (const cell of cells) {
      expect(cell.attrs?.borders).toMatchObject({
        top: { style: 'none' },
        left: { style: 'none' },
        bottom: { style: 'none' },
        right: { style: 'none' },
      })
    }
  })

  it('a stock with a vertical gap gets a spanning gutter row between label rows', () => {
    const spaced = { ...layout, grid: { ...layout.grid!, gapY: 180 } }
    const cells = Array.from({ length: 4 }, (_, i) => [paragraph(`r${i + 1}`)])
    const doc = labelSheetDoc(cells, spaced)
    const rows = doc.content![0]!.content!
    // two label rows, one gutter row between them
    expect(rows).toHaveLength(3)
    expect(rows[1]!.attrs?.heightTwips).toBe(180)
    expect(rows[1]!.attrs?.heightRule).toBe('exact')
    expect(rows[1]!.content).toHaveLength(1)
    expect(rows[1]!.content![0]!.attrs?.colspan).toBe(5)
    expect(rows[1]!.content![0]!.attrs?.borders).toMatchObject({ top: { style: 'none' } })
  })

  it('a single record still gets a full first row', () => {
    const doc = labelSheetDoc([[paragraph('only')]], layout)
    expect(doc.content![0]!.content![0]!.content).toHaveLength(5)
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

  it('counts the blocks a label cell cannot carry', () => {
    const doc: PmNode = {
      type: 'doc',
      content: [
        paragraph('name'),
        { type: 'docHeading', attrs: { level: 1 } },
        { type: 'docTable', attrs: {}, content: [] },
        { type: 'docProtected', attrs: {} },
      ],
    }
    expect(unsupportedLabelBlocks(doc)).toBe(2)
    expect(unsupportedLabelBlocks({ type: 'doc', content: [paragraph('x')] })).toBe(0)
  })
})
