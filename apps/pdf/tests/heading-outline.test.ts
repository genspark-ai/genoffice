import { describe, expect, it } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import {
  buildHeadingOutline,
  collectLines,
  outlineFromLines,
  remapOutlinePages,
  type TextLine,
} from '../src/renderer/heading-outline'

interface FakeItem {
  str?: string
  transform?: number[]
  hasEOL?: boolean
}

const item = (str: string, size: number, x: number, y: number, hasEOL = false): FakeItem => ({
  str,
  transform: [size, 0, 0, size, x, y],
  hasEOL,
})

const line = (pageIndex: number, text: string, size: number, y: number): TextLine => ({
  pageIndex,
  text,
  size,
  y,
})

function bodyLines(pageIndex: number, from: number, count: number): TextLine[] {
  return Array.from({ length: count }, (_, i) =>
    line(pageIndex, `Body paragraph text number ${i} on page ${pageIndex}`, 10, from - i * 12),
  )
}

describe('collectLines', () => {
  it('joins items on one baseline and splits on baseline jumps or EOL', () => {
    const lines = collectLines(
      [
        item('Chapter ', 18, 72, 700),
        item('One', 18, 150, 700),
        item('Body text', 10, 72, 670, true),
        item('more body', 10, 72, 656),
      ],
      3,
    )
    expect(lines).toEqual([
      { pageIndex: 3, text: 'Chapter One', size: 18, y: 700 },
      { pageIndex: 3, text: 'Body text', size: 10, y: 670 },
      { pageIndex: 3, text: 'more body', size: 10, y: 656 },
    ])
  })

  it('takes the size most characters use and ignores rotated runs', () => {
    const lines = collectLines(
      [
        item('1', 14, 72, 700),
        item(' Introduction', 18, 80, 700),
        { str: 'DRAFT', transform: [0, 40, -40, 0, 300, 300] },
      ],
      0,
    )
    expect(lines).toEqual([{ pageIndex: 0, text: '1 Introduction', size: 18, y: 700 }])
  })
})

describe('outlineFromLines', () => {
  it('nests headings by size tier and points each at its page', () => {
    const lines = [
      line(0, 'User Guide', 24, 720),
      ...bodyLines(0, 690, 10),
      line(1, '1 Getting started', 16, 720),
      ...bodyLines(1, 690, 10),
      line(1, '1.1 Install', 13, 560),
      ...bodyLines(1, 540, 10),
      line(2, '2 Usage', 16, 720),
      ...bodyLines(2, 690, 10),
    ]
    const outline = outlineFromLines(lines)!
    expect(outline).toHaveLength(1)
    expect(outline[0]!.title).toBe('User Guide')
    expect(outline[0]!.dest).toEqual([0, { name: 'XYZ' }, null, 744, null])
    const chapters = outline[0]!.items!
    expect(chapters.map((c) => c.title)).toEqual(['1 Getting started', '2 Usage'])
    expect(chapters[0]!.items!.map((c) => c.title)).toEqual(['1.1 Install'])
    expect(chapters[1]!.dest).toEqual([2, { name: 'XYZ' }, null, 736, null])
  })

  it('merges a wrapped heading and drops running headers and numeric lines', () => {
    const lines = [
      line(0, 'ACME Corporation', 12, 780),
      line(0, 'A very long chapter title that', 16, 720),
      line(0, 'wraps onto a second line', 16, 700),
      ...bodyLines(0, 670, 10),
      line(1, 'ACME Corporation', 12, 780),
      line(1, '2', 16, 720),
      line(1, 'Second chapter', 16, 700),
      ...bodyLines(1, 670, 10),
      line(2, 'ACME Corporation', 12, 780),
      ...bodyLines(2, 690, 10),
    ]
    const outline = outlineFromLines(lines)!
    expect(outline.map((n) => n.title)).toEqual([
      'A very long chapter title that wraps onto a second line',
      'Second chapter',
    ])
  })

  it('rejects sentences, stat callouts and letter-spaced decoration', () => {
    const noise = [
      line(0, 'Since 2023 the region has been in a sustained conflict cycle.', 13, 700),
      line(0, '900 \u6b21', 20, 680),
      line(0, '110 \u4e07', 20, 660),
      line(0, 'A', 20, 640),
      line(0, 'M E R R I C K', 14, 620),
      line(0, '\u63a8 \u8350 \u6253 \u5361 \u5730 \uff1a \u8001 \u95e8 \u4e1c', 14, 600),
      line(0, '2024-2025 Industry Report', 18, 580),
      line(0, '\u76ee \u6b21', 16, 560),
    ]
    const outline = outlineFromLines([...noise, ...bodyLines(0, 540, 30)])!
    expect(outline.map((n) => n.title)).toEqual(['2024-2025 Industry Report'])
    expect(outline[0]!.items!.map((n) => n.title)).toEqual(['\u76ee \u6b21'])
  })

  it('gives up without a heading structure', () => {
    expect(outlineFromLines(bodyLines(0, 700, 20))).toBeNull()
    expect(outlineFromLines([line(0, 'Title only', 20, 700), ...bodyLines(0, 670, 5)])).toBeNull()
    // Mostly large text: the body-size guess is not trustworthy
    const table = Array.from({ length: 10 }, (_, i) => line(0, `Cell ${i} label`, 14, 700 - i * 20))
    expect(outlineFromLines([...table, ...bodyLines(0, 400, 4)])).toBeNull()
  })
})

describe('buildHeadingOutline', () => {
  it('reads every page and honors cancellation', async () => {
    const pages: FakeItem[][] = [
      [
        item('Report', 20, 72, 720, true),
        ...Array.from({ length: 8 }, (_, i) => item(`body ${i} text`, 10, 72, 690 - i * 12, true)),
      ],
      [
        item('Results', 20, 72, 720, true),
        ...Array.from({ length: 8 }, (_, i) => item(`body ${i} text`, 10, 72, 690 - i * 12, true)),
      ],
    ]
    const doc = {
      numPages: pages.length,
      getPage: async (n: number) => ({ getTextContent: async () => ({ items: pages[n - 1] }) }),
    } as unknown as PDFDocumentProxy
    const outline = await buildHeadingOutline(doc)
    expect(outline?.map((n) => n.title)).toEqual(['Report', 'Results'])
    expect(await buildHeadingOutline(doc, () => true)).toBeNull()
  })
})

describe('remapOutlinePages', () => {
  it('follows the save page map, hoists children of deleted pages, keeps bookmarks', () => {
    const nodes = [
      {
        title: 'A',
        dest: [0, { name: 'XYZ' }, null, 700, null],
        items: [{ title: 'A.1', dest: [1, { name: 'XYZ' }, null, 600, null] }],
      },
      { title: 'B', dest: [2, { name: 'XYZ' }, null, 700, null] },
      { title: 'Bookmark', dest: 'named-dest' },
    ]
    // Page 0 deleted, pages 1 and 2 swapped
    const out = remapOutlinePages(
      nodes,
      new Map([
        [1, 1],
        [2, 0],
      ]),
    )
    expect(out).toEqual([
      { title: 'A.1', dest: [1, { name: 'XYZ' }, null, 600, null] },
      { title: 'B', dest: [0, { name: 'XYZ' }, null, 700, null] },
      { title: 'Bookmark', dest: 'named-dest' },
    ])
  })
})
