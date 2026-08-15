import { describe, expect, it } from 'vitest'
import type { PDFDocumentProxy } from 'pdfjs-dist'
import { buildSearchIndex, searchInIndex, type SearchIndex } from '../src/renderer/search'

interface FakeItem {
  str?: string
  transform?: number[]
  width?: number
  height?: number
  hasEOL?: boolean
}

function fakeDoc(pages: FakeItem[][]): PDFDocumentProxy {
  return {
    numPages: pages.length,
    getPage: async (n: number) => ({
      getTextContent: async () => ({ items: pages[n - 1] }),
    }),
  } as unknown as PDFDocumentProxy
}

const item = (str: string, x: number, y: number, w: number, h: number): FakeItem => ({
  str,
  transform: [1, 0, 0, 1, x, y],
  width: w,
  height: h,
})

describe('buildSearchIndex', () => {
  it('concatenates item text per page and records char ranges', async () => {
    const doc = fakeDoc([[item('Hello ', 10, 700, 60, 12), item('World', 70, 700, 50, 12)]])
    const index = await buildSearchIndex(doc)
    expect(index).toHaveLength(1)
    expect(index[0]!.text).toBe('Hello World')
    expect(index[0]!.lower).toBe('hello world')
    expect(index[0]!.items).toEqual([
      { start: 0, end: 6, x: 10, y: 700, w: 60, h: 12 },
      { start: 6, end: 11, x: 70, y: 700, w: 50, h: 12 },
    ])
  })

  it('inserts newlines for hasEOL and skips empty/invalid items', async () => {
    const doc = fakeDoc([
      [
        { ...item('line1', 0, 0, 10, 10), hasEOL: true },
        { str: '', hasEOL: true }, // empty text still contributes the EOL
        { transform: [1, 0, 0, 1, 0, 0] }, // no str -> skipped entirely
        item('line2', 0, 0, 10, 10),
      ],
    ])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.text).toBe('line1\n\nline2')
  })

  it('derives height from the transform when height is missing', async () => {
    const doc = fakeDoc([[{ str: 'x', transform: [1, 0, 3, 4, 0, 0], width: 5 }]])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.items[0]!.h).toBe(5) // hypot(3, 4)
  })

  it('flags rotated runs and leaves upright ones unflagged', async () => {
    const doc = fakeDoc([
      [
        item('upright', 10, 700, 60, 12),
        { str: 'rotated', transform: [0, 12, -12, 0, 200, 400], width: 60, height: 12 },
      ],
    ])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.items[0]!.rot).toBeUndefined()
    expect(index[0]!.items[1]!.rot).toBe(true)
  })

  it('synthetic italic shear (c ≠ 0, horizontal baseline) is not flagged as rotated', async () => {
    // ~12° shear as writers emit for fake italics: b = 0, c = tan(12°) × size
    const doc = fakeDoc([
      [{ str: 'emphasis', transform: [12, 0, 2.55, 12, 100, 700], width: 48, height: 12 }],
    ])
    const index = await buildSearchIndex(doc)
    expect(index[0]!.items[0]!.rot).toBeUndefined()
  })
})

describe('searchInIndex', () => {
  const entry = (text: string, items: SearchIndex[number]['items']): SearchIndex[number] => ({
    text,
    lower: text.toLowerCase(),
    items,
  })

  it('returns empty for an empty query', () => {
    const index = [entry('abc', [{ start: 0, end: 3, x: 0, y: 0, w: 30, h: 10 }])]
    expect(searchInIndex(index, '')).toEqual([])
  })

  it('finds case-insensitive matches with interpolated rects', () => {
    const index = [entry('Hello World', [{ start: 0, end: 11, x: 0, y: 700, w: 110, h: 12 }])]
    const matches = searchInIndex(index, 'WORLD')
    expect(matches).toHaveLength(1)
    expect(matches[0]!.pageIndex).toBe(0)
    // 'World' spans chars 6..11 of 11 -> x from 60 to 110
    expect(matches[0]!.rects).toHaveLength(1)
    const [x1, y1, x2, y2] = matches[0]!.rects[0]!
    expect(x1).toBeCloseTo(60)
    expect(y1).toBe(700)
    expect(x2).toBeCloseTo(110)
    expect(y2).toBe(712)
  })

  it('spans multiple items with one rect per item', () => {
    const index = [
      entry('abcdef', [
        { start: 0, end: 3, x: 0, y: 0, w: 30, h: 10 },
        { start: 3, end: 6, x: 30, y: 0, w: 30, h: 10 },
      ]),
    ]
    const matches = searchInIndex(index, 'cd')
    expect(matches).toHaveLength(1)
    expect(matches[0]!.rects).toHaveLength(2)
    expect(matches[0]!.rects[0]).toEqual([20, 0, 30, 10])
    expect(matches[0]!.rects[1]).toEqual([30, 0, 40, 10])
  })

  it('reports every occurrence and the correct page index', () => {
    const index = [
      entry('nothing here', [{ start: 0, end: 12, x: 0, y: 0, w: 120, h: 10 }]),
      entry('foo bar foo', [{ start: 0, end: 11, x: 0, y: 0, w: 110, h: 10 }]),
    ]
    const matches = searchInIndex(index, 'foo')
    expect(matches).toHaveLength(2)
    expect(matches.every((m) => m.pageIndex === 1)).toBe(true)
  })

  it('skips matches falling in EOL-only gaps with no item coverage', () => {
    // '\n' at chars 5..6 belongs to no item -> no rects -> match dropped
    const index = [
      entry('hello\nworld', [
        { start: 0, end: 5, x: 0, y: 0, w: 50, h: 10 },
        { start: 6, end: 11, x: 0, y: -20, w: 50, h: 10 },
      ]),
    ]
    expect(searchInIndex(index, 'hello')).toHaveLength(1)
    expect(searchInIndex(index, 'world')).toHaveLength(1)
    // The match itself spans the newline; rects come from both surrounding items
    expect(searchInIndex(index, 'hello\nworld')[0]!.rects).toHaveLength(2)
  })

  it('caps results at 1000 matches', () => {
    const text = 'a'.repeat(2000)
    const index = [entry(text, [{ start: 0, end: 2000, x: 0, y: 0, w: 2000, h: 10 }])]
    expect(searchInIndex(index, 'a')).toHaveLength(1000)
  })
})
