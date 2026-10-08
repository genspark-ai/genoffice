import { describe, expect, it } from 'vitest'

import { bandIndices, planPrintPages, type PaginationInput } from '../src/renderer/print-paginate'

function input(overrides: Partial<PaginationInput> = {}): PaginationInput {
  return {
    areas: [{ startRow: 0, endRow: 9, startColumn: 0, endColumn: 3 }],
    rowHeightPt: () => 15,
    columnWidthPt: () => 60,
    titleRows: null,
    titleColumns: null,
    rowBreaks: [],
    colBreaks: [],
    pageWidthPt: 150,
    pageHeightPt: 50,
    headingWidthPt: 0,
    headingHeightPt: 0,
    ...overrides,
  }
}

describe('bandIndices', () => {
  it('packs whole items and never splits one', () => {
    expect(bandIndices([0, 1, 2, 3, 4], () => 15, 50, new Set())).toEqual([
      { start: 0, end: 2 },
      { start: 3, end: 4 },
    ])
  })

  it('gives an oversized item a band of its own', () => {
    expect(bandIndices([0, 1, 2], (i) => (i === 1 ? 100 : 10), 50, new Set())).toEqual([
      { start: 0, end: 0 },
      { start: 1, end: 1 },
      { start: 2, end: 2 },
    ])
  })

  it('ends the band before a manual break', () => {
    expect(bandIndices([0, 1, 2, 3], () => 1, 100, new Set([2]))).toEqual([
      { start: 0, end: 1 },
      { start: 2, end: 3 },
    ])
    // a break on the first item does not open an empty band
    expect(bandIndices([0, 1], () => 1, 100, new Set([0]))).toEqual([{ start: 0, end: 1 }])
  })
})

describe('planPrintPages', () => {
  it('pages down then over by default', () => {
    // 10 rows of 15pt on 50pt pages = 4 row bands; 4 columns of 60pt on 150pt = 2 column bands.
    const pages = planPrintPages(input())
    expect(pages).toHaveLength(8)
    expect(pages.slice(0, 4).every((page) => page.colStart === 0 && page.colEnd === 1)).toBe(true)
    expect(pages[0]).toEqual({ area: 0, rowStart: 0, rowEnd: 2, colStart: 0, colEnd: 1 })
    expect(pages[3]).toEqual({ area: 0, rowStart: 9, rowEnd: 9, colStart: 0, colEnd: 1 })
    expect(pages[4]).toEqual({ area: 0, rowStart: 0, rowEnd: 2, colStart: 2, colEnd: 3 })
  })

  it('pages over then down when asked', () => {
    const pages = planPrintPages(input({ order: 'overThenDown' }))
    expect(pages[1]).toEqual({ area: 0, rowStart: 0, rowEnd: 2, colStart: 2, colEnd: 3 })
  })

  it('reserves the heading strip and the repeated titles on every page', () => {
    // 50 - 10 heading - 15 title row = 25pt for body rows: one 15pt row per page.
    const pages = planPrintPages(
      input({
        titleRows: { start: 0, end: 0 },
        headingHeightPt: 10,
        headingWidthPt: 30,
        titleColumns: { start: 0, end: 0 },
        pageWidthPt: 150,
      }),
    )
    // body rows 1..9 → 9 bands; body columns 1..3 fit (150 - 30 - 60 = 60) one per band → 3.
    expect(pages).toHaveLength(27)
    expect(pages[0]).toEqual({ area: 0, rowStart: 1, rowEnd: 1, colStart: 1, colEnd: 1 })
  })

  it('skips hidden rows and columns and starts every area on a new page', () => {
    const pages = planPrintPages(
      input({
        areas: [
          { startRow: 0, endRow: 1, startColumn: 0, endColumn: 0 },
          { startRow: 5, endRow: 6, startColumn: 0, endColumn: 1 },
        ],
        rowHeightPt: (row) => (row === 5 ? 0 : 15),
        columnWidthPt: (column) => (column === 1 ? 0 : 60),
      }),
    )
    expect(pages).toEqual([
      { area: 0, rowStart: 0, rowEnd: 1, colStart: 0, colEnd: 0 },
      { area: 1, rowStart: 6, rowEnd: 6, colStart: 0, colEnd: 0 },
    ])
  })

  it('drops an area without visible body cells', () => {
    expect(
      planPrintPages(
        input({
          areas: [{ startRow: 0, endRow: 0, startColumn: 0, endColumn: 0 }],
          titleRows: { start: 0, end: 0 },
        }),
      ),
    ).toEqual([])
  })

  it('honours manual row and column breaks', () => {
    const pages = planPrintPages(
      input({ pageHeightPt: 1000, pageWidthPt: 1000, rowBreaks: [4], colBreaks: [1, 3] }),
    )
    expect(pages.map((page) => [page.rowStart, page.rowEnd, page.colStart, page.colEnd])).toEqual([
      [0, 3, 0, 0],
      [4, 9, 0, 0],
      [0, 3, 1, 2],
      [4, 9, 1, 2],
      [0, 3, 3, 3],
      [4, 9, 3, 3],
    ])
  })
})
