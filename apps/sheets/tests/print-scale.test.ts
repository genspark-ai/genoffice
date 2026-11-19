import { describe, expect, it } from 'vitest'

import { countPages, fitToPageScale } from '../src/renderer/print-scale'

const A4_PORTRAIT = { printableWidthPt: 494.6, printableHeightPt: 733.5 }

function rows(count: number, heightPt = 15): number[] {
  return Array.from({ length: count }, () => heightPt)
}

describe('countPages', () => {
  it('packs whole rows onto pages and repeats the header on each', () => {
    // 10 rows of 100 on a 250 page with a 50 header: 2 rows fit per page.
    expect(countPages([{ repeatedHeightPt: 50, rowHeightsPt: rows(10, 100) }], 250)).toBe(5)
    expect(countPages([{ repeatedHeightPt: 0, rowHeightsPt: rows(10, 100) }], 250)).toBe(5)
    expect(countPages([{ repeatedHeightPt: 0, rowHeightsPt: rows(10, 100) }], 300)).toBe(4)
  })

  it('starts every area on a new page', () => {
    const area = { repeatedHeightPt: 0, rowHeightsPt: rows(2, 10) }
    expect(countPages([area, area, area], 1000)).toBe(3)
  })

  it('gives a row taller than the page its own page instead of looping', () => {
    expect(countPages([{ repeatedHeightPt: 0, rowHeightsPt: [500, 10, 500] }], 100)).toBe(3)
    expect(countPages([{ repeatedHeightPt: 0, rowHeightsPt: [] }], 100)).toBe(1)
  })
})

describe('fitToPageScale', () => {
  const tall = { repeatedHeightPt: 0, rowHeightsPt: rows(200) } // 3000pt of rows

  it('is 100% when both axes are automatic or already fit', () => {
    expect(
      fitToPageScale({
        ...A4_PORTRAIT,
        fitToWidth: 0,
        fitToHeight: 0,
        contentWidthPt: 5000,
        areas: [tall],
      }),
    ).toBe(1)
    expect(
      fitToPageScale({
        ...A4_PORTRAIT,
        fitToWidth: 1,
        fitToHeight: 1,
        contentWidthPt: 200,
        areas: [{ repeatedHeightPt: 0, rowHeightsPt: rows(10) }],
      }),
    ).toBe(1)
  })

  it('never enlarges past 100% (Excel fit only shrinks)', () => {
    expect(
      fitToPageScale({
        ...A4_PORTRAIT,
        fitToWidth: 3,
        fitToHeight: 3,
        contentWidthPt: 100,
        areas: [{ repeatedHeightPt: 0, rowHeightsPt: rows(2) }],
      }),
    ).toBe(1)
  })

  it('fits the width to N pages across', () => {
    const one = fitToPageScale({
      ...A4_PORTRAIT,
      fitToWidth: 1,
      fitToHeight: 0,
      contentWidthPt: 989.2,
      areas: [{ repeatedHeightPt: 0, rowHeightsPt: rows(3) }],
    })
    expect(one).toBeCloseTo(0.5, 5)
    const two = fitToPageScale({
      ...A4_PORTRAIT,
      fitToWidth: 2,
      fitToHeight: 0,
      contentWidthPt: 989.2,
      areas: [{ repeatedHeightPt: 0, rowHeightsPt: rows(3) }],
    })
    expect(two).toBe(1)
  })

  it('fits the height to N pages tall with whole rows', () => {
    const one = fitToPageScale({
      ...A4_PORTRAIT,
      fitToWidth: 0,
      fitToHeight: 1,
      contentWidthPt: 100,
      areas: [tall],
    })
    // 733.5 / 3000 = 0.2445 — the search settles at or just under it.
    expect(one).toBeLessThanOrEqual(0.2445)
    expect(one).toBeGreaterThan(0.23)
    expect(countPages([tall], A4_PORTRAIT.printableHeightPt / one)).toBe(1)
    const two = fitToPageScale({
      ...A4_PORTRAIT,
      fitToWidth: 0,
      fitToHeight: 2,
      contentWidthPt: 100,
      areas: [tall],
    })
    expect(two).toBeGreaterThan(one)
    expect(countPages([tall], A4_PORTRAIT.printableHeightPt / two)).toBeLessThanOrEqual(2)
  })

  it('steps down when pagination waste pushes a row onto an extra page', () => {
    // 5 rows of 100 under a 100 header on a 300 page: at 100% a page holds
    // the header plus two rows, so the sheet takes three pages.
    const area = { repeatedHeightPt: 100, rowHeightsPt: rows(5, 100) }
    const three = fitToPageScale({
      printableWidthPt: 1000,
      printableHeightPt: 300,
      fitToWidth: 0,
      fitToHeight: 3,
      contentWidthPt: 100,
      areas: [area],
    })
    expect(three).toBe(1)
    const two = fitToPageScale({
      printableWidthPt: 1000,
      printableHeightPt: 300,
      fitToWidth: 0,
      fitToHeight: 2,
      contentWidthPt: 100,
      areas: [area],
    })
    // Plain division (600 / 600 = 100%) still yields three pages because of
    // the repeated header; the search shrinks until header + 3 rows fit a
    // page (capacity ≥ 400pt → scale ≤ 75%).
    expect(two).toBeLessThanOrEqual(0.7500001)
    expect(two).toBeGreaterThan(0.7)
    expect(countPages([area], 300 / two)).toBe(2)
  })

  it('takes the smaller of the width and height scales', () => {
    const scale = fitToPageScale({
      ...A4_PORTRAIT,
      fitToWidth: 1,
      fitToHeight: 1,
      contentWidthPt: 989.2, // width alone → 0.5
      areas: [tall], // height alone → ~0.24
    })
    expect(scale).toBeLessThan(0.25)
    const wide = fitToPageScale({
      ...A4_PORTRAIT,
      fitToWidth: 1,
      fitToHeight: 1,
      contentWidthPt: 4946, // width alone → 0.1
      areas: [tall],
    })
    expect(wide).toBeCloseTo(0.1, 5)
  })

  it('budgets the pages per print area and takes the smallest fit', () => {
    // Three areas that each fit a page at 100%: fitToHeight=1 must not
    // count them as three pages and grind down to the floor.
    const small = { repeatedHeightPt: 0, rowHeightsPt: rows(10) }
    expect(
      fitToPageScale({
        ...A4_PORTRAIT,
        fitToWidth: 0,
        fitToHeight: 1,
        contentWidthPt: 100,
        areas: [small, small, small],
      }),
    ).toBe(1)
    // One tall area among small ones sets the scale; the small ones ride along.
    const mixed = fitToPageScale({
      ...A4_PORTRAIT,
      fitToWidth: 0,
      fitToHeight: 1,
      contentWidthPt: 100,
      areas: [small, tall, small],
    })
    expect(mixed).toBeLessThanOrEqual(0.2445)
    expect(mixed).toBeGreaterThan(0.23)
    expect(countPages([tall], A4_PORTRAIT.printableHeightPt / mixed)).toBe(1)
    // An empty area contributes nothing.
    expect(
      fitToPageScale({
        ...A4_PORTRAIT,
        fitToWidth: 0,
        fitToHeight: 1,
        contentWidthPt: 100,
        areas: [{ repeatedHeightPt: 0, rowHeightsPt: [] }, small],
      }),
    ).toBe(1)
  })

  it('clamps at the 10% floor for absurd content', () => {
    expect(
      fitToPageScale({
        ...A4_PORTRAIT,
        fitToWidth: 0,
        fitToHeight: 1,
        contentWidthPt: 100,
        areas: [{ repeatedHeightPt: 0, rowHeightsPt: rows(100_000) }],
      }),
    ).toBe(0.1)
  })
})
