// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { markTableSeamSlices, seamWindow, TABLE_SEAM_PX } from '../src/renderer/pagination-slices'
import type { BlockBox, PageSlice } from '../src/renderer/pagination-types'

const block = (tag: string, top: number, height: number): BlockBox => ({
  el: document.createElement(tag),
  top,
  height,
})

describe('markTableSeamSlices', () => {
  it('flags pages that open with a native table, never the first page', () => {
    const blocks = [block('p', 0, 200), block('table', 200.4, 300), block('p', 500.4, 100)]
    const slices: PageSlice[] = [
      { start: 0, end: 200.4, section: 0 },
      { start: 200.4, end: 500.4, section: 0 },
      { start: 500.4, end: 600.4, section: 0 },
    ]
    markTableSeamSlices(slices, blocks)
    expect(slices.map((s) => s.leadTable ?? false)).toEqual([false, true, false])
  })

  it('tolerates sub-pixel drift between the slice start and the table top', () => {
    const blocks = [block('p', 0, 100), block('table', 100.3, 50)]
    const slices: PageSlice[] = [
      { start: 0, end: 100, section: 0 },
      { start: 100, end: 150.3, section: 0 },
    ]
    markTableSeamSlices(slices, blocks)
    expect(slices[1].leadTable).toBe(true)
  })

  it('ignores mid-block cuts, non-table leads and column-region pages', () => {
    const blocks = [block('table', 0, 400), block('table', 400, 100), block('div', 500, 100)]
    const slices: PageSlice[] = [
      { start: 0, end: 250, section: 0 },
      { start: 250, end: 400, section: 0 },
      { start: 400, end: 500, section: 0, regions: [] },
      { start: 500, end: 600, section: 0 },
    ]
    markTableSeamSlices(slices, blocks)
    expect(slices.some((s) => s.leadTable)).toBe(false)
  })

  it('flags pages that open inside a table cut by the break, never its lead page', () => {
    const blocks = [block('p', 0, 100), block('table', 100, 500), block('p', 600, 100)]
    const slices: PageSlice[] = [
      { start: 0, end: 300, section: 0 },
      { start: 300, end: 600, section: 0 },
      { start: 600, end: 700, section: 0 },
    ]
    markTableSeamSlices(slices, blocks)
    expect(slices.map((s) => s.cutTable ?? false)).toEqual([false, true, false])
    expect(slices.some((s) => s.leadTable)).toBe(false)
  })

  it('does not flag a break that lands on the table bottom edge or in a column page', () => {
    const blocks = [block('table', 0, 300), block('p', 300, 100), block('table', 400, 300)]
    const slices: PageSlice[] = [
      { start: 0, end: 300, section: 0 },
      { start: 300, end: 550, section: 0 },
      { start: 550, end: 700, section: 0, regions: [] },
    ]
    markTableSeamSlices(slices, blocks)
    expect(slices.some((s) => s.cutTable)).toBe(false)
  })

  it('flags a page opening on the bottom edge of the previous table', () => {
    const blocks = [block('p', 0, 100), block('table', 100, 500), block('p', 600, 100)]
    const slices: PageSlice[] = [
      { start: 0, end: 600, section: 0 },
      { start: 600, end: 700, section: 0 },
    ]
    markTableSeamSlices(slices, blocks)
    expect(slices.map((s) => s.tailTable ?? false)).toEqual([false, true])
    expect(slices.some((s) => s.cutTable || s.leadTable)).toBe(false)
  })

  it('reads the table box bottom under a folded inter-block gap', () => {
    const gapped = { ...block('table', 100, 508), spaceAfterPx: 8 }
    const blocks = [block('p', 0, 100), gapped, block('p', 608, 100)]
    const slices: PageSlice[] = [
      { start: 0, end: 608, section: 0 },
      { start: 608, end: 708, section: 0 },
    ]
    markTableSeamSlices(slices, blocks)
    expect(slices[1].tailTable).toBeUndefined()
    expect(slices[1].cutTable).toBeUndefined()
  })

  it('prefers the lead flag when a table opens right under the previous one', () => {
    const blocks = [block('table', 0, 300), block('table', 300, 300)]
    const slices: PageSlice[] = [
      { start: 0, end: 300, section: 0 },
      { start: 300, end: 600, section: 0 },
    ]
    markTableSeamSlices(slices, blocks)
    expect(slices[1].leadTable).toBe(true)
    expect(slices[1].tailTable).toBeUndefined()
  })

  it('shaves a whole CSS pixel, inside the table margin', () => {
    expect(TABLE_SEAM_PX).toBe(1)
  })
})

describe('seamWindow', () => {
  const s = (extra: Partial<PageSlice> = {}): PageSlice => ({
    start: 0,
    end: 100,
    section: 0,
    ...extra,
  })

  it('opens a cut page up and grows the page before it', () => {
    expect(seamWindow(s({ cutTable: true }), undefined)).toEqual({ lift: 1, extend: 0 })
    expect(seamWindow(s(), s({ cutTable: true }))).toEqual({ lift: 0, extend: 1 })
  })

  it('moves the pixel from a tail page to the page before it', () => {
    expect(seamWindow(s({ tailTable: true }), undefined)).toEqual({ lift: -1, extend: 0 })
    expect(seamWindow(s(), s({ tailTable: true }))).toEqual({ lift: 0, extend: 1 })
  })

  it('shaves the page before a lead table and leaves lifted tail pages alone', () => {
    expect(seamWindow(s(), s({ leadTable: true }))).toEqual({ lift: 0, extend: -1 })
    expect(seamWindow(s({ tailTable: true, liftTop: 4 }), undefined)).toEqual({
      lift: 0,
      extend: 0,
    })
  })
})
