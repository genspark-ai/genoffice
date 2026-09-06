// @vitest-environment jsdom
import { describe, expect, it } from 'vitest'
import { markTableSeamSlices, TABLE_SEAM_PX } from '../src/renderer/pagination-slices'
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

  it('shaves a whole CSS pixel, inside the table margin', () => {
    expect(TABLE_SEAM_PX).toBe(1)
  })
})
