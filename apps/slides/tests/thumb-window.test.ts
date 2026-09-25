// #763: the rail's windowing maths. Pure so the geometry can be checked without a browser.
import { describe, expect, it } from 'vitest'
import {
  DEFAULT_OVERSCAN_PX,
  THUMB_BORDER_PX,
  THUMB_GAP_PX,
  WINDOWING_MIN_SLIDES,
  insertionIndexAt,
  rowOffsets,
  thumbRowHeight,
  visibleRowRange,
} from '../src/renderer/thumb-window'

const wide = { widthPx: 1280, heightPx: 720 }

describe('thumbRowHeight', () => {
  it('matches the CSS box: scaled stage plus border and margin', () => {
    // 126 px wide → 70.9 px tall stage → 71 + 4 border + 10 margin
    expect(thumbRowHeight(wide, 126)).toBe(71 + THUMB_BORDER_PX + THUMB_GAP_PX)
  })

  it('follows the aspect ratio of the slide itself', () => {
    const tall = { widthPx: 720, heightPx: 1280 }
    expect(thumbRowHeight(tall, 126)).toBe(Math.round(126 * (1280 / 720)) + 14)
  })

  it('never collapses a slide to zero height', () => {
    expect(thumbRowHeight({ widthPx: 0, heightPx: 0 }, 126)).toBeGreaterThanOrEqual(1)
  })
})

describe('rowOffsets', () => {
  it('is a prefix sum ending in the total height', () => {
    expect(rowOffsets([10, 20, 30])).toEqual([0, 10, 30, 60])
  })

  it('handles an empty rail', () => {
    expect(rowOffsets([])).toEqual([0])
  })
})

describe('visibleRowRange', () => {
  const offsets = rowOffsets(new Array(400).fill(100)) // 400 rows of 100px

  it('mounts the viewport plus the overscan on either side', () => {
    const { start, end } = visibleRowRange(offsets, 10_000, 800)
    // visible band [10000-600, 10000+800+600) = [9400, 11400) → rows 94..113
    expect(start).toBe(94)
    expect(end).toBe(114)
    expect((end - start) * 100).toBeLessThan(800 + 2 * DEFAULT_OVERSCAN_PX + 200)
  })

  it('mounts a bounded slice of a long deck', () => {
    const { start, end } = visibleRowRange(offsets, 5_000, 800)
    expect(end - start).toBeLessThan(30)
    expect(offsets.length - 1).toBe(400)
  })

  it('starts at the top when the rail is not scrolled', () => {
    expect(visibleRowRange(offsets, 0, 800)).toEqual({ start: 0, end: 14 })
  })

  it('clamps at the end instead of running past it', () => {
    const { start, end } = visibleRowRange(offsets, 39_900, 800)
    expect(end).toBe(400)
    expect(start).toBeLessThan(400)
  })

  it('handles an empty rail and a zero-height viewport', () => {
    expect(visibleRowRange(rowOffsets([]), 0, 800)).toEqual({ start: 0, end: 0 })
    const { start, end } = visibleRowRange(offsets, 1_000, 0)
    expect(end).toBeGreaterThan(start)
  })
})

describe('insertionIndexAt', () => {
  const offsets = rowOffsets([100, 100, 100])

  it('returns the first row whose middle is below the pointer (same rule as the element scan)', () => {
    expect(insertionIndexAt(offsets, 10)).toBe(0) // above row 0's middle (50)
    expect(insertionIndexAt(offsets, 60)).toBe(1) // past row 0's middle
    expect(insertionIndexAt(offsets, 160)).toBe(2) // past row 1's middle (150)
    expect(insertionIndexAt(offsets, 260)).toBe(3) // past the last middle (250) → append
  })

  it('returns the end for a pointer past the last row', () => {
    expect(insertionIndexAt(offsets, 10_000)).toBe(3)
  })
})

describe('windowing threshold', () => {
  it('is high enough that short decks render as they always did', () => {
    expect(WINDOWING_MIN_SLIDES).toBeGreaterThanOrEqual(30)
    expect(WINDOWING_MIN_SLIDES).toBeLessThanOrEqual(120)
  })
})
