import { describe, expect, it } from 'vitest'
import {
  layoutMultipage,
  multipageColumns,
  multipageSlotAt,
  multipageZoomForColumns,
  MULTIPAGE_GUTTER,
  MULTIPAGE_PANE_PAD,
} from '../src/renderer/multipage'

const LETTER_W = 816
const LETTER_H = 1056

describe('multipageColumns', () => {
  it('stays single-column at 100% in a pane narrower than two pages', () => {
    expect(multipageColumns(1400, LETTER_W, 100)).toBe(1)
  })

  it('flows two pages side by side once the zoom lets them fit', () => {
    expect(multipageColumns(1400, LETTER_W, 50)).toBe(3)
    expect(multipageColumns(1400, LETTER_W, 80)).toBe(2)
  })

  it('counts the markup column in the slot width', () => {
    expect(multipageColumns(1400, LETTER_W + 200, 80)).toBe(1)
  })

  it('never returns less than one column', () => {
    expect(multipageColumns(300, LETTER_W, 500)).toBe(1)
    expect(multipageColumns(0, LETTER_W, 100)).toBe(1)
  })
})

describe('multipageZoomForColumns', () => {
  it('round-trips with multipageColumns', () => {
    const paneW = 1400
    for (const cols of [2, 3, 4]) {
      const z = Math.floor(multipageZoomForColumns(paneW, LETTER_W, cols))
      expect(multipageColumns(paneW, LETTER_W, z)).toBe(cols)
      expect(multipageColumns(paneW, LETTER_W, z + 1)).toBeLessThanOrEqual(cols)
    }
  })

  it('one column is the width fit', () => {
    const z = multipageZoomForColumns(1400, LETTER_W, 1)
    expect(z).toBeCloseTo(((1400 - MULTIPAGE_PANE_PAD) / LETTER_W) * 100)
  })
})

describe('layoutMultipage', () => {
  it('places pages row-major with a gutter between columns and rows', () => {
    const l = layoutMultipage([LETTER_H, LETTER_H, LETTER_H], 2, LETTER_W)
    expect(l.cols).toBe(2)
    expect(l.slots).toEqual([
      { left: 0, top: 0, width: LETTER_W, height: LETTER_H },
      { left: LETTER_W + MULTIPAGE_GUTTER, top: 0, width: LETTER_W, height: LETTER_H },
      { left: 0, top: LETTER_H + MULTIPAGE_GUTTER, width: LETTER_W, height: LETTER_H },
    ])
    expect(l.width).toBe(2 * LETTER_W + MULTIPAGE_GUTTER)
    expect(l.height).toBe(2 * LETTER_H + MULTIPAGE_GUTTER)
  })

  it('a row takes its tallest page; a short final row does not widen the grid', () => {
    const l = layoutMultipage([LETTER_H, 1400, LETTER_H], 3, LETTER_W)
    expect(l.height).toBe(1400)
    expect(l.width).toBe(3 * LETTER_W + 2 * MULTIPAGE_GUTTER)
    const one = layoutMultipage([LETTER_H], 3, LETTER_W)
    expect(one.width).toBe(LETTER_W)
  })

  it('handles an empty document', () => {
    const l = layoutMultipage([], 2, LETTER_W)
    expect(l.slots).toEqual([])
    expect(l.width).toBe(0)
    expect(l.height).toBe(0)
  })
})

describe('multipageSlotAt', () => {
  const l = layoutMultipage([LETTER_H, LETTER_H, LETTER_H, LETTER_H], 2, LETTER_W)

  it('maps points inside a slot and in the gutter after it', () => {
    expect(multipageSlotAt(l, 10, 10)).toBe(0)
    expect(multipageSlotAt(l, LETTER_W + 5, 10)).toBe(0)
    expect(multipageSlotAt(l, LETTER_W + MULTIPAGE_GUTTER + 5, 10)).toBe(1)
    expect(multipageSlotAt(l, 10, LETTER_H + MULTIPAGE_GUTTER + 5)).toBe(2)
  })

  it('a point right of the last column lands on that row’s last slot', () => {
    expect(multipageSlotAt(l, 5000, 10)).toBe(1)
  })
})
