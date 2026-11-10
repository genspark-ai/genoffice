import { describe, expect, it } from 'vitest'
import { BooleanNumber } from '@univerjs/core'
import {
  type GridInsetSource,
  gridTopInset,
  shouldShowVisualDeleteButton,
  visualDeleteButtonPosition,
} from '../src/renderer/visual-delete-button'

// Grid area below the column header: 20px header on a canvas starting at y=100.
const bounds = { left: 50, top: 120, right: 850, bottom: 620 }
const size = 20
const gap = 4

describe('visual delete button visibility', () => {
  it('is absent for an unselected visual (hover does not count)', () => {
    expect(shouldShowVisualDeleteButton({ selected: false, textEditing: false })).toBe(false)
  })

  it('is present for the selected visual', () => {
    expect(shouldShowVisualDeleteButton({ selected: true, textEditing: false })).toBe(true)
  })

  it('stays hidden while the selected shape is in text edit', () => {
    expect(shouldShowVisualDeleteButton({ selected: true, textEditing: true })).toBe(false)
  })
})

describe('visual delete button placement', () => {
  it('sits above the top-right corner, outside the frame', () => {
    const frame = { left: 300, top: 300, right: 324, bottom: 324 }
    const position = visualDeleteButtonPosition(frame, bounds, size, gap)
    expect(position).toEqual({ left: 304, top: 276, placement: 'above' })
    // Never overlaps the visual, even one smaller than the button itself.
    expect(position.top + size).toBeLessThanOrEqual(frame.top)
  })

  it('flips below when the visual touches the top of the grid', () => {
    const frame = { left: 300, top: 120, right: 400, bottom: 180 }
    expect(visualDeleteButtonPosition(frame, bounds, size, gap)).toEqual({
      left: 380,
      top: 184,
      placement: 'below',
    })
  })

  it('flips below when only part of the gap is available above', () => {
    const frame = { left: 300, top: 130, right: 400, bottom: 180 }
    expect(visualDeleteButtonPosition(frame, bounds, size, gap).placement).toBe('below')
  })

  it('falls back inside the visible top-right corner when neither side fits', () => {
    const frame = { left: 300, top: 120, right: 400, bottom: 700 }
    expect(visualDeleteButtonPosition(frame, bounds, size, gap)).toEqual({
      left: 380,
      top: 124,
      placement: 'inside',
    })
  })

  it('uses the visible edge of a frame scrolled under the header for the inside fallback', () => {
    const frame = { left: 300, top: 60, right: 400, bottom: 700 }
    expect(visualDeleteButtonPosition(frame, bounds, size, gap).top).toBe(bounds.top + gap)
  })

  it('clamps horizontally so a visual at the right edge keeps a reachable button', () => {
    const frame = { left: 800, top: 300, right: 900, bottom: 360 }
    expect(visualDeleteButtonPosition(frame, bounds, size, gap).left).toBe(bounds.right - size)
    const sliver = { left: 20, top: 300, right: 40, bottom: 360 }
    expect(visualDeleteButtonPosition(sliver, bounds, size, gap).left).toBe(bounds.left)
  })
})

function fakeWorksheet(options: {
  zoom?: number
  headerHidden?: boolean
  freeze?: { ySplit: number; startRow: number }
  hiddenRows?: readonly number[]
}): GridInsetSource {
  const hidden = new Set(options.hiddenRows ?? [])
  return {
    getZoom: () => options.zoom ?? 1,
    getFreeze: () => options.freeze ?? { ySplit: 0, startRow: 0 },
    getRowHeight: (row) => 20 + row,
    getSheet: () => ({
      getConfig: () => ({
        columnHeader: {
          height: 20,
          hidden: options.headerHidden ? BooleanNumber.TRUE : BooleanNumber.FALSE,
        },
      }),
      getRowVisible: (row) => !hidden.has(row),
    }),
  }
}

describe('grid top inset', () => {
  it('is the column header height by default', () => {
    expect(gridTopInset(fakeWorksheet({}))).toBe(20)
  })

  it('drops to zero with headings hidden and scales with zoom', () => {
    expect(gridTopInset(fakeWorksheet({ headerHidden: true }))).toBe(0)
    expect(gridTopInset(fakeWorksheet({ zoom: 1.5 }))).toBe(30)
  })

  it('adds the visible frozen rows below the header', () => {
    // Rows 0-2 frozen: 20 + 21 + 22 on top of the 20px header.
    expect(gridTopInset(fakeWorksheet({ freeze: { ySplit: 3, startRow: 3 } }))).toBe(83)
    expect(
      gridTopInset(fakeWorksheet({ freeze: { ySplit: 3, startRow: 3 }, hiddenRows: [1] })),
    ).toBe(62)
  })
})
