import { describe, expect, it } from 'vitest'

import {
  mirrorSpanX,
  RTL_BORDER_TYPE_SWAP,
  rtlFreezeLayout,
  rtlFreezeXSplit,
  rtlHorizontalScrollState,
  rtlMaxViewportScrollX,
  rtlRevealScrollX,
} from '../src/renderer/rtl-grid-mirror'
import {
  EMU_PER_PIXEL,
  markerFrom,
  mirrorCornerX,
  walkMarker,
} from '../src/renderer/WorkbookVisuals'

describe('mirrorSpanX', () => {
  it('mirrors a grid-space span around the total width', () => {
    expect(mirrorSpanX(10, 30, 100)).toEqual({ startX: 70, endX: 90 })
    expect(mirrorSpanX(0, 100, 100)).toEqual({ startX: 0, endX: 100 })
  })

  it('keeps the span width and ordering', () => {
    const { startX, endX } = mirrorSpanX(37.5, 42.25, 512)
    expect(endX - startX).toBeCloseTo(4.75)
    expect(startX).toBeLessThan(endX)
  })

  it('mirrors inside the grid when coordinates carry the header offset', () => {
    // Grid [0, 100] shifted by a 46px row header: span [56, 76] → [70+46, 90+46].
    expect(mirrorSpanX(56, 76, 100, 46)).toEqual({ startX: 116, endX: 136 })
  })

  it('round-trips', () => {
    const once = mirrorSpanX(12, 20, 88, 46)
    const twice = mirrorSpanX(once.startX, once.endX, 88, 46)
    expect(twice).toEqual({ startX: 12, endX: 20 })
  })
})

describe('rtlFreezeLayout', () => {
  // Real geometry from a production Hebrew workbook (xSplit=2 → columns A+B
  // frozen, widths 115+246): engine 1240px, header strip 46px, grid 3244px.
  const base = {
    engineWidth: 1240,
    scaleX: 1,
    headerWidth: 46,
    totalWidth: 3244,
    freezeStartX: 0,
    freezeEndX: 361,
  }

  it('docks the frozen pane against the right-edge header strip', () => {
    const { stripLeft, freezeGap, paneLeft } = rtlFreezeLayout(base)
    expect(stripLeft).toBe(1240 - 46 - 1)
    expect(freezeGap).toBe(361)
    expect(paneLeft).toBe(stripLeft - 361)
  })

  it('pins the mirrored frozen band exactly onto the dock', () => {
    const { paneLeft, frozenScrollX } = rtlFreezeLayout(base)
    // Scene x rendered at the pane's left edge must be the mirrored band
    // start: headerWidth + totalWidth - freezeEndX.
    expect(paneLeft + frozenScrollX).toBeCloseTo(base.headerWidth + base.totalWidth - 361)
  })

  it('reserves the frozen band inside the main pane inset', () => {
    const { mainRightInset } = rtlFreezeLayout(base)
    expect(mainRightInset).toBe((46 + 1 + 361) * 1)
  })

  it('degenerates to the plain RTL anchor without a column freeze', () => {
    const layout = rtlFreezeLayout({ ...base, freezeStartX: 0, freezeEndX: 0 })
    expect(layout.freezeGap).toBe(0)
    expect(layout.paneLeft).toBe(layout.stripLeft)
    expect(layout.mainRightInset).toBe((46 + 1) * 1)
  })

  it('never produces a negative gap (mirrored inputs would)', () => {
    // Regression guard: feeding scene-space (mirrored) coords made the gap
    // negative (start 3129 > end 2883) and collapsed the freeze viewports.
    const layout = rtlFreezeLayout({ ...base, freezeStartX: 3129, freezeEndX: 2883 })
    expect(layout.freezeGap).toBe(0)
  })

  it('handles a scrolled freeze anchor (freezeStartX > 0)', () => {
    const layout = rtlFreezeLayout({ ...base, freezeStartX: 115, freezeEndX: 361 })
    expect(layout.freezeGap).toBe(246)
    expect(layout.paneLeft + layout.frozenScrollX).toBeCloseTo(46 + 3244 - 361)
  })
})

describe('rtlMaxViewportScrollX', () => {
  it('meets the frozen band seamlessly at max scroll', () => {
    // Scene = header + grid; main pane width = engine - header - 1 - gap.
    const sceneWidth = 46 + 3244
    const gap = 361
    const paneWidth = 1240 - 46 - 1 - gap
    const max = rtlMaxViewportScrollX(sceneWidth, paneWidth, 1, 0, gap)
    // Right edge of the main pane at max scroll = mirrored frozen band start.
    expect(max + paneWidth).toBeCloseTo(46 + 3244 - gap)
  })

  it('matches the old cap when there is no freeze padding', () => {
    expect(rtlMaxViewportScrollX(3290, 1193, 1)).toBeCloseTo(3290 - 1193)
  })

  it('ignores an inverted padding window', () => {
    expect(rtlMaxViewportScrollX(3290, 1193, 1, 3129, 2883)).toBeCloseTo(3290 - 1193)
  })

  it('divides the viewport width by the zoom scale', () => {
    expect(rtlMaxViewportScrollX(3290, 1000, 2, 0, 100)).toBeCloseTo(3290 - 100 - 500)
  })

  it('keeps the columns hidden behind a mid-sheet band out of the pane', () => {
    // 60 x 74px columns behind a 46px header; freeze dropped while scrolled so
    // that column K (index 10) is the band: ten columns (740px) are hidden
    // behind it. RTL padding = [0, hidden + band] (header-less LTR x).
    const sceneWidth = 46 + 60 * 74
    const paneWidth = 1119
    const max = rtlMaxViewportScrollX(sceneWidth, paneWidth, 1, 0, 740 + 74)
    // pane's right edge at home = mirrored start of column K
    expect(max + paneWidth).toBe(46 + 60 * 74 - 11 * 74)
    // the stock LTR window [740, 814] would have overshot by the hidden width
    expect(rtlMaxViewportScrollX(sceneWidth, paneWidth, 1, 740, 814) - max).toBe(740)
  })
})

describe('rtlHorizontalScrollState', () => {
  // 10 columns of 100px in a 400px pane: scene x grows toward the visual
  // right, home (column A flush right) is the max scroll 600.
  const bounds = { minScrollX: 0, maxScrollX: 600 }
  const offsetRelativeToColumn = (scrollX: number) => ({
    column: Math.floor(scrollX / 100),
    columnOffset: scrollX % 100,
  })

  it('records a mid-sheet pixel as its column + offset', () => {
    expect(rtlHorizontalScrollState(250, bounds, offsetRelativeToColumn)).toEqual({
      sheetViewStartColumn: 2,
      offsetX: 50,
    })
  })

  it('maps the home edge and anything past it to the column 0 / offset 0 sentinel', () => {
    expect(rtlHorizontalScrollState(600, bounds, offsetRelativeToColumn)).toEqual({
      sheetViewStartColumn: 0,
      offsetX: 0,
    })
    expect(rtlHorizontalScrollState(9999, bounds, offsetRelativeToColumn)).toEqual({
      sheetViewStartColumn: 0,
      offsetX: 0,
    })
  })

  it('clamps a wheel overshoot past the far (max-column) edge', () => {
    expect(
      rtlHorizontalScrollState(-80, { minScrollX: 20, maxScrollX: 600 }, offsetRelativeToColumn),
    ).toEqual({
      sheetViewStartColumn: 0,
      offsetX: 20,
    })
  })

  it('treats a sheet narrower than the pane as always home', () => {
    expect(
      rtlHorizontalScrollState(120, { minScrollX: 0, maxScrollX: 0 }, offsetRelativeToColumn),
    ).toEqual({
      sheetViewStartColumn: 0,
      offsetX: 0,
    })
  })
})

describe('rtlRevealScrollX', () => {
  const view = { scrollX: 1000, paneWidth: 400 }

  it('leaves a fully visible column alone', () => {
    expect(rtlRevealScrollX({ startX: 1100, endX: 1200 }, view)).toBeNull()
    expect(rtlRevealScrollX({ startX: 1000, endX: 1400 }, view)).toBeNull()
  })

  it('scrolls the least distance for a column hidden toward the visual left (higher index)', () => {
    // flush with the pane's left edge
    expect(rtlRevealScrollX({ startX: 700, endX: 800 }, view)).toBe(700)
    // partially clipped on the left counts too
    expect(rtlRevealScrollX({ startX: 950, endX: 1050 }, view)).toBe(950)
  })

  it('scrolls the least distance for a column hidden toward the visual right (lower index)', () => {
    // flush with the pane's right edge
    expect(rtlRevealScrollX({ startX: 1600, endX: 1700 }, view)).toBe(1300)
    expect(rtlRevealScrollX({ startX: 1350, endX: 1450 }, view)).toBe(1050)
  })
})

describe('rtlFreezeXSplit', () => {
  const oldFreeze = { startColumn: 2, xSplit: 2 }

  it('at home the split is the dropped column itself', () => {
    expect(
      rtlFreezeXSplit({ column: 5, lowestVisibleColumn: 2, oldFreeze, droppedInFrozenBand: false }),
    ).toBe(5)
  })

  it('anchors a scrolled pane at its lowest visible column (not the visual-left max column)', () => {
    // pane scrolled to show columns 10.. (lowest, at the visual right) — drop on 13
    expect(
      rtlFreezeXSplit({
        column: 13,
        lowestVisibleColumn: 10,
        oldFreeze,
        droppedInFrozenBand: false,
      }),
    ).toBe(5)
  })

  it('anchors a drop inside the frozen band at the first frozen column, like LTR', () => {
    expect(
      rtlFreezeXSplit({ column: 1, lowestVisibleColumn: 10, oldFreeze, droppedInFrozenBand: true }),
    ).toBe(1)
  })

  it('never returns a negative split', () => {
    expect(
      rtlFreezeXSplit({
        column: 3,
        lowestVisibleColumn: 10,
        oldFreeze,
        droppedInFrozenBand: false,
      }),
    ).toBe(0)
  })
})

describe('RTL_BORDER_TYPE_SWAP', () => {
  it('is an involution (swapping twice restores every type)', () => {
    for (const [from, to] of Object.entries(RTL_BORDER_TYPE_SWAP)) {
      expect(RTL_BORDER_TYPE_SWAP[to]).toBe(from)
    }
  })

  it('maps left to right and mirrors diagonals horizontally', () => {
    expect(RTL_BORDER_TYPE_SWAP.l).toBe('r')
    expect(RTL_BORDER_TYPE_SWAP.tl_br).toBe('bl_tr')
    expect(RTL_BORDER_TYPE_SWAP.t).toBeUndefined()
    expect(RTL_BORDER_TYPE_SWAP.b).toBeUndefined()
  })
})

describe('RTL float drag mirror', () => {
  it('mirrorCornerX is an involution swapping east/west and fixing n/s', () => {
    const corners = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const
    for (const corner of corners) expect(mirrorCornerX(mirrorCornerX(corner))).toBe(corner)
    expect(mirrorCornerX('e')).toBe('w')
    expect(mirrorCornerX('ne')).toBe('nw')
    expect(mirrorCornerX('sw')).toBe('se')
    expect(mirrorCornerX('n')).toBe('n')
    expect(mirrorCornerX('s')).toBe('s')
  })

  // Forward mirror: a marker at logical x renders at totalWidth - x, so the
  // visual box is [W - logicalRight, W - logicalLeft]. commitDrag's inverse
  // (negated screen dx, mirrored corner) must preserve screen-space intent.
  const columnWidth = () => 64
  const totalWidth = 10 * 64
  const maxColumn = 9
  const logicalPx = (marker: { index: number; offset: number }) => marker.index * 64 + marker.offset
  const visualBox = (fromX: { index: number; offset: number }, toX: typeof fromX) => ({
    left: totalWidth - logicalPx(toX),
    right: totalWidth - logicalPx(fromX),
  })

  it('move: negated logical shift lands the box exactly +dx on screen', () => {
    const fromX = markerFrom(2, 10 * EMU_PER_PIXEL)
    const toX = markerFrom(5, 30 * EMU_PER_PIXEL)
    const before = visualBox(fromX, toX)
    const screenDx = 100
    const movedFrom = walkMarker(fromX, -screenDx, columnWidth, maxColumn)
    const movedTo = walkMarker(toX, -screenDx, columnWidth, maxColumn)
    const after = visualBox(movedFrom, movedTo)
    expect(after.left).toBeCloseTo(before.left + screenDx)
    expect(after.right - after.left).toBeCloseTo(before.right - before.left)
  })

  it('east resize: mirrored corner grows the visual right edge, left pinned', () => {
    const fromX = markerFrom(2, 10 * EMU_PER_PIXEL)
    const toX = markerFrom(5, 30 * EMU_PER_PIXEL)
    const before = visualBox(fromX, toX)
    const screenDx = 50
    // screen 'e' handle on RTL → logical west edge, negated dx
    expect(mirrorCornerX('e')).toBe('w')
    const resizedFrom = walkMarker(fromX, -screenDx, columnWidth, maxColumn)
    const after = visualBox(resizedFrom, toX)
    expect(after.right).toBeCloseTo(before.right + screenDx)
    expect(after.left).toBeCloseTo(before.left)
  })
})
