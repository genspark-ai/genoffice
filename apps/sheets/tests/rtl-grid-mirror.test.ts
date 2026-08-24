import { describe, expect, it } from 'vitest'

import { mirrorSpanX, RTL_BORDER_TYPE_SWAP } from '../src/renderer/rtl-grid-mirror'
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
