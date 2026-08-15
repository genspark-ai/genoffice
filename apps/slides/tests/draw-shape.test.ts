import { describe, it, expect } from 'vitest'
import {
  isLineDrawKind,
  isStraightLineKind,
  defaultDrawSize,
  resolveDrawRect,
} from '../src/renderer/draw-shape'

describe('draw kind classification', () => {
  it('recognizes line kinds', () => {
    for (const k of ['line', 'lineArrow', 'lineArrowDouble', 'lineBent', 'lineCurved']) {
      expect(isLineDrawKind(k)).toBe(true)
    }
    expect(isLineDrawKind('rect')).toBe(false)
    expect(isLineDrawKind('ellipse')).toBe(false)
  })

  it('recognizes straight lines (zero-height default)', () => {
    expect(isStraightLineKind('line')).toBe(true)
    expect(isStraightLineKind('lineArrowDouble')).toBe(true)
    expect(isStraightLineKind('lineBent')).toBe(false)
  })
})

describe('defaultDrawSize (single click = PowerPoint 1in defaults)', () => {
  it('shapes default to a 96x96 square (1x1 inch)', () => {
    expect(defaultDrawSize('rect')).toEqual({ w: 96, h: 96 })
    expect(defaultDrawSize('ellipse')).toEqual({ w: 96, h: 96 })
    expect(defaultDrawSize('octagon')).toEqual({ w: 96, h: 96 })
  })

  it('straight lines default to a 1in horizontal stroke', () => {
    expect(defaultDrawSize('line')).toEqual({ w: 96, h: 0 })
    expect(defaultDrawSize('lineArrow')).toEqual({ w: 96, h: 0 })
  })

  it('bent/curved connectors need a real box for their elbow geometry', () => {
    expect(defaultDrawSize('lineBent')).toEqual({ w: 96, h: 96 })
    expect(defaultDrawSize('lineCurved')).toEqual({ w: 96, h: 96 })
  })
})

describe('resolveDrawRect: shapes', () => {
  it('free drag down-right maps to the dragged rectangle', () => {
    expect(resolveDrawRect('rect', 10, 20, 110, 60, false)).toEqual({ x: 10, y: 20, w: 100, h: 40 })
  })

  it('drag up-left normalizes to a positive-size rectangle', () => {
    expect(resolveDrawRect('rect', 110, 60, 10, 20, false)).toEqual({ x: 10, y: 20, w: 100, h: 40 })
  })

  it('shift constrains to a square using the larger delta, anchored at the start point', () => {
    expect(resolveDrawRect('rect', 10, 20, 110, 60, true)).toEqual({ x: 10, y: 20, w: 100, h: 100 })
  })

  it('shift square follows the drag direction (up-left drag grows up-left)', () => {
    expect(resolveDrawRect('rect', 110, 120, 30, 60, true)).toEqual({ x: 30, y: 40, w: 80, h: 80 })
  })
})

describe('resolveDrawRect: lines', () => {
  it('down-right drag: box spans the two endpoints, no flip', () => {
    expect(resolveDrawRect('lineArrow', 0, 0, 100, 50, false)).toEqual({
      x: 0,
      y: 0,
      w: 100,
      h: 50,
    })
  })

  it('up-right drag: flipV so the head lands on the drag end', () => {
    expect(resolveDrawRect('lineArrow', 0, 50, 100, 0, false)).toEqual({
      x: 0,
      y: 0,
      w: 100,
      h: 50,
      flipV: true,
    })
  })

  it('down-left drag: flipH', () => {
    expect(resolveDrawRect('lineArrow', 100, 0, 0, 50, false)).toEqual({
      x: 0,
      y: 0,
      w: 100,
      h: 50,
      flipH: true,
    })
  })

  it('up-left drag: flipH + flipV', () => {
    expect(resolveDrawRect('lineArrow', 100, 50, 0, 0, false)).toEqual({
      x: 0,
      y: 0,
      w: 100,
      h: 50,
      flipH: true,
      flipV: true,
    })
  })

  it('shift snaps a near-horizontal line to exactly horizontal', () => {
    expect(resolveDrawRect('line', 0, 0, 100, 10, true)).toEqual({ x: 0, y: 0, w: 100, h: 0 })
  })

  it('shift snaps a near-vertical line to exactly vertical', () => {
    expect(resolveDrawRect('line', 50, 0, 44, 80, true)).toEqual({ x: 50, y: 0, w: 0, h: 80 })
  })

  it('shift snaps a diagonal-ish line to exactly 45 degrees', () => {
    const r = resolveDrawRect('line', 0, 0, 90, 110, true)
    expect(r.x).toBe(0)
    expect(r.y).toBe(0)
    expect(r.w).toBe(r.h)
    expect(r.w).toBeGreaterThan(0)
  })
})
