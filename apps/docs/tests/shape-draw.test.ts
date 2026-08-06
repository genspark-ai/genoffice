import { describe, it, expect } from 'vitest'
import {
  DEFAULT_SHAPE_EMU,
  drawRectToEmu,
  resolveDrawRect,
} from '../src/renderer/editor/shape-draw'

describe('resolveDrawRect', () => {
  it('free drag down-right maps to the dragged rectangle', () => {
    expect(resolveDrawRect(10, 20, 110, 60, false)).toEqual({ x: 10, y: 20, w: 100, h: 40 })
  })

  it('drag up-left normalizes to a positive-size rectangle', () => {
    expect(resolveDrawRect(110, 60, 10, 20, false)).toEqual({ x: 10, y: 20, w: 100, h: 40 })
  })

  it('shift constrains to a square using the larger delta, anchored at the start', () => {
    expect(resolveDrawRect(10, 20, 110, 60, true)).toEqual({ x: 10, y: 20, w: 100, h: 100 })
  })

  it('shift square follows the drag direction (up-left drag grows up-left)', () => {
    expect(resolveDrawRect(110, 120, 30, 60, true)).toEqual({ x: 30, y: 40, w: 80, h: 80 })
  })
})

describe('drawRectToEmu', () => {
  it('converts drawn px to EMU at 9525 EMU/px', () => {
    expect(drawRectToEmu({ x: 0, y: 0, w: 96, h: 48 }, 1)).toEqual({
      widthEmu: 96 * 9525,
      heightEmu: 48 * 9525,
    })
  })

  it('divides out the page CSS zoom', () => {
    expect(drawRectToEmu({ x: 0, y: 0, w: 192, h: 96 }, 2)).toEqual({
      widthEmu: 96 * 9525,
      heightEmu: 48 * 9525,
    })
  })

  it('floors degenerate drags at 1px so the OOXML extent stays valid', () => {
    expect(drawRectToEmu({ x: 0, y: 0, w: 0, h: 10 }, 1).widthEmu).toBe(9525)
  })
})

describe('single-click default', () => {
  it('is Word parity: 1x1 inch', () => {
    expect(DEFAULT_SHAPE_EMU).toBe(914400)
  })
})
