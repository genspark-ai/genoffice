import { describe, expect, it } from 'vitest'

import { anchorAxisMarkers } from '../src/renderer/shape-draw'

const col20 = () => 20

describe('anchorAxisMarkers', () => {
  it('walks LTR from the reference cell left edge', () => {
    const { from, to } = anchorAxisMarkers({ index: 0, offset: 0 }, 0, 25, 30, 1, false, col20, 100)
    expect(from).toEqual({ index: 1, offset: 5 })
    expect(to).toEqual({ index: 2, offset: 15 })
  })

  it('divides screen distances by the zoom factor', () => {
    const { from, to } = anchorAxisMarkers({ index: 0, offset: 0 }, 0, 50, 60, 2, false, col20, 100)
    expect(from).toEqual({ index: 1, offset: 5 })
    expect(to).toEqual({ index: 2, offset: 15 })
  })

  it('starts the walk from a scrolled origin marker', () => {
    const { from } = anchorAxisMarkers({ index: 4, offset: 0 }, 200, 225, 10, 1, false, col20, 100)
    expect(from).toEqual({ index: 5, offset: 5 })
  })

  it('RTL: the from marker is the rectangle visual right edge, walked leftward', () => {
    // Reference cell's visual right edge at 100px; rect spans [30, 60] on
    // screen, so its right edge sits 40 logical px into the column axis.
    const { from, to } = anchorAxisMarkers(
      { index: 0, offset: 0 },
      100,
      30,
      30,
      1,
      true,
      col20,
      100,
    )
    expect(from).toEqual({ index: 2, offset: 0 })
    expect(to).toEqual({ index: 3, offset: 10 })
  })

  it('RTL: a rect drawn past the reference right edge clamps at the sheet start', () => {
    const { from, to } = anchorAxisMarkers(
      { index: 0, offset: 0 },
      100,
      90,
      30,
      1,
      true,
      col20,
      100,
    )
    expect(from).toEqual({ index: 0, offset: 0 })
    expect(to).toEqual({ index: 1, offset: 10 })
  })

  it('RTL mirror equals LTR of the mirrored rectangle', () => {
    const total = 200
    const rect = { x: 35, w: 48 }
    const mirroredX = total - (rect.x + rect.w)
    const rtl = anchorAxisMarkers(
      { index: 0, offset: 0 },
      total,
      rect.x,
      rect.w,
      1,
      true,
      col20,
      100,
    )
    const ltr = anchorAxisMarkers(
      { index: 0, offset: 0 },
      0,
      mirroredX,
      rect.w,
      1,
      false,
      col20,
      100,
    )
    expect(rtl).toEqual(ltr)
  })
})
