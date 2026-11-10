import { describe, expect, it } from 'vitest'

import {
  clampExplicitTo,
  EMU_PER_PIXEL,
  markerFrom,
  walkMarker,
} from '../src/renderer/WorkbookVisuals'

const columnWidth = (): number => 64

describe('clampExplicitTo', () => {
  it('clamps a real <xdr:to> offset at the cell edge like Excel', () => {
    // closedxml_picture-webp.xlsx: to = col 0 + 3524250 EMU (~370px) while
    // col 0 is ~64px wide — Excel shows a col-sized sliver, not 370px.
    const to = markerFrom(0, 3524250)
    expect(to.offset).toBeCloseTo(3524250 / EMU_PER_PIXEL)
    expect(clampExplicitTo(to, true, columnWidth)).toEqual({ index: 0, offset: 64 })
  })

  it('keeps in-cell offsets untouched', () => {
    const to = markerFrom(3, 10 * EMU_PER_PIXEL)
    expect(clampExplicitTo(to, true, columnWidth)).toBe(to)
  })

  it('never clamps synthesized markers (oneCellAnchor ext, absoluteAnchor)', () => {
    const to = markerFrom(0, 3524250)
    expect(clampExplicitTo(to, false, columnWidth)).toBe(to)
  })

  it('drag commits walk from the clamped edge, not the raw overflow', () => {
    // commitDrag seeds its to marker this way: without the clamp, walkMarker
    // normalizes the ~370px overflow into real columns and the committed
    // anchor re-expands the visual to its walked size.
    const to = clampExplicitTo(markerFrom(0, 3524250), true, columnWidth)
    expect(walkMarker(to, 10, columnWidth, 100)).toEqual({ index: 1, offset: 10 })
    // Normal in-cell anchors take the identical path (clamp is a no-op).
    const normal = clampExplicitTo(markerFrom(2, 10 * EMU_PER_PIXEL), true, columnWidth)
    expect(walkMarker(normal, 10, columnWidth, 100)).toEqual({ index: 2, offset: 20 })
  })
})
