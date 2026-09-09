import { describe, expect, it } from 'vitest'

import {
  shapeRunFontSize,
  shapeTextOverflowClass,
  shapeTextScaleStyle,
  shapeZoomedPx,
} from '../src/renderer/shape-text-scale'

// Mirrors the browser: the float box is `frameWidth * zoom` wide, so 100cqw
// resolves to that many px.
function resolvePx(expression: string, frameWidth: number, zoom: number): number {
  const unit = expression.match(/^calc\(100cqw \/ ([\d.]+)\)$/)
  if (unit) return (frameWidth * zoom) / Number(unit[1])
  const scaled = expression.match(/^calc\(var\(--shape-px, 1px\) \* ([\d.]+)\)$/)
  if (scaled) return resolvePx(shapeZoomedPx(frameWidth), frameWidth, zoom) * Number(scaled[1])
  const px = expression.match(/^([\d.]+)px$/)
  if (!px) throw new Error(`unresolvable ${expression}`)
  return Number(px[1])
}

describe('shape text zoom scaling', () => {
  it('one shape px is the float width over the logical frame width', () => {
    expect(shapeZoomedPx(250)).toBe('calc(100cqw / 250)')
    expect(shapeTextScaleStyle(250)).toEqual({ '--shape-px': 'calc(100cqw / 250)' })
  })

  it.each([0.6, 1, 1.25])('resolves to the zoom factor at zoom %s', (zoom) => {
    expect(resolvePx(shapeZoomedPx(250), 250, zoom)).toBeCloseTo(zoom, 6)
  })

  it.each([0.6, 1, 1.25])('an 11pt run is 14.667px times the zoom at %s', (zoom) => {
    expect(resolvePx(shapeRunFontSize(11), 250, zoom)).toBeCloseTo(11 * (4 / 3) * zoom, 2)
  })

  it('falls back to logical px without a frame', () => {
    expect(shapeZoomedPx(undefined)).toBe('1px')
    expect(shapeZoomedPx(0)).toBe('1px')
    expect(resolvePx(shapeZoomedPx(undefined), 250, 0.6)).toBe(1)
  })
})

describe('shape text overflow mapping', () => {
  it('clips only when bodyPr asks for clip or ellipsis on either axis', () => {
    expect(shapeTextOverflowClass(undefined, undefined)).toBe('')
    expect(shapeTextOverflowClass('overflow', 'overflow')).toBe('')
    expect(shapeTextOverflowClass('clip', 'clip')).toBe(' shape-text-clip')
    expect(shapeTextOverflowClass('clip', undefined)).toBe(' shape-text-clip')
    expect(shapeTextOverflowClass(undefined, 'ellipsis')).toBe(' shape-text-clip')
    expect(shapeTextOverflowClass('bogus', 'overflow')).toBe('')
  })
})
