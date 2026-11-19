import { describe, it, expect } from 'vitest'
import {
  emuToPx,
  ptToPx,
  rotToDeg,
  rotToRad,
  makeViewport,
  rectToPx,
  placeTransform,
  EMU_PER_PX_96,
} from '../src/coords'
import type { Transform } from '@genoffice/pptx-engine'

describe('2.1 coordinate system', () => {
  it('emuToPx: 1px = 9525 EMU at scale 1', () => {
    expect(emuToPx(9525, 1)).toBeCloseTo(1, 6)
    expect(emuToPx(EMU_PER_PX_96 * 100, 1)).toBeCloseTo(100, 6)
  })

  it('ptToPx: 18pt = 24px at 96dpi', () => {
    expect(ptToPx(18, 1)).toBeCloseTo(24, 6)
    expect(ptToPx(72, 1)).toBeCloseTo(96, 6)
  })

  it('rotToDeg: 60000 units = 1 degree', () => {
    expect(rotToDeg(60000)).toBe(1)
    expect(rotToDeg(2700000)).toBe(45)
  })

  it('makeViewport: 16:9 slide maps to correct aspect', () => {
    // Standard 16:9: 12192000 x 6858000 EMU
    const vp = makeViewport({ cx: 12192000, cy: 6858000 }, 1280)
    expect(vp.widthPx).toBe(1280)
    expect(vp.heightPx).toBeCloseTo(720, 0) // 16:9
  })

  it('rectToPx: scales geometry by viewport scale', () => {
    const size = { cx: 9525 * 1000, cy: 9525 * 562 } // natural 1000x562 px
    const vp = makeViewport(size, 500) // scaled to half
    const px = rectToPx({ x: 9525 * 100, y: 9525 * 50, cx: 9525 * 200, cy: 9525 * 100 }, vp)
    expect(px.x).toBeCloseTo(50, 4)
    expect(px.y).toBeCloseTo(25, 4)
    expect(px.w).toBeCloseTo(100, 4)
    expect(px.h).toBeCloseTo(50, 4)
  })

  it('returns 0 for non-finite coordinate inputs', () => {
    expect(emuToPx(NaN)).toBe(0)
    expect(emuToPx(Infinity)).toBe(0)
    expect(ptToPx(NaN)).toBe(0)
    expect(ptToPx(12, Infinity)).toBe(0)
    expect(rotToDeg(NaN)).toBe(0)
    expect(rotToRad(Infinity)).toBe(0)
  })

  it('placeTransform: computes center + rotation + parent offset', () => {
    const size = { cx: 9525 * 1000, cy: 9525 * 1000 }
    const vp = makeViewport(size, 1000) // scale 1
    const t: Transform = {
      offset: { x: 9525 * 100, y: 9525 * 100, cx: 9525 * 200, cy: 9525 * 80 },
      rot: 2700000, // 45deg
      flipH: true,
      flipV: false,
    }
    const box = placeTransform(t, vp, { x: 10, y: 20 })
    expect(box.x).toBeCloseTo(110, 4)
    expect(box.y).toBeCloseTo(120, 4)
    expect(box.w).toBeCloseTo(200, 4)
    expect(box.h).toBeCloseTo(80, 4)
    expect(box.rotationDeg).toBe(45)
    expect(box.flipH).toBe(true)
    expect(box.centerX).toBeCloseTo(210, 4)
    expect(box.centerY).toBeCloseTo(160, 4)
  })
})

describe('group child scaling of quarter-turned children', () => {
  it('placeTransform: a 90° child in a non-uniformly scaled group scales its visual box', () => {
    const vp = makeViewport({ cx: 9525 * 1000, cy: 9525 * 1000 }, 1000)
    const t: Transform = {
      offset: { x: 0, y: 0, cx: 9525 * 200, cy: 9525 * 50 },
      rot: 5400000,
      flipH: false,
      flipV: false,
    }
    // visual box 50 wide × 200 tall → 50 × 600 on screen, so the shape's own w/h swap factors
    const box = placeTransform(t, vp, { x: 0, y: 0, scaleX: 1, scaleY: 3 })
    expect(box.w).toBeCloseTo(600, 4)
    expect(box.h).toBeCloseTo(50, 4)
    expect(box.centerX).toBeCloseTo(100, 4)
    expect(box.centerY).toBeCloseTo(75, 4)
    // unrotated children keep the plain per-axis scaling
    const flat = placeTransform({ ...t, rot: 0 }, vp, { x: 0, y: 0, scaleX: 1, scaleY: 3 })
    expect(flat.w).toBeCloseTo(200, 4)
    expect(flat.h).toBeCloseTo(150, 4)
  })
})
