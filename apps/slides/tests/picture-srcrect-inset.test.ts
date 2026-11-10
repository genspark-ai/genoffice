/**
 * cropToKonva: negative srcRect values are insets — the image occupies a sub-rect
 * of the frame and the remaining bands stay empty, matching PowerPoint (common in
 * Google Slides exports). Non-negative crops keep the legacy full-frame mapping.
 */
import { describe, it, expect } from 'vitest'
import { cropToKonva } from '../src/renderer/konva-adapter'

const img = { width: 1000, height: 500 } as HTMLImageElement
const pic = (srcRect?: { l: number; t: number; r: number; b: number }) =>
  ({ srcRect, box: { w: 800, h: 400 } }) as any

describe('cropToKonva', () => {
  it('returns {} without srcRect or image', () => {
    expect(cropToKonva(pic(undefined), img)).toEqual({})
    expect(cropToKonva(pic({ l: 0.1, t: 0, r: 0, b: 0 }), undefined)).toEqual({})
  })

  it('positive crop maps to source pixels and keeps the full frame', () => {
    const out = cropToKonva(pic({ l: 0.1, t: 0.05, r: 0.1, b: 0.05 }), img)
    expect(out.crop!.x).toBe(100)
    expect(out.crop!.y).toBe(25)
    expect(out.crop!.width).toBeCloseTo(800)
    expect(out.crop!.height).toBeCloseTo(450)
    expect(out.x).toBeUndefined()
    expect(out.width).toBeUndefined()
  })

  it('negative top inset leaves a blank band and shows the full image below it', () => {
    // t=-0.25: source span is 1.25 of the image height mapped onto 400px → the
    // image starts 1/5 down the frame (80px) and fills the remaining 320px.
    const out = cropToKonva(pic({ l: 0, t: -0.25, r: 0, b: 0 }), img)
    expect(out.crop).toEqual({ x: 0, y: 0, width: 1000, height: 500 })
    expect(out.x).toBe(0)
    expect(out.y).toBeCloseTo(80)
    expect(out.width).toBeCloseTo(800)
    expect(out.height).toBeCloseTo(320)
  })

  it('mixed negative and positive components combine', () => {
    // l=-0.2 inset, r=0.2 crop: spanX = 1 - (-0.2) - 0.2 = 1; visible source 0..0.8
    const out = cropToKonva(pic({ l: -0.2, t: 0, r: 0.2, b: 0 }), img)
    expect(out.crop!.x).toBe(0)
    expect(out.crop!.width).toBeCloseTo(800)
    expect(out.crop!.height).toBeCloseTo(500)
    expect(out.x).toBeCloseTo(160) // 800 * (0 - (-0.2)) / 1
    expect(out.width).toBeCloseTo(640) // 800 * 0.8 / 1
    expect(out.y).toBe(0)
    expect(out.height).toBeCloseTo(400)
  })
})
