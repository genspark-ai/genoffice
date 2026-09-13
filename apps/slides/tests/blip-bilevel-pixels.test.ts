import { describe, it, expect } from 'vitest'
import {
  biLevelPixels,
  processedImageKey,
  hasBlipEffects,
  tracePictureClip,
} from '../src/renderer/konva-adapter'

const rgba = (...v: number[]) => new Uint8ClampedArray(v)

describe('biLevelPixels', () => {
  it('thresholds luminance to pure black or white and keeps alpha', () => {
    // mid gray (lum 0.5), dark navy (lum ~0.14), near-white, translucent light
    const px = rgba(128, 128, 128, 255, 20, 30, 80, 255, 240, 240, 240, 255, 200, 200, 200, 90)
    biLevelPixels(px, 0.25)
    expect(Array.from(px)).toEqual([
      255, 255, 255, 255, 0, 0, 0, 255, 255, 255, 255, 255, 255, 255, 255, 90,
    ])
  })

  it('the threshold is inclusive and 50% is the default cut', () => {
    const px = rgba(127, 127, 127, 255, 128, 128, 128, 255)
    biLevelPixels(px, 0.5)
    expect(Array.from(px.slice(0, 3))).toEqual([0, 0, 0])
    expect(Array.from(px.slice(4, 7))).toEqual([255, 255, 255])
  })

  it('saturated colors are cut by luminance, not by any single channel', () => {
    // pure blue lum 0.114 -> black at 25%; pure green lum 0.587 -> white
    const px = rgba(0, 0, 255, 255, 0, 255, 0, 255)
    biLevelPixels(px, 0.25)
    expect(Array.from(px.slice(0, 3))).toEqual([0, 0, 0])
    expect(Array.from(px.slice(4, 7))).toEqual([255, 255, 255])
  })
})

describe('processedImageKey / hasBlipEffects', () => {
  it('keys a biLevel variant apart from the raw and duotone variants', () => {
    const raw = processedImageKey('data:x', {})
    const bl = processedImageKey('data:x', { biLevel: 0.25 })
    const duo = processedImageKey('data:x', { duotone: ['#000000', '#FFFFFF'] })
    expect(raw).toBe('data:x')
    expect(new Set([raw, bl, duo]).size).toBe(3)
    expect(hasBlipEffects({})).toBe(false)
    expect(hasBlipEffects({ biLevel: 0 })).toBe(true)
  })
})

describe('tracePictureClip', () => {
  const recorder = () => {
    const calls: string[] = []
    const ctx = {
      moveTo: (x: number, y: number) => calls.push(`M${x},${y}`),
      lineTo: (x: number, y: number) => calls.push(`L${x},${y}`),
      arcTo: (...a: number[]) => calls.push(`A${a.join(',')}`),
      closePath: () => calls.push('Z'),
    }
    return { ctx, calls }
  }

  it('traces a polygon clip and closes it', () => {
    const { ctx, calls } = recorder()
    expect(tracePictureClip(ctx, { polygonPoints: [0, 0, 10, 0, 5, 8] }, 10, 8)).toBeUndefined()
    expect(calls).toEqual(['M0,0', 'L10,0', 'L5,8', 'Z'])
  })

  it('clamps the corner radius to half the box', () => {
    const { ctx, calls } = recorder()
    tracePictureClip(ctx, { cornerRadiusPx: 50 }, 20, 10)
    expect(calls[0]).toBe('M5,0')
    expect(calls.filter((c) => c.startsWith('A')).length).toBe(4)
  })
})
