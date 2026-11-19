import { describe, it, expect, vi } from 'vitest'

// react-konva's node entry requires the native 'canvas' package; only the pure helper is under test
vi.mock('react-konva', () => {
  const stub = () => null
  return {
    Stage: stub,
    Layer: stub,
    Rect: stub,
    Group: stub,
    Transformer: stub,
    Line: stub,
    Arrow: stub,
    Text: stub,
    Ellipse: stub,
    Image: stub,
    Path: stub,
    Circle: stub,
    Arc: stub,
  }
})

import {
  canvasPixelRatio,
  DENSE_SLIDE_NODE_COUNT,
  LIGHT_SLIDE_NODE_COUNT,
} from '../src/renderer/SlideCanvas'

describe('canvasPixelRatio', () => {
  it('follows zoom above 100% so the bitmap matches the displayed resolution (blur)', () => {
    expect(canvasPixelRatio(2, 1.54)).toBeCloseTo(3.08)
    expect(canvasPixelRatio(1, 1.54)).toBeCloseTo(1.54)
    expect(canvasPixelRatio(1, 2)).toBe(2)
  })

  it('never drops below devicePixelRatio when zoomed out', () => {
    expect(canvasPixelRatio(2, 0.25)).toBe(2)
    expect(canvasPixelRatio(2, 0.7)).toBe(2)
    expect(canvasPixelRatio(1, 0.5)).toBe(1)
  })

  it('light pages get the full dpr×zoom up to 6 (crisp content and selection chrome at deep zoom)', () => {
    expect(canvasPixelRatio(2, 3)).toBe(6)
    expect(canvasPixelRatio(2, 3, LIGHT_SLIDE_NODE_COUNT - 1)).toBe(6)
    expect(canvasPixelRatio(3, 3)).toBe(6) // capped at 6
  })

  it('mid pages keep the historical 3 cap', () => {
    expect(canvasPixelRatio(2, 3, LIGHT_SLIDE_NODE_COUNT)).toBe(3)
    expect(canvasPixelRatio(3, 2, LIGHT_SLIDE_NODE_COUNT)).toBe(3)
    expect(canvasPixelRatio(2, 1, LIGHT_SLIDE_NODE_COUNT)).toBe(2) // below the cap nothing changes
  })

  it('treats a missing/zero devicePixelRatio as 1', () => {
    expect(canvasPixelRatio(0, 1)).toBe(1)
    expect(canvasPixelRatio(0, 2)).toBe(2)
  })

  it('caps dense slides at min(dpr, 2) regardless of zoom (raster-pressure mitigation)', () => {
    const dense = DENSE_SLIDE_NODE_COUNT
    expect(canvasPixelRatio(2, 2, dense)).toBe(2) // zoom upscale skipped
    expect(canvasPixelRatio(3, 1, dense)).toBe(2) // hard cap below dpr 3
    expect(canvasPixelRatio(1, 3, dense)).toBe(1)
    expect(canvasPixelRatio(0, 2, dense)).toBe(1)
    expect(canvasPixelRatio(2, 2, dense - 1)).toBe(3) // mid tier: the 3 cap still applies
  })
})
