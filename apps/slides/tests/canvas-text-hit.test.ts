import type { ShapeRenderNode } from '@genoffice/pptx-render'
import { describe, expect, it } from 'vitest'
import { needsTextFrameHitArea } from '../src/renderer/text-hit-area'

function shape(presetGeometry?: string, hasText = true): ShapeRenderNode {
  return {
    type: 'shape',
    presetGeometry,
    ...(hasText ? { text: { lines: [] } } : {}),
  } as unknown as ShapeRenderNode
}

describe('canvas text-frame hit area', () => {
  it('covers text frames regardless of their geometry', () => {
    expect(needsTextFrameHitArea(shape('rect'))).toBe(true)
    expect(needsTextFrameHitArea(shape('ellipse'))).toBe(true)
    expect(needsTextFrameHitArea(shape('roundRect'))).toBe(true)
    expect(needsTextFrameHitArea(shape('custom'))).toBe(true)
  })

  it('does not add a text-frame target to shapes without text', () => {
    expect(needsTextFrameHitArea(shape('ellipse', false))).toBe(false)
  })
})
