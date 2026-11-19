/**
 * Connector lines with a gradient stroke (<a:ln><a:gradFill>): theme divider
 * rules fade in from transparent ends, so drawing them in the first stop's
 * color alone made them vanish from the canvas while PowerPoint shows them.
 */
import { describe, expect, it } from 'vitest'
import type { RenderStroke } from '@genoffice/pptx-render'
import {
  connectorHeadColor,
  connectorStrokeProps,
  strokeToKonva,
} from '../src/renderer/konva-adapter'

const fading: RenderStroke = {
  color: '#4472C400',
  widthPx: 1,
  widthPt: 0.75,
  gradient: {
    angleDeg: 0,
    stops: [
      { pos: 0, color: '#4472C400' },
      { pos: 0.5, color: '#4472C4' },
      { pos: 1, color: '#4472C400' },
    ],
  },
}

describe('connector gradient strokes', () => {
  it('carries the gradient ramp onto the line, spanning the connector box', () => {
    const props = connectorStrokeProps(strokeToKonva(fading, { w: 400, h: 2 }))
    expect(props.strokeWidth).toBe(1)
    expect(props.strokeLinearGradientStartPoint).toEqual({ x: 0, y: 1 })
    expect(props.strokeLinearGradientEndPoint).toEqual({ x: 400, y: 1 })
    const stops = props.strokeLinearGradientColorStops!
    expect(stops[0]).toBe(0)
    expect(stops[1]).toBe('rgba(68,114,196,0.000)')
    expect(stops).toContain('#4472C4')
  })

  it('keeps a plain stroke plain and defaults a missing one to a 1px black line', () => {
    expect(
      connectorStrokeProps(strokeToKonva({ color: '#FF0000', widthPx: 3, widthPt: 2.25 })),
    ).toEqual({
      stroke: '#FF0000',
      strokeWidth: 3,
    })
    expect(connectorStrokeProps(strokeToKonva(undefined))).toEqual({
      stroke: '#000000',
      strokeWidth: 1,
    })
  })

  it('fills arrowheads with the most opaque stop instead of a transparent end', () => {
    expect(connectorHeadColor(strokeToKonva(fading, { w: 400, h: 2 }))).toBe('#4472C4')
    expect(connectorHeadColor(strokeToKonva({ color: '#FF0000', widthPx: 1, widthPt: 0.75 }))).toBe(
      '#FF0000',
    )
  })
})
