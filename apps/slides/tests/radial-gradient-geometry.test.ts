import { describe, it, expect } from 'vitest'
import { fillToKonva, radialCircleGeometry } from '../src/renderer/konva-adapter'

// Google Slides corner radial: <a:path path="circle"><a:fillToRect b="100%" r="100%"/></a:path>
// <a:tileRect l="-100%" t="-100%"/> -> focus at the shape's top-left, 100% ring on the
// bottom-right corner (measured on PowerPoint: top-left pixel is the pos-0 color, the
// far corner reaches the pos-1 color)
describe('radialCircleGeometry', () => {
  it('top-left fillToRect focus with a tileRect grown up/left runs corner to corner', () => {
    const g = radialCircleGeometry(580, 305, { x: 0, y: 0 }, { l: -1, t: -1, r: 0, b: 0 })
    expect(g.cx).toBe(0)
    expect(g.cy).toBe(0)
    expect(g.r).toBeCloseTo(Math.hypot(580, 305), 6)
  })

  it('mirrored idiom (fillToRect l=t=100%, tileRect r=b=-100%) focuses bottom-right', () => {
    const g = radialCircleGeometry(400, 200, { x: 1, y: 1 }, { l: 0, t: 0, r: -1, b: -1 })
    expect(g.cx).toBe(400)
    expect(g.cy).toBe(200)
    expect(g.r).toBeCloseTo(Math.hypot(400, 200), 6)
  })

  it('default centered case keeps the ring on the shape corner', () => {
    const g = radialCircleGeometry(400, 200)
    expect(g.cx).toBe(200)
    expect(g.cy).toBe(100)
    expect(g.r).toBeCloseTo(Math.hypot(200, 100), 6)
    expect(radialCircleGeometry(400, 200, { x: 0.5, y: 0.5 }, { l: 0, t: 0, r: 0, b: 0 })).toEqual(
      g,
    )
  })

  it('positive tileRect insets never shrink the ring below the shape corner', () => {
    const g = radialCircleGeometry(400, 200, { x: 0.5, y: 0.5 }, { l: 0.2, t: 0.2, r: 0.2, b: 0.2 })
    expect(g.r).toBeCloseTo(Math.hypot(200, 100), 6)
  })

  it('fillToKonva circle fill uses the derived center and radius', () => {
    const r = fillToKonva(
      {
        kind: 'gradient',
        stops: [
          { pos: 0, color: '#5DE0E6' },
          { pos: 1, color: '#004AAD' },
        ],
        angleDeg: 0,
        radial: true,
        path: 'circle',
        center: { x: 0, y: 0 },
        tileRect: { l: -1, t: -1, r: 0, b: 0 },
      },
      580,
      305,
    )
    expect(r.fillRadialGradientStartPoint).toEqual({ x: 0, y: 0 })
    expect(r.fillRadialGradientEndRadius).toBeCloseTo(Math.hypot(580, 305), 6)
    expect(r.fillRadialGradientColorStops!.slice(-2)).toEqual([1, '#004AAD'])
  })
})
