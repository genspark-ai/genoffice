import { describe, it, expect, beforeEach, afterEach } from 'vitest'
import { fillToKonva } from '../src/renderer/konva-adapter'

/** jsdom has no 2d context: stub it with a recorder that returns fixed pixels. */
const drawCalls: unknown[][] = []
let pixels: number[] = []
const origGetContext = HTMLCanvasElement.prototype.getContext

beforeEach(() => {
  drawCalls.length = 0
  HTMLCanvasElement.prototype.getContext = function () {
    return {
      drawImage: (...args: unknown[]) => void drawCalls.push(args),
      getImageData: () => ({ data: Uint8ClampedArray.from(pixels) }),
    } as unknown as CanvasRenderingContext2D
  } as unknown as typeof HTMLCanvasElement.prototype.getContext
})

afterEach(() => {
  HTMLCanvasElement.prototype.getContext = origGetContext
})

const imageFill = (extra = {}) =>
  ({ kind: 'image', dataUrl: 'data:x', mode: 'stretch', ...extra }) as any

const imgOf = (w: number, h: number) => ({ width: w, height: h }) as unknown as HTMLImageElement

const imagesMap = (img: HTMLImageElement) => new Map([['data:x', img]])

describe('fillToKonva degenerate stretch textures (tdf146223)', () => {
  it('flattens a 2×2 stretch blip to its mean color', () => {
    // 0, 57, 32, 92 gray pixels → mean 45 = #2d2d2d
    pixels = [0, 0, 0, 255, 57, 57, 57, 255, 32, 32, 32, 255, 92, 92, 92, 255]
    const r = fillToKonva(imageFill(), 1280, 720, imagesMap(imgOf(2, 2)))
    expect(r).toEqual({ fill: '#2d2d2d' })
  })

  it('keeps alphaModFix as node opacity on the flattened fill', () => {
    pixels = [255, 0, 0, 255, 255, 0, 0, 255, 255, 0, 0, 255, 255, 0, 0, 255]
    const fill = { ...imageFill({ alpha: 0.7 }), dataUrl: 'data:red' }
    const r = fillToKonva(fill, 1280, 720, new Map([['data:red', imgOf(2, 2)]]))
    expect(r.fill).toBe('#ff0000')
    expect(r.opacity).toBe(0.7)
  })

  it('does not flatten real-size images', () => {
    const r = fillToKonva(imageFill(), 1280, 720, imagesMap(imgOf(155, 93)))
    expect(r.fillPatternImage).toBeTruthy()
    expect(r.fill).toBeUndefined()
  })
})

describe('fillToKonva stretch fillRect insets (tdf153466)', () => {
  it('composites the image into a transparent-padded tile covering the whole shape', () => {
    const fr = { l: 0.55, t: 0.56, r: 0, b: 0 }
    const r = fillToKonva(imageFill({ fillRect: fr }), 1280, 720, imagesMap(imgOf(155, 93)))
    const tile = r.fillPatternImage as unknown as HTMLCanvasElement
    expect(tile.tagName).toBe('CANVAS')
    // tile spans the whole shape: image size / covered fraction
    expect(tile.width).toBe(Math.round(155 / 0.45))
    expect(tile.height).toBe(Math.round(93 / 0.44))
    // pattern scale stretches the tile (not the raw image) over the shape, no offset needed
    expect(r.fillPatternScaleX).toBeCloseTo(1280 / tile.width, 5)
    expect(r.fillPatternScaleY).toBeCloseTo(720 / tile.height, 5)
    expect(r.fillPatternX).toBeUndefined()
    // the image was drawn into the inset subrect of the tile
    const [, dx, dy, dw, dh] = drawCalls.at(-1) as number[]
    expect(dx).toBeCloseTo(0.55 * tile.width, 3)
    expect(dy).toBeCloseTo(0.56 * tile.height, 3)
    expect(dw).toBeCloseTo(0.45 * tile.width, 3)
    expect(dh).toBeCloseTo(0.44 * tile.height, 3)
  })

  it('plain stretch still scales the raw image over the shape', () => {
    const img = imgOf(155, 93)
    const r = fillToKonva(imageFill(), 1280, 720, imagesMap(img))
    expect(r.fillPatternImage).toBe(img)
    expect(r.fillPatternScaleX).toBeCloseTo(1280 / 155, 5)
  })
})

describe('a:lin scaled gradients stretch the angle with the box aspect (prod_048 measured)', () => {
  it('45° scaled on a wide box turns near-vertical: direction ∝ (h·cosθ, w·sinθ)', () => {
    const r = fillToKonva(
      {
        kind: 'gradient',
        angleDeg: 45,
        scaled: true,
        stops: [
          { pos: 0, color: '#FF0000' },
          { pos: 1, color: '#00B050' },
        ],
      } as any,
      1000,
      100,
    )
    const s = r.fillLinearGradientStartPoint!
    const e = r.fillLinearGradientEndPoint!
    const [dx, dy] = [e.x - s.x, e.y - s.y]
    // (h, w)/|.| = (100, 1000) normalized -> dy/dx = 10
    expect(dy / dx).toBeCloseTo(10, 5)
    expect(dx).toBeGreaterThan(0)
  })

  it('unscaled 45° keeps the true diagonal direction', () => {
    const r = fillToKonva(
      {
        kind: 'gradient',
        angleDeg: 45,
        stops: [
          { pos: 0, color: '#FF0000' },
          { pos: 1, color: '#00B050' },
        ],
      } as any,
      1000,
      100,
    )
    const s = r.fillLinearGradientStartPoint!
    const e = r.fillLinearGradientEndPoint!
    expect((e.y - s.y) / (e.x - s.x)).toBeCloseTo(1, 5)
  })

  it('axis-aligned angles are unchanged by scaling', () => {
    const r = fillToKonva(
      {
        kind: 'gradient',
        angleDeg: 90,
        scaled: true,
        stops: [
          { pos: 0, color: '#FF0000' },
          { pos: 1, color: '#00B050' },
        ],
      } as any,
      1000,
      100,
    )
    const s = r.fillLinearGradientStartPoint!
    const e = r.fillLinearGradientEndPoint!
    expect(e.x - s.x).toBeCloseTo(0, 5)
    expect(e.y - s.y).toBeCloseTo(100, 5)
  })
})

describe('gradient ramps interpolate in linear sRGB (tdf105739)', () => {
  it('subdivides a two-stop ramp with linear-light midpoints', () => {
    const r = fillToKonva(
      {
        kind: 'gradient',
        angleDeg: 45,
        stops: [
          { pos: 0, color: '#FF0000' },
          { pos: 1, color: '#00B050' },
        ],
      } as any,
      100,
      100,
    )
    const stops = r.fillLinearGradientColorStops!
    // 2 original + 7 inserted midpoints, position/color interleaved
    expect(stops.length).toBe((2 + 7) * 2)
    const mid = stops[stops.indexOf(0.5) + 1] as string
    // PowerPoint-measured midpoint (188,129,55); raw sRGB blending would give (128,88,40)
    const [r0, g0, b0] = mid.match(/\d+/g)!.map(Number)
    expect(Math.abs(r0 - 188)).toBeLessThanOrEqual(2)
    expect(Math.abs(g0 - 129)).toBeLessThanOrEqual(2)
    expect(Math.abs(b0 - 55)).toBeLessThanOrEqual(2)
  })

  it('alpha fades interpolate straight: color ramps toward the transparent stop hue', () => {
    const r = fillToKonva(
      {
        kind: 'gradient',
        angleDeg: 0,
        stops: [
          { pos: 0, color: '#FFFFFF00' },
          { pos: 1, color: '#29354D' },
        ],
      } as any,
      100,
      100,
    )
    const stops = r.fillLinearGradientColorStops!
    // PowerPoint-measured (controlled probe over black + white backdrops, both stop
    // orders): the midpoint color is the linear-sRGB straight mix toward white,
    // NOT the premultiplied navy — alpha ramps linearly.
    const mid = stops[stops.indexOf(0.5) + 1] as string
    const [r0, g0, b0, a0] = mid.match(/[\d.]+/g)!.map(Number)
    expect(Math.abs(r0 - 189)).toBeLessThanOrEqual(2)
    expect(Math.abs(g0 - 190)).toBeLessThanOrEqual(2)
    expect(Math.abs(b0 - 194)).toBeLessThanOrEqual(2)
    expect(a0).toBeCloseTo(0.5, 2)
  })

  it('transparent stop color is masked when two visible stops remain (fade stays navy)', () => {
    const r = fillToKonva(
      {
        kind: 'gradient',
        angleDeg: 0,
        stops: [
          { pos: 0, color: '#FFFFFF00' },
          { pos: 0.5, color: '#29354D' },
          { pos: 1, color: '#29354D' },
        ],
      } as any,
      100,
      100,
    )
    const stops = r.fillLinearGradientColorStops!
    // PowerPoint drops fully-transparent stop colors from the color ramp when ≥2 visible
    // stops remain: the whole fade keeps the navy hue and only alpha ramps (probe-measured).
    const first = stops[1] as string
    expect(first).toBe('rgba(41,53,77,0)')
    const mid = stops[stops.indexOf(0.25) + 1] as string
    const [r0, g0, b0, a0] = mid.match(/[\d.]+/g)!.map(Number)
    expect(Math.abs(r0 - 41)).toBeLessThanOrEqual(1)
    expect(Math.abs(g0 - 53)).toBeLessThanOrEqual(1)
    expect(Math.abs(b0 - 77)).toBeLessThanOrEqual(1)
    expect(a0).toBeCloseTo(0.5, 2)
  })

  it('keeps equal-color stop pairs unsubdivided', () => {
    const r = fillToKonva(
      {
        kind: 'gradient',
        angleDeg: 0,
        stops: [
          { pos: 0, color: '#123456' },
          { pos: 1, color: '#123456' },
        ],
      } as any,
      100,
      100,
    )
    expect(r.fillLinearGradientColorStops!.length).toBe(4)
  })
})
