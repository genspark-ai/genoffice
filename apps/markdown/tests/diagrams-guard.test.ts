import { describe, expect, it, vi } from 'vitest'
import {
  DIAGRAM_MAX_CANVAS_DIM_PX,
  clampDiagramMaxWidth,
  computeDiagramRasterSize,
  parseSvgViewBox,
  pinSvgIntrinsicSize,
} from '../src/renderer/editor/diagrams'
import {
  isWaveJson,
  prevalidateWaveSource,
  renderWavedrom,
  scopeSkinStyles,
} from '../src/renderer/editor/wavedrom'

describe('diagram viewBox guards', () => {
  it('parses comma-separated viewBox values', () => {
    expect(parseSvgViewBox('<svg viewBox="0,0,200,100">')).toEqual({ width: 200, height: 100 })
    expect(parseSvgViewBox('<svg viewBox="0 0 200,100">')).toEqual({ width: 200, height: 100 })
    expect(parseSvgViewBox("<svg viewBox='0, 0, 64, 48'>")).toEqual({ width: 64, height: 48 })
    expect(parseSvgViewBox('<svg viewBox="0 0 200 100">')).toEqual({ width: 200, height: 100 })
  })

  it('returns null for a missing or invalid viewBox', () => {
    expect(parseSvgViewBox('<svg width="100" height="50">')).toBeNull()
    expect(parseSvgViewBox('<svg viewBox="0 0 0 100">')).toBeNull()
    expect(parseSvgViewBox('<svg viewBox="0 0 -5 100">')).toBeNull()
    expect(computeDiagramRasterSize(0, 100, 620)).toBeNull()
    expect(computeDiagramRasterSize(Number.NaN, 100, 620)).toBeNull()
  })

  it('strips single-quoted width/height when pinning intrinsic size', () => {
    const out = pinSvgIntrinsicSize(
      "<svg width='100%' height='50' viewBox='0 0 200 100'>",
      200,
      100,
    )
    expect(out).toContain('width="200"')
    expect(out).toContain('height="100"')
    expect(out).not.toContain('100%')
  })
})

describe('diagram size caps', () => {
  it('clamps maxWidthPx to a positive finite value within a sane cap', () => {
    expect(clampDiagramMaxWidth(Number.NaN)).toBe(620)
    expect(clampDiagramMaxWidth(-5)).toBe(620)
    expect(clampDiagramMaxWidth(0)).toBe(620)
    expect(clampDiagramMaxWidth(Number.POSITIVE_INFINITY)).toBe(620)
    expect(clampDiagramMaxWidth(300)).toBe(300)
    expect(clampDiagramMaxWidth(100_000)).toBeLessThanOrEqual(2048)
  })

  it('caps the backing canvas so huge diagrams cannot allocate unbounded bitmaps', () => {
    const small = computeDiagramRasterSize(100, 80, 620)
    expect(small).toEqual({ widthPx: 100, heightPx: 80, canvasWidth: 200, canvasHeight: 160 })
    const huge = computeDiagramRasterSize(5000, 5000, 620)!
    expect(Math.max(huge.canvasWidth, huge.canvasHeight)).toBe(DIAGRAM_MAX_CANVAS_DIM_PX)
    expect(huge.canvasWidth).toBeLessThanOrEqual(DIAGRAM_MAX_CANVAS_DIM_PX)
    expect(huge.canvasHeight).toBeLessThanOrEqual(DIAGRAM_MAX_CANVAS_DIM_PX)
    // doc size still fits the clamped max width
    expect(huge.widthPx).toBeLessThanOrEqual(620)
  })
})

describe('wavedrom guards', () => {
  it('prevalidates the wave source with an actionable message', () => {
    expect(prevalidateWaveSource('')).toMatch(/empty/)
    expect(prevalidateWaveSource('   ')).toMatch(/empty/)
    expect(prevalidateWaveSource('[1, 2]')).toMatch(/starting with "\{"/)
    expect(prevalidateWaveSource(42)).toMatch(/WaveJSON/)
    expect(prevalidateWaveSource('{ signal: [] }')).toBeNull()
  })

  it('rejects empty WaveJSON arrays as non-WaveJSON', () => {
    expect(isWaveJson({ signal: [] })).toBe(false)
    expect(isWaveJson({ signal: [{ wave: 'p...' }] })).toBe(true)
    expect(isWaveJson({ signal: ['oops'] })).toBe(false)
    expect(isWaveJson({ foo: 1 })).toBe(false)
  })

  it('returns ok:false instead of throwing when the renderer fails to load', async () => {
    const result = await renderWavedrom("{ signal: [{ name: 'clk', wave: 'p...' }] }", {
      load: () => Promise.reject(new Error('boom')),
    })
    expect(result.ok).toBe(false)
    if (!result.ok) expect(result.error).toMatch(/failed to load.*boom/)
  })

  it('does not pay for the bundle when prevalidation fails', async () => {
    const load = vi.fn(() => Promise.reject(new Error('must not be called')))
    const result = await renderWavedrom('   ', { load })
    expect(result.ok).toBe(false)
    expect(load).not.toHaveBeenCalled()
  })

  it('scopes every style block and never double-prefixes wd- classes', () => {
    const svg =
      '<svg class="WaveDrom">' +
      '<style>.s1{fill:red}text{fill:blue}</style>' +
      '<style>.s2{fill:green}</style>' +
      '<rect class="s1"/><rect class="wd-s2"/>' +
      '</svg>'
    const out = scopeSkinStyles(svg)
    expect(out).toContain('.wd-s1{')
    expect(out).toContain('.wd-s2{')
    expect(out).toContain('svg.WaveDrom text{')
    expect(out).not.toContain('wd-wd-')
    expect(out).toContain('class="wd-s1"')
    expect(out).toContain('class="wd-s2"')
    // both style blocks were rewritten (global flag, not first-only)
    expect(out.match(/\.wd-s2\{/g)?.length).toBe(1)
  })
})
