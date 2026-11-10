import { describe, expect, it } from 'vitest'
import { createRenderContext } from '../src/generate/render-context'

describe('createRenderContext degenerate inputs', () => {
  it('matches defaults for honest settings', () => {
    const ctx = createRenderContext({})
    expect(ctx.pageWidthDxa).toBe(794 * 15)
    expect(ctx.contentDxa).toBeGreaterThan(0)
  })

  it('degrades non-finite page size, squeeze, and margins to finite geometry', () => {
    const ctx = createRenderContext({
      pageSizePx: { width: Infinity, height: NaN },
      pageFitSqueeze: Infinity,
      viewportWidthPx: -100,
      marginsPx: { top: NaN, left: Infinity, right: -50 },
    })
    for (const v of [
      ctx.pageWidthDxa,
      ctx.pageHeightDxa,
      ctx.contentDxa,
      ctx.pageMargins.top,
      ctx.pageMargins.left,
      ctx.pageMargins.right,
      ctx.measurementScale,
    ]) {
      expect(Number.isFinite(v)).toBe(true)
    }
    expect(ctx.pxToTwips(NaN)).toBe(0)
    expect(ctx.pxToTwips(10)).toBeGreaterThan(0)
  })

  it('falls back to A4 for a zero page size (display:none root)', () => {
    const ctx = createRenderContext({ pageSizePx: { width: 0, height: 0 } })
    expect(ctx.pageWidthDxa).toBe(794 * 15)
    expect(ctx.pageHeightDxa).toBe(1123 * 15)
  })

  it('floors the content width when the horizontal margins overrun the page', () => {
    // each margin is clamped on its own, but two healthy ones can still add up
    // past the page: the negative width then scales every floating image to a
    // negative extent, which Word reports as a corrupt file
    const ctx = createRenderContext({ marginsPx: { left: 700, right: 700 } })
    expect(ctx.contentDxa).toBeGreaterThan(0)
  })
})
