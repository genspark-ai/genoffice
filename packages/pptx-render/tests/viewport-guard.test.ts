import { describe, it, expect } from 'vitest'
import { makeViewport } from '../src/coords'

const DEFAULT_CX = 9144000
const DEFAULT_CY = 6858000

function assertFiniteViewport(vp: { widthPx: number; heightPx: number; scale: number }) {
  for (const v of [vp.widthPx, vp.heightPx, vp.scale]) {
    expect(Number.isFinite(v)).toBe(true)
    expect(v).toBeGreaterThan(0)
  }
}

describe('makeViewport degenerate guard', () => {
  it('normal path is unchanged', () => {
    const vp = makeViewport({ cx: 12192000, cy: 6858000 }, 1280)
    expect(vp.widthPx).toBe(1280)
    expect(vp.heightPx).toBeCloseTo(720, 0)
    expect(vp.scale).toBeCloseTo(1, 6)
  })

  it('zero cx falls back to the default slide width', () => {
    const guarded = makeViewport({ cx: 0, cy: DEFAULT_CY }, 960)
    const expected = makeViewport({ cx: DEFAULT_CX, cy: DEFAULT_CY }, 960)
    expect(guarded).toEqual(expected)
    assertFiniteViewport(guarded)
  })

  it('negative, NaN, and Infinity inputs stay finite', () => {
    const bad = [-1, -9144000, 0, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY]
    for (const cx of bad) {
      for (const cy of bad) {
        for (const fit of bad) {
          const vp = makeViewport({ cx, cy }, fit)
          assertFiniteViewport(vp)
        }
      }
    }
  })

  it('degenerate fit width falls back to scale 1', () => {
    const vp = makeViewport({ cx: DEFAULT_CX, cy: DEFAULT_CY }, 0)
    expect(vp.widthPx).toBeCloseTo(DEFAULT_CX / 9525, 6)
    expect(vp.heightPx).toBeCloseTo(DEFAULT_CY / 9525, 6)
    expect(vp.scale).toBeCloseTo(1, 6)
  })
})
