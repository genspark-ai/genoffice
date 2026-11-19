import { describe, expect, it } from 'vitest'
import { cubicBezier } from '../src/renderer/video-export/transitions'

describe('cubicBezier', () => {
  it('pins the ends and is monotonic', () => {
    const ease = cubicBezier(0.25, 0.1, 0.25, 1)
    expect(ease(0)).toBe(0)
    expect(ease(1)).toBe(1)
    let prev = 0
    for (let i = 1; i <= 20; i++) {
      const v = ease(i / 20)
      expect(v).toBeGreaterThanOrEqual(prev)
      prev = v
    }
  })

  it('matches the CSS ease curve at its midpoint', () => {
    // cubic-bezier(0.25, 0.1, 0.25, 1) at x = 0.5 is ≈ 0.8024
    expect(cubicBezier(0.25, 0.1, 0.25, 1)(0.5)).toBeCloseTo(0.8024, 3)
  })

  it('ease-out stays above linear', () => {
    const out = cubicBezier(0, 0, 0.58, 1)
    expect(out(0.3)).toBeGreaterThan(0.3)
  })
})
