import { describe, expect, it } from 'vitest'
import { bodyContextOf } from '../src/analyze/blocks'

const mk = (n: number) =>
  Array.from({ length: n }, (_, i) => ({
    box: { x0: i, y0: i, x1: i + 1, y1: i + 1 },
    text: 'x',
    fontSize: 12,
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
  })) as never[]

describe('bodyContextOf without spreads', () => {
  it('handles a 200k-line page that used to blow the argument limit', () => {
    // Math.min(...xs) threw RangeError past V8's argument limit (~124.6k here);
    // pages with that many lines exist in CAD/map exports.
    const ctx = bodyContextOf(mk(200_000) as never)
    expect(ctx.bodyLeft).toBe(0)
    expect(ctx.bodyRight).toBe(200_000)
  })

  it('keeps small-page results identical', () => {
    const ctx = bodyContextOf(mk(10) as never)
    expect(ctx).toEqual({ bodyLeft: 0, bodyRight: 10 })
  })
})
