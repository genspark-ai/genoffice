import { describe, expect, it } from 'vitest'
import { mapBounded } from '../src/main/map-bounded'

describe('mapBounded', () => {
  it('keeps the order of the items and never runs more than the limit at once', async () => {
    let running = 0
    let most = 0
    const out = await mapBounded([5, 1, 4, 2, 3, 0, 6], 3, async (n) => {
      most = Math.max(most, ++running)
      await new Promise((resolve) => setTimeout(resolve, n))
      running--
      return n * 2
    })
    expect(out).toEqual([10, 2, 8, 4, 6, 0, 12])
    expect(most).toBe(3)
  })

  it('makes nothing of nothing', async () => {
    expect(await mapBounded([], 4, async () => 1)).toEqual([])
  })
})
