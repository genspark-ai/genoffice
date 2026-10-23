import { describe, expect, it } from 'vitest'
import { defaultMaxLiveTabs, pickTabsToSleep } from '../src/main/tab-hibernation'

const GB = 2 ** 30
const policy = { maxLiveTabs: 2, idleMs: 10_000 }

describe('defaultMaxLiveTabs', () => {
  it('scales the live-tab cap with physical memory', () => {
    expect(defaultMaxLiveTabs(8 * GB)).toBe(4)
    expect(defaultMaxLiveTabs(16 * GB)).toBe(8)
    expect(defaultMaxLiveTabs(32 * GB)).toBe(12)
    expect(defaultMaxLiveTabs(64 * GB)).toBe(12)
  })
})

describe('pickTabsToSleep', () => {
  const now = 100_000
  const idle = (id: string, ago: number) => ({ id, lastActiveAt: now - ago })

  it('returns nothing while the live count is within the cap', () => {
    expect(pickTabsToSleep([idle('a', 60_000)], 2, policy, now)).toEqual([])
  })

  it('sleeps only as many idle tabs as exceed the cap, oldest first', () => {
    const picked = pickTabsToSleep(
      [idle('fresh', 1_000), idle('old', 60_000), idle('older', 90_000)],
      4,
      policy,
      now,
    )
    expect(picked).toEqual(['older', 'old'])
  })

  it('a tab idle for less than idleMs is never picked even when the cap is exceeded', () => {
    expect(pickTabsToSleep([idle('a', 5_000), idle('b', 2_000)], 5, policy, now)).toEqual([])
  })

  it('over budget the single oldest candidate goes regardless of cap and idleness', () => {
    expect(pickTabsToSleep([idle('a', 1_000), idle('b', 3_000)], 1, policy, now, true)).toEqual([
      'b',
    ])
    expect(pickTabsToSleep([], 1, policy, now, true)).toEqual([])
  })
})
