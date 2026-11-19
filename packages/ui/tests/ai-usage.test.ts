import { describe, expect, it } from 'vitest'
import { addTokenUsage, cacheHitPercent, formatTokenCount } from '../src/AiUsageChip'
import { AI_USAGE_LABELS } from '../src/strings-ai-usage'

describe('formatTokenCount', () => {
  it('keeps small counts verbatim', () => {
    expect(formatTokenCount(0)).toBe('0')
    expect(formatTokenCount(999)).toBe('999')
  })

  it('abbreviates thousands and millions', () => {
    expect(formatTokenCount(1000)).toBe('1k')
    expect(formatTokenCount(1234)).toBe('1.2k')
    expect(formatTokenCount(1234567)).toBe('1.2M')
    expect(formatTokenCount(2_000_000)).toBe('2M')
  })
})

describe('addTokenUsage', () => {
  it('sums both sides across turns', () => {
    expect(
      addTokenUsage(addTokenUsage(undefined, { inputTokens: 100, outputTokens: 10 }), {
        inputTokens: 250,
        outputTokens: 35,
      }),
    ).toEqual({ inputTokens: 350, outputTokens: 45 })
  })

  it('keeps sides the provider never reported absent instead of zeroing them', () => {
    expect(addTokenUsage(undefined, { inputTokens: 5 })).toEqual({ inputTokens: 5 })
    expect(addTokenUsage({ inputTokens: 5 }, { outputTokens: 7 })).toEqual({
      inputTokens: 5,
      outputTokens: 7,
    })
  })

  it('does not mutate the accumulator', () => {
    const prev = { inputTokens: 1, outputTokens: 1 }
    addTokenUsage(prev, { inputTokens: 2, outputTokens: 2 })
    expect(prev).toEqual({ inputTokens: 1, outputTokens: 1 })
  })
})

describe('AI_USAGE_LABELS', () => {
  it('covers every ui language', () => {
    expect(Object.keys(AI_USAGE_LABELS).length).toBeGreaterThanOrEqual(21)
    for (const [lang, labels] of Object.entries(AI_USAGE_LABELS)) {
      expect(labels.runTitle.length, lang).toBeGreaterThan(0)
      expect(labels.sessionTitle.length, lang).toBeGreaterThan(0)
      expect(labels.cacheHit, lang).toContain('{pct}')
    }
  })
})

describe('cacheHitPercent', () => {
  it('is the cached share of the input', () => {
    expect(cacheHitPercent({ inputTokens: 1000, cacheReadTokens: 820 })).toBe(82)
    expect(cacheHitPercent({ inputTokens: 1000, cacheReadTokens: 0 })).toBe(0)
  })

  it('is unknown when the provider reports no cache figure', () => {
    expect(cacheHitPercent({ inputTokens: 1000 })).toBeNull()
    expect(cacheHitPercent({ cacheReadTokens: 5 })).toBeNull()
  })

  it('sums cache figures across turns', () => {
    expect(
      addTokenUsage(
        { inputTokens: 100, cacheReadTokens: 0, cacheWriteTokens: 90 },
        { inputTokens: 120, cacheReadTokens: 90 },
      ),
    ).toEqual({ inputTokens: 220, cacheReadTokens: 90, cacheWriteTokens: 90 })
  })
})
