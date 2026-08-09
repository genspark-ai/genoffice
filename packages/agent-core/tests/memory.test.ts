import { describe, expect, it } from 'vitest'

import {
  MAX_MEMORIES,
  MAX_MEMORY_BUDGET_CHARS,
  MAX_MEMORY_CHARS,
  buildMemoryPrompt,
  memoriesWithinBudget,
  normalizeMemoryText,
  type UserMemory,
} from '../src/memory'

const at = (createdAt: number, text: string, id = `m${createdAt}`): UserMemory => ({
  id,
  text,
  createdAt,
})

describe('normalizeMemoryText', () => {
  it('collapses whitespace so a pasted block stays one line in the prompt', () => {
    expect(normalizeMemoryText('  prefers\n\n  short   sentences ')).toBe('prefers short sentences')
  })

  it('caps a single memory: anything longer belongs in a rule or a skill', () => {
    expect(normalizeMemoryText('x'.repeat(MAX_MEMORY_CHARS + 50))).toHaveLength(MAX_MEMORY_CHARS)
  })

  it('returns empty for nothing worth storing', () => {
    for (const value of ['', '   ', null, undefined]) {
      expect(normalizeMemoryText(value)).toBe('')
    }
  })
})

describe('memoriesWithinBudget', () => {
  it('orders newest first, so a recent preference wins over a stale one', () => {
    const kept = memoriesWithinBudget([at(1, 'old'), at(3, 'newest'), at(2, 'middle')])
    expect(kept.map((m) => m.text)).toEqual(['newest', 'middle', 'old'])
  })

  it('drops the oldest once the character budget is spent', () => {
    // each entry is a fifth of the budget, so only five can fit
    const size = Math.floor(MAX_MEMORY_BUDGET_CHARS / 5)
    const many = Array.from({ length: 12 }, (_, i) => at(i, 'x'.repeat(size), `m${i}`))
    const kept = memoriesWithinBudget(many)

    expect(kept.length).toBeLessThan(12)
    const spent = kept.reduce((n, m) => n + m.text.length + 3, 0)
    expect(spent).toBeLessThanOrEqual(MAX_MEMORY_BUDGET_CHARS)
    // the survivors are the newest ones
    expect(kept[0]!.id).toBe('m11')
  })

  it('caps the count even when every entry is tiny', () => {
    const many = Array.from({ length: MAX_MEMORIES + 40 }, (_, i) => at(i, 'a', `m${i}`))
    expect(memoriesWithinBudget(many)).toHaveLength(MAX_MEMORIES)
  })
})

describe('buildMemoryPrompt', () => {
  it('says nothing at all when there is nothing remembered', () => {
    expect(buildMemoryPrompt([])).toBe('')
  })

  it('lists the memories and tells the model the user can override them', () => {
    const prompt = buildMemoryPrompt([at(2, 'writes in British English'), at(1, 'dislikes emoji')])
    expect(prompt).toContain('- writes in British English')
    expect(prompt).toContain('- dislikes emoji')
    expect(prompt.toLowerCase()).toContain('override')
  })

  it('never exceeds the budget it promises', () => {
    const many = Array.from({ length: 200 }, (_, i) => at(i, 'y'.repeat(200), `m${i}`))
    // heading and preamble aside, the entries themselves stay inside the cap
    const entryChars = buildMemoryPrompt(many)
      .split('\n')
      .filter((l) => l.startsWith('- '))
      .reduce((n, l) => n + l.length + 1, 0)
    expect(entryChars).toBeLessThanOrEqual(MAX_MEMORY_BUDGET_CHARS)
  })
})
