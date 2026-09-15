import { describe, expect, it } from 'vitest'

import { TranslationMemory } from '../src/memory'

describe('TranslationMemory', () => {
  it('round-trips save / lookup by source text', () => {
    const mem = new TranslationMemory()
    mem.save({
      sourceLang: 'en-US',
      targetLang: 'zh-CN',
      sourceText: 'Hello',
      translatedText: '你好',
    })
    const hit = mem.lookup('en-US', 'zh-CN', 'Hello')
    expect(hit?.translatedText).toBe('你好')
  })

  it('treats whitespace-only variants of the source as the same key', () => {
    const mem = new TranslationMemory()
    mem.save({
      sourceLang: 'en-US',
      targetLang: 'zh-CN',
      sourceText: 'Hello world',
      translatedText: '你好世界',
    })
    expect(mem.lookup('en-US', 'zh-CN', '  Hello   world  ')?.translatedText).toBe('你好世界')
  })

  it('returns null for missing entries', () => {
    const mem = new TranslationMemory()
    expect(mem.lookup('en-US', 'zh-CN', 'anything')).toBeNull()
  })

  it('saveMany counts saved vs skipped units', () => {
    const mem = new TranslationMemory()
    const r = mem.saveMany({
      scene: 'doc',
      sourceLang: 'en-US',
      targetLang: 'zh-CN',
      units: [
        { unitId: 'a', sourceText: 'Hello', translatedText: '你好' },
        { unitId: 'b', sourceText: 'World', translatedText: '' },
        { unitId: 'c', sourceText: '   ', translatedText: ' ' },
      ],
    })
    expect(r.ok).toBe(true)
    expect(r.savedCount).toBe(1)
    expect(r.skippedCount).toBe(2)
  })

  it('evicts the oldest entries when capacity is exceeded', () => {
    const mem = new TranslationMemory({ maxEntries: 10 })
    for (let i = 0; i < 25; i++) {
      mem.save({
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        sourceText: `s-${i}`,
        translatedText: `t-${i}`,
      })
    }
    expect(mem.size()).toBeLessThanOrEqual(10)
  })
})
