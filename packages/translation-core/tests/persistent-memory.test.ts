import { describe, expect, it } from 'vitest'

import { PersistentTranslationMemory } from '../src/persistent-memory'
import type { PersistentMemoryFileSystem } from '../src/persistent-memory'

class MemoryFS implements PersistentMemoryFileSystem {
  files = new Map<string, string>()
  async mkdir(_p: string) {}
  async readFile(p: string) {
    const v = this.files.get(p)
    if (v === undefined) {
      const err = new Error('ENOENT') as NodeJS.ErrnoException
      err.code = 'ENOENT'
      throw err
    }
    return v
  }
  async writeFile(p: string, contents: string) {
    this.files.set(p, contents)
  }
  async rename(from: string, to: string) {
    const v = this.files.get(from)
    if (v !== undefined) {
      this.files.set(to, v)
      this.files.delete(from)
    }
  }
  async readdir(p: string) {
    const prefix = p.endsWith('/') ? p : p + '/'
    return Array.from(this.files.keys())
      .filter((k) => k.startsWith(prefix))
      .filter((k) => !k.endsWith('.tmp'))
  }
}

describe('PersistentTranslationMemory', () => {
  it('round-trips an exact entry through save / flush / load', async () => {
    const fs = new MemoryFS()
    const tm = new PersistentTranslationMemory({ baseDir: '/tmp/tm', fileSystem: fs })
    tm.save({ sourceLang: 'en-US', targetLang: 'zh-CN', sourceText: 'Hello', translatedText: '你好' })
    expect(tm.isDirty()).toBe(true)
    await tm.flush()
    expect(tm.isDirty()).toBe(false)

    const reloaded = new PersistentTranslationMemory({ baseDir: '/tmp/tm', fileSystem: fs })
    await reloaded.load()
    const hit = reloaded.lookup('en-US', 'zh-CN', 'Hello')
    expect(hit?.translatedText).toBe('你好')
    expect(reloaded.size()).toBe(1)
  })

  it('fuzzy lookup finds near matches with confidence < 1', () => {
    const tm = new PersistentTranslationMemory({ baseDir: '/tmp/tm', fileSystem: new MemoryFS() })
    tm.save({
      sourceLang: 'en-US',
      targetLang: 'zh-CN',
      sourceText: 'machine wash cold',
      translatedText: '冷水机洗',
    })
    const hit = tm.fuzzyLookup('en-US', 'zh-CN', 'machine wash, cold')
    expect(hit).not.toBeNull()
    // Default threshold is 0.7 (matches LumosAI), so the hit must clear it.
    expect(hit?.confidence).toBeGreaterThan(0.7)
    // And it must NOT be an exact match (comma vs no comma).
    expect(hit?.confidence).toBeLessThan(1)
    expect(hit?.translatedText).toBe('冷水机洗')
  })

  it('fuzzy lookup returns confidence=1 for exact matches', () => {
    const tm = new PersistentTranslationMemory({ baseDir: '/tmp/tm', fileSystem: new MemoryFS() })
    tm.save({
      sourceLang: 'en-US',
      targetLang: 'zh-CN',
      sourceText: 'Hello',
      translatedText: '你好',
    })
    const hit = tm.fuzzyLookup('en-US', 'zh-CN', 'Hello')
    expect(hit?.confidence).toBe(1)
  })

  it('fuzzy lookup returns null when below threshold', () => {
    const tm = new PersistentTranslationMemory({ baseDir: '/tmp/tm', fileSystem: new MemoryFS(), fuzzyThreshold: 0.95 })
    tm.save({
      sourceLang: 'en-US',
      targetLang: 'zh-CN',
      sourceText: 'completely different topic',
      translatedText: '完全不相关的主题',
    })
    const hit = tm.fuzzyLookup('en-US', 'zh-CN', 'Hello world')
    expect(hit).toBeNull()
  })

  it('isolates pairs into separate on-disk files', async () => {
    const fs = new MemoryFS()
    const tm = new PersistentTranslationMemory({ baseDir: '/tmp/tm', fileSystem: fs })
    tm.save({ sourceLang: 'en-US', targetLang: 'zh-CN', sourceText: 'A', translatedText: '甲' })
    tm.save({ sourceLang: 'en-US', targetLang: 'ja-JP', sourceText: 'A', translatedText: 'エー' })
    tm.save({ sourceLang: 'fr-FR', targetLang: 'zh-CN', sourceText: 'A', translatedText: '啊' })
    await tm.flush()
    const reloaded = new PersistentTranslationMemory({ baseDir: '/tmp/tm', fileSystem: fs })
    await reloaded.load()
    expect(reloaded.lookup('en-US', 'zh-CN', 'A')?.translatedText).toBe('甲')
    expect(reloaded.lookup('en-US', 'ja-JP', 'A')?.translatedText).toBe('エー')
    expect(reloaded.lookup('fr-FR', 'zh-CN', 'A')?.translatedText).toBe('啊')
    expect(reloaded.size()).toBe(3)
  })

  it('exportTmx() produces valid 1.4b XML', () => {
    const tm = new PersistentTranslationMemory({ baseDir: '/tmp/tm', fileSystem: new MemoryFS() })
    tm.save({ sourceLang: 'en-US', targetLang: 'zh-CN', sourceText: 'Hello', translatedText: '你好' })
    tm.save({ sourceLang: 'en-US', targetLang: 'zh-CN', sourceText: 'World', translatedText: '世界' })
    const xml = tm.exportTmx()
    expect(xml).toContain('<?xml version="1.0" encoding="UTF-8"?>')
    expect(xml).toContain('<tmx version="1.4b">')
    expect(xml.match(/<tu>/g)).toHaveLength(2)
    expect(xml).toContain('<tuv xml:lang="en-US">')
    expect(xml).toContain('<tuv xml:lang="zh-CN">')
    // XML escaping — make sure & and < round-trip safely.
    tm.save({
      sourceLang: 'en-US',
      targetLang: 'zh-CN',
      sourceText: 'A & B',
      translatedText: '<C>',
    })
    const xml2 = tm.exportTmx()
    expect(xml2).toContain('A &amp; B')
    expect(xml2).toContain('&lt;C&gt;')
  })

  it('isDirty tracks per-pair changes', () => {
    const tm = new PersistentTranslationMemory({ baseDir: '/tmp/tm', fileSystem: new MemoryFS() })
    expect(tm.isDirty()).toBe(false)
    tm.save({ sourceLang: 'en-US', targetLang: 'zh-CN', sourceText: 'A', translatedText: '甲' })
    expect(tm.isDirty()).toBe(true)
  })

  it('flush is a no-op when nothing changed', async () => {
    const fs = new MemoryFS()
    const tm = new PersistentTranslationMemory({ baseDir: '/tmp/tm', fileSystem: fs })
    await tm.flush()
    expect(fs.files.size).toBe(0)
  })
})
