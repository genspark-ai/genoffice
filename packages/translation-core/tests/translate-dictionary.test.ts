/**
 * Unit coverage for the KB + LLM dictionary builder.
 *
 * The model call is injected, so these tests pin the parts that would
 * otherwise only be observable after a real provider round-trip: segment
 * mining, batching, KB seeding, and the post-translation KB rule pass.
 */
import { describe, expect, it } from 'vitest'
import { mkdtempSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { tmpdir } from 'node:os'

import {
  applyKbRules,
  batchSegments,
  buildDictionary,
  defaultDictionaryPath,
  mineSegments,
} from '../src/dictionary'
import { KnowledgeBase } from '../src/knowledge-base'

describe('mineSegments', () => {
  it('splits lines, collapses whitespace and dedupes', () => {
    const segments = mineSegments(
      ['Product Specification', '  Product   Specification  ', '', 'Fabric: 100% cotton', '180'].join('\n'),
    )
    expect(segments).toEqual(['Product Specification', 'Fabric: 100% cotton'])
  })

  it('drops punctuation-only and too-short lines', () => {
    expect(mineSegments('a\n--\n|||\nhello')).toEqual(['hello'])
  })

  it('honours the minChars threshold', () => {
    expect(mineSegments('ab\nabc\nabcd', 3)).toEqual(['abc', 'abcd'])
  })
})

describe('batchSegments', () => {
  it('groups under the char budget', () => {
    const batches = batchSegments(['aaaa', 'bbbb', 'cccc'], 9)
    expect(batches).toEqual([['aaaa', 'bbbb'], ['cccc']])
  })

  it('keeps a single oversized segment in its own batch', () => {
    const batches = batchSegments(['x'.repeat(10)], 4)
    expect(batches).toEqual([['x'.repeat(10)]])
  })
})

describe('defaultDictionaryPath', () => {
  it('slugs the source name into the dictionary directory', () => {
    const p = defaultDictionaryPath('/data', '/tmp/My Spec (v2).docx')
    expect(p.startsWith('/data/translation-dictionaries/')).toBe(true)
    expect(p).toContain('My_Spec__v2_')
    expect(p.endsWith('.json')).toBe(true)
  })
})

describe('applyKbRules', () => {
  const seed = (kb: KnowledgeBase) => {
    kb.upsert({
      id: 't1',
      scope: 'company',
      priority: 3,
      sourceTerm: '克重',
      targetTerm: 'GSM',
    })
    kb.upsert({
      id: 'f1',
      scope: 'company',
      priority: 3,
      forbiddenText: 'all cotton',
      replacement: '100% cotton',
    })
  }

  it('rewrites forbidden phrasing to its replacement', () => {
    const kb = new KnowledgeBase()
    seed(kb)
    const { text } = applyKbRules('this is all cotton fabric', kb, {
      sourceLang: 'en-US',
      targetLang: 'zh-CN',
    })
    expect(text).toBe('this is 100% cotton fabric')
  })

  it('drops forbidden phrasing when no replacement is configured', () => {
    const kb = new KnowledgeBase()
    kb.upsert({
      id: 'f2',
      scope: 'company',
      priority: 3,
      forbiddenText: 'cheap quality',
    })
    const { text } = applyKbRules('a cheap quality shirt', kb, {
      sourceLang: 'en-US',
      targetLang: 'zh-CN',
    })
    expect(text).toBe('a  shirt')
  })

  it('forces mandatory terms and reports which ones matched', () => {
    const kb = new KnowledgeBase()
    seed(kb)
    const { text, matchedTerms } = applyKbRules('克重 180', kb, {
      sourceLang: 'zh-CN',
      targetLang: 'en-US',
    })
    expect(text).toBe('GSM 180')
    expect(matchedTerms).toContain('克重')
  })

  it('leaves text alone when nothing matches', () => {
    const kb = new KnowledgeBase()
    seed(kb)
    const { text, matchedTerms } = applyKbRules('nothing here', kb, {
      sourceLang: 'en-US',
      targetLang: 'zh-CN',
    })
    expect(text).toBe('nothing here')
    expect(matchedTerms).toEqual([])
  })
})

describe('buildDictionary', () => {
  function writeInput(contents: string): string {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-'))
    const path = join(dir, 'input.txt')
    writeFileSync(path, contents, 'utf8')
    return path
  }

  it('writes a sorted dictionary from the injected translator', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-out-'))
    const inputPath = writeInput('Hello world\nGood morning\nHello world\n')
    const result = await buildDictionary(
      {
        inputPath,
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        outputPath: join(dir, 'dict.json'),
        dataDir: dir,
      },
      {
        translateBatch: async ({ units }) => ({
          ok: true,
          units: units.map((u) => ({
            unitId: u.unitId,
            sourceText: u.sourceText,
            translatedText: u.sourceText === 'Hello world' ? '你好世界' : '早上好',
            status: 'translated' as const,
          })),
        }),
      },
    )
    expect(result.ok).toBe(true)
    expect(result.totalSegments).toBe(2)
    expect(result.llmEntries).toBe(2)
    const written = JSON.parse(readFileSync(result.dictionaryPath!, 'utf8')) as Record<string, string>
    expect(Object.keys(written)).toEqual(['Good morning', 'Hello world'])
    expect(written['Hello world']).toBe('你好世界')
    expect(result.segments).toEqual([
      { source: 'Hello world', target: '你好世界', origin: 'llm' },
      { source: 'Good morning', target: '早上好', origin: 'llm' },
    ])
  })

  it('records segments the model failed on instead of dropping them silently', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-miss-'))
    const inputPath = writeInput('Translate me\nMe too\n')
    const result = await buildDictionary(
      {
        inputPath,
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        outputPath: join(dir, 'dict.json'),
        dataDir: dir,
      },
      {
        translateBatch: async ({ units }) => ({
          ok: true,
          units: units.map((u, i) => ({
            unitId: u.unitId,
            sourceText: u.sourceText,
            ...(i === 0
              ? { status: 'failed' as const, errorMessage: 'provider down' }
              : { translatedText: '我也一样', status: 'translated' as const }),
          })),
        }),
      },
    )
    expect(result.ok).toBe(true)
    expect(result.missed).toEqual(['Translate me'])
    expect(result.llmEntries).toBe(1)
  })

  it('surfaces a provider failure as an error result instead of throwing', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-err-'))
    const inputPath = writeInput('Hello\n')
    const result = await buildDictionary(
      {
        inputPath,
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        outputPath: join(dir, 'dict.json'),
        dataDir: dir,
      },
      {
        translateBatch: async () => {
          throw new Error('socket hang up')
        },
      },
    )
    expect(result.ok).toBe(false)
    expect(result.error).toContain('socket hang up')
  })

  it('skips the model entirely when useLlm is false', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-nolm-'))
    const inputPath = writeInput('Hello\nWorld\n')
    let called = false
    const result = await buildDictionary(
      {
        inputPath,
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        outputPath: join(dir, 'dict.json'),
        dataDir: dir,
        useLlm: false,
      },
      {
        translateBatch: async () => {
          called = true
          return { ok: true, units: [] }
        },
      },
    )
    expect(called).toBe(false)
    expect(result.ok).toBe(true)
    expect(result.llmEntries).toBe(0)
  })

  it('rejects a file with no extractable text', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-empty-'))
    const inputPath = join(dir, 'empty.txt')
    writeFileSync(inputPath, '\n\n   \n', 'utf8')
    const result = await buildDictionary(
      { inputPath, sourceLang: 'en-US', targetLang: 'zh-CN', dataDir: dir },
      { translateBatch: async () => ({ ok: true, units: [] }) },
    )
    expect(result.ok).toBe(false)
    expect(result.error).toContain('no translatable segments')
  })

  it('reports a missing input file', async () => {
    const result = await buildDictionary(
      { inputPath: '/nope/missing.txt', sourceLang: 'en-US', targetLang: 'zh-CN', dataDir: '/tmp' },
      { translateBatch: async () => ({ ok: true, units: [] }) },
    )
    expect(result.ok).toBe(false)
  })
})
