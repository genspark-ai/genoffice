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
  fillDictionaryGaps,
  mineSegments,
  readDictionaryFile,
  segmentFormatForPath,
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

  it('strips the docx heading and list markers the extractor adds', () => {
    // The handlers substitute on runs, so `# 供应商交付说明` can never match the
    // `供应商交付说明` run — mining the decorated text left every heading and list
    // item in the source language while coverage claimed they were done.
    const segments = mineSegments('# 供应商交付说明\n- 第一项\n## 二级标题', 2, 'docx')
    expect(segments).toEqual(['供应商交付说明', '第一项', '二级标题'])
  })

  it('splits docx table rows into the cells the handler actually rewrites', () => {
    expect(mineSegments('供应商 | 数量 | 交期', 2, 'docx')).toEqual(['供应商', '数量', '交期'])
  })

  it('drops the slide headers the pptx extractor synthesizes', () => {
    const text = '## Slide 1\n封面标题\n正文一段\n\n## Slide 2\n结论'
    expect(mineSegments(text, 2, 'pptx')).toEqual(['封面标题', '正文一段', '结论'])
  })

  it('drops the sheet headers the xlsx extractor synthesizes', () => {
    const text = '# 报价明细\n物料 | 单价\n\n# 汇总'
    expect(mineSegments(text, 2, 'xlsx')).toEqual(['物料', '单价'])
  })

  it('leaves flat formats untouched', () => {
    // A pdf line is real text; splitting it on `|` would invent keys the
    // document does not contain.
    expect(mineSegments('# 看起来像标题\na | b', 2, 'flat')).toEqual(['# 看起来像标题', 'a | b'])
  })

  it('maps extensions onto the decoration rules', () => {
    expect(segmentFormatForPath('/tmp/a.DOCX')).toBe('docx')
    expect(segmentFormatForPath('/tmp/a.pptx')).toBe('pptx')
    expect(segmentFormatForPath('/tmp/a.ppt')).toBe('pptx')
    expect(segmentFormatForPath('/tmp/a.xlsm')).toBe('xlsx')
    expect(segmentFormatForPath('/tmp/a.pdf')).toBe('flat')
    expect(segmentFormatForPath('/tmp/a.txt')).toBe('flat')
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

  it('explains how to fix a PDF with no text layer', async () => {
    // Regression: this used to surface the generic 'no extractable text'
    // message, which gives the user nothing to act on. Scans and PDFs whose
    // CJK fonts carry no ToUnicode map land here.
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-pdf-'))
    const inputPath = join(__dirname, 'fixtures', 'pdf-without-text-layer.pdf')
    const result = await buildDictionary(
      { inputPath, sourceLang: 'zh-CN', targetLang: 'en-US', dataDir: dir },
      { translateBatch: async () => ({ ok: true, units: [] }) },
    )
    expect(result.ok).toBe(false)
    expect(result.error).toContain('no text layer')
    expect(result.error).toContain('OCR')
  })

  it('reports a missing input file', async () => {
    const result = await buildDictionary(
      { inputPath: '/nope/missing.txt', sourceLang: 'en-US', targetLang: 'zh-CN', dataDir: '/tmp' },
      { translateBatch: async () => ({ ok: true, units: [] }) },
    )
    expect(result.ok).toBe(false)
  })
})

describe('coverage + gap filling', () => {
  /** Same shape as the outer suite's helper; `writeInput` is describe-scoped. */
  function writeInput(contents: string): string {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-input-'))
    const file = join(dir, 'input.txt')
    writeFileSync(file, contents, 'utf8')
    return file
  }

  /** A stub that upper-cases, so we can tell translated from untouched. */
  const upperBatch = () => async (input: { units: Array<{ sourceText: string }> }) => ({
    ok: true,
    units: input.units.map((u, i) => ({
      unitId: `u${i}`,
      sourceText: u.sourceText,
      translatedText: u.sourceText.toUpperCase(),
      status: 'translated' as const,
    })),
  })

  it('reports coverage on a KB-only build', async () => {
    // With useLlm:false nothing from the file is translated, and the result has
    // to say so — this is the case the UI previously showed as a plain success.
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-cov-'))
    const inputPath = writeInput('Alpha\nBeta\nGamma\n')
    const result = await buildDictionary(
      {
        inputPath,
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        outputPath: join(dir, 'dict.json'),
        dataDir: dir,
        useLlm: false,
      },
      { translateBatch: async () => ({ ok: true, units: [] }) },
    )
    expect(result.ok).toBe(true)
    expect(result.coverage?.total).toBe(3)
    expect(result.coverage?.covered).toBe(0)
    expect(result.coverage?.uncovered).toEqual(['Alpha', 'Beta', 'Gamma'])
  })

  it('counts model-translated segments as covered', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-cov2-'))
    const inputPath = writeInput('Alpha\nBeta\n')
    const result = await buildDictionary(
      {
        inputPath,
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        outputPath: join(dir, 'dict.json'),
        dataDir: dir,
      },
      { translateBatch: upperBatch() },
    )
    expect(result.coverage?.covered).toBe(2)
    expect(result.coverage?.ratio).toBe(1)
  })

  it('fills only the uncovered segments and keeps hand edits', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-fill-'))
    const inputPath = writeInput('Alpha\nBeta\nGamma\n')
    const partial = join(dir, 'partial.json')
    writeFileSync(partial, JSON.stringify({ Alpha: '阿尔法(手改)' }), 'utf8')

    let sent: string[] = []
    const result = await fillDictionaryGaps(
      {
        inputPath,
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        dictionaryPath: partial,
        outputPath: join(dir, 'complete.json'),
        dataDir: dir,
      },
      {
        translateBatch: async (input) => {
          sent = input.units.map((u) => u.sourceText)
          return upperBatch()(input)
        },
      },
    )
    expect(result.ok).toBe(true)
    expect(sent).toEqual(['Beta', 'Gamma'])
    expect(result.added).toBe(2)
    expect(result.coverageBefore?.covered).toBe(1)
    expect(result.coverageAfter?.covered).toBe(3)
    expect(result.stillUncovered).toEqual([])
    const written = readDictionaryFile(join(dir, 'complete.json'))
    expect(written?.entries.Alpha).toBe('阿尔法(手改)')
    expect(written?.entries.Beta).toBe('BETA')
    // The original is left alone so the user can diff the two.
    expect(readDictionaryFile(partial)?.entries).toEqual({ Alpha: '阿尔法(手改)' })
  })

  it('re-translates partially rewritten segments as whole strings', async () => {
    // The KB matched a term inside a longer sentence, producing mixed-language
    // output. The gap-fill pass has to send the whole segment so an exact key
    // overrides the substring hit.
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-partial-'))
    const inputPath = writeInput('产品验收报告已提交。\n')
    const partial = join(dir, 'partial.json')
    writeFileSync(partial, JSON.stringify({ 产品验收报告: 'Product Acceptance Report' }), 'utf8')

    let sent: string[] = []
    const result = await fillDictionaryGaps(
      {
        inputPath,
        sourceLang: 'zh-CN',
        targetLang: 'en-US',
        dictionaryPath: partial,
        outputPath: join(dir, 'complete.json'),
        dataDir: dir,
      },
      {
        translateBatch: async (input) => {
          sent = input.units.map((u) => u.sourceText)
          return {
            ok: true,
            units: input.units.map((u, i) => ({
              unitId: `u${i}`,
              sourceText: u.sourceText,
              translatedText: 'The product acceptance report has been submitted.',
              status: 'translated' as const,
            })),
          }
        },
      },
    )
    expect(sent).toEqual(['产品验收报告已提交。'])
    expect(result.coverageBefore?.partial).toEqual(['产品验收报告已提交。'])
    expect(result.coverageAfter?.partial).toEqual([])
    expect(result.coverageAfter?.exact).toBe(1)
    expect(result.coverageAfter?.ratio).toBe(1)
  })

  it('makes no model call when the dictionary already covers the file', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-fill2-'))
    const inputPath = writeInput('Alpha\n')
    const full = join(dir, 'full.json')
    writeFileSync(full, JSON.stringify({ Alpha: '阿尔法' }), 'utf8')
    let called = false
    const result = await fillDictionaryGaps(
      {
        inputPath,
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        dictionaryPath: full,
        dataDir: dir,
      },
      {
        translateBatch: async () => {
          called = true
          return { ok: true, units: [] }
        },
      },
    )
    expect(called).toBe(false)
    expect(result.added).toBe(0)
    expect(result.coverageAfter?.ratio).toBe(1)
  })

  it('reports segments the provider failed on as still uncovered', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-dict-fill3-'))
    const inputPath = writeInput('Alpha\nBeta\n')
    const empty = join(dir, 'empty.json')
    writeFileSync(empty, '{}', 'utf8')
    const result = await fillDictionaryGaps(
      {
        inputPath,
        sourceLang: 'en-US',
        targetLang: 'zh-CN',
        dictionaryPath: empty,
        outputPath: join(dir, 'out.json'),
        dataDir: dir,
      },
      {
        translateBatch: async (input) => ({
          ok: false,
          units: input.units.map((u) => ({
            unitId: u.unitId,
            sourceText: u.sourceText,
            status: 'failed' as const,
          })),
          error: 'provider down',
        }),
      },
    )
    expect(result.ok).toBe(true)
    expect(result.added).toBe(0)
    expect(result.stillUncovered).toEqual(['Alpha', 'Beta'])
  })
})
