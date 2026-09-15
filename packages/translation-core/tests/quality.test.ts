import { describe, expect, it } from 'vitest'

import { assessBatchQuality, assessQuality } from '../src/quality'

describe('assessQuality', () => {
  it('flags empty output', () => {
    const r = assessQuality('hello world', '')
    expect(r.empty).toBe(true)
    expect(r.overallScore).toBe(0)
    expect(r.warnings).toContain('empty')
  })

  it('flags a refusal marker', () => {
    const r = assessQuality('translate me', 'I cannot translate that for you.')
    expect(r.warnings).toContain('refusal-marker')
  })

  it('flags an untranslated echo of a meaningful input', () => {
    const r = assessQuality(
      'Hello world, this is a longer test case.',
      'Hello world, this is a longer test case.',
    )
    expect(r.warnings).toContain('untranslated')
    expect(r.untranslated).toBe(true)
  })

  it('returns a clean report for plausible translations', () => {
    const r = assessQuality('Hello world', '你好世界')
    expect(r.warnings).toEqual([])
    expect(r.overallScore).toBe(1)
  })

  it('flags too-short / too-long relative to source length', () => {
    const r1 = assessQuality('a'.repeat(40), 'x')
    expect(r1.warnings).toContain('too-short')
    const r2 = assessQuality('hi', 'a'.repeat(40))
    expect(r2.warnings).toContain('too-long')
  })
})

describe('assessBatchQuality', () => {
  it('averages the unit-level scores and dedupes warnings', () => {
    const r = assessBatchQuality([
      { sourceText: 'hello', translatedText: '你好' },
      { sourceText: 'world', translatedText: '' },
    ])
    expect(r.warnings).toContain('empty')
    expect(r.overallScore).toBeCloseTo(0.5, 1)
  })

  it('returns a clean report for an empty batch', () => {
    const r = assessBatchQuality([])
    expect(r.overallScore).toBe(1)
    expect(r.warnings).toEqual([])
  })
})
