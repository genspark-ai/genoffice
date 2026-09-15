import { describe, expect, it } from 'vitest'

import {
  buildTranslateSystemPrompt,
  buildTranslationPrompt,
  extractTranslationText,
  normalizeSourceLang,
} from '../src/prompt'

describe('buildTranslateSystemPrompt', () => {
  it('mentions source / target languages and the preserve-format rule', () => {
    const sys = buildTranslateSystemPrompt({
      sourceLang: 'auto',
      targetLang: 'zh-CN',
      preserveFormat: true,
    })
    expect(sys).toMatch(/Source language: Auto-detect\./)
    expect(sys).toMatch(/Target language: Simplified Chinese\./)
    expect(sys).toMatch(/Preserve the original formatting/)
  })

  it('switches off the preserve rule when requested', () => {
    const sys = buildTranslateSystemPrompt({
      sourceLang: 'en-US',
      targetLang: 'ja-JP',
      preserveFormat: false,
    })
    expect(sys).toMatch(/Return only the translated text/)
    expect(sys).not.toMatch(/Preserve the original formatting/)
  })
})

describe('buildTranslationPrompt', () => {
  it('wraps the source in <source_text> delimiters', () => {
    const p = buildTranslationPrompt('Hello')
    expect(p).toContain('<source_text>\nHello\n</source_text>')
    expect(p).toContain('Translate the literal text')
  })
})

describe('extractTranslationText', () => {
  it('strips <think> blocks', () => {
    expect(extractTranslationText('<think>chain of thought</think>Hello')).toBe('Hello')
    expect(extractTranslationText('<think>still going\nmulti-line</think>\nDone')).toBe('Done')
  })

  it('returns null for empty after stripping', () => {
    expect(extractTranslationText('<think>only thoughts</think>')).toBeNull()
    expect(extractTranslationText('   \n  ')).toBeNull()
  })

  it('returns the trimmed text when there is no think block', () => {
    expect(extractTranslationText('  Hello world  ')).toBe('Hello world')
  })
})

describe('normalizeSourceLang', () => {
  it('maps undefined / "auto" to "auto"', () => {
    expect(normalizeSourceLang(undefined)).toBe('auto')
    expect(normalizeSourceLang('auto')).toBe('auto')
  })
  it('passes through a real code', () => {
    expect(normalizeSourceLang('en-US')).toBe('en-US')
  })
})
