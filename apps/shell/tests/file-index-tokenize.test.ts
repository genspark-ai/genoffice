import { describe, expect, it } from 'vitest'
import {
  parseQuery,
  termChars,
  tokenExprs,
  toMatchExpression,
  tokenize,
} from '../src/main/file-index/tokenize'
import { buildSnippet, containsAny, excerpt } from '../src/main/file-index/snippet'
import { markText } from '../src/shared/text-marks'

describe('tokenize', () => {
  it('splits CJK runs into bigrams and keeps Latin words whole', () => {
    const t = tokenize('\u65b0\u80fd\u6e90 report 2026')
    expect(t.bi).toEqual(['\u65b0\u80fd', '\u80fd\u6e90', 'report', '2026'])
    expect(t.uni).toEqual(['\u65b0', '\u80fd', '\u6e90', 'report', '2026'])
  })

  it('emits a lone CJK character as a unigram in both streams', () => {
    expect(tokenize('\u7b2c3\u5b63\u5ea6').bi).toEqual(['\u7b2c', '3', '\u5b63\u5ea6'])
  })

  it('folds case and fullwidth forms', () => {
    expect(tokenize('ＡＢＣ Report').bi).toEqual(['abc', 'report'])
  })

  it('handles kana and hangul as CJK', () => {
    expect(tokenize('\u30c7\u30d7\u30ed\u30a4').bi).toEqual([
      '\u30c7\u30d7',
      '\u30d7\u30ed',
      '\u30ed\u30a4',
    ])
    expect(tokenize('\uc11c\uc6b8 \uc9c0\ud558\ucca0').bi).toEqual([
      '\uc11c\uc6b8',
      '\uc9c0\ud558',
      '\ud558\ucca0',
    ])
  })
})

describe('parseQuery / toMatchExpression', () => {
  it('ANDs whitespace terms as phrases with Latin prefix', () => {
    const expr = toMatchExpression(parseQuery('\u65b0\u80fd\u6e90 rep'))
    expect(expr).toBe(
      '({name path body}: "\u65b0\u80fd \u80fd\u6e90") AND ({name path body}: "rep" *)',
    )
  })

  it('routes a single CJK character to the unigram columns', () => {
    expect(toMatchExpression(parseQuery('\u544a'))).toBe('({name_u path_u body_u}: "\u544a")')
  })

  it('routes a term mixing a lone CJK character into a longer run to unigrams', () => {
    expect(toMatchExpression(parseQuery('\u7b2c1\u5b63\u5ea6'))).toBe(
      '({name_u path_u body_u}: "\u7b2c 1 \u5b63 \u5ea6")',
    )
  })

  it('keeps quoted text as one phrase without prefix and honours exclusion', () => {
    const expr = toMatchExpression(parseQuery('"annual report" -draft'))
    expect(expr).toBe('(({name path body}: "annual report")) NOT ({name path body}: "draft" *)')
  })

  it('returns null for an empty or punctuation-only query', () => {
    expect(toMatchExpression(parseQuery('  '))).toBeNull()
    expect(toMatchExpression(parseQuery('!!! ...'))).toBeNull()
  })
})

describe('tokenExprs', () => {
  it('emits one expression per distinct token with the typing prefix on the last word', () => {
    const [cjk, latin] = parseQuery('\u65b0\u80fd\u6e90 rep').include
    expect(tokenExprs(cjk!)).toEqual([
      { text: '\u65b0\u80fd', expr: '{name path body}: "\u65b0\u80fd"', at: [0, 1] },
      { text: '\u80fd\u6e90', expr: '{name path body}: "\u80fd\u6e90"', at: [1, 2] },
    ])
    expect(tokenExprs(latin!)).toEqual([
      { text: 'rep', expr: '{name path body}: "rep" *', at: [0, 1, 2] },
    ])
  })

  it('merges the offsets of a repeated token and keeps the prefix on the final word', () => {
    const [t] = parseQuery('\u6280\u672f\u4e0e\u6280\u672f-v2').include
    const map = new Map(tokenExprs(t!).map((e) => [e.text, e]))
    expect(map.get('\u6280\u672f')!.at).toEqual([0, 1, 3, 4])
    expect(map.get('v2')!.expr).toBe('{name path body}: "v2" *')
    expect(termChars(t!)).toBe(7)
  })

  it('counts code points so astral CJK characters cover exactly their offsets', () => {
    const [t] = parseQuery('\u{20000}\u{20001}\u{20002}\u{20003}').include
    const exprs = tokenExprs(t!)
    expect(exprs.map((e) => e.at)).toEqual([
      [0, 1],
      [1, 2],
      [2, 3],
    ])
    expect(termChars(t!)).toBe(4)
  })
})

describe('buildSnippet', () => {
  it('highlights compatibility characters at their original positions', () => {
    expect(markText('eﬃcient', ['effi'])).toEqual([
      { text: 'eﬃ', hit: true },
      { text: 'cient', hit: false },
    ])
    expect(markText('eﬃcient', ['ffi'])).toEqual([
      { text: 'e', hit: false },
      { text: 'ﬃ', hit: true },
      { text: 'cient', hit: false },
    ])
    expect(markText('ﬃcient', ['ffi'])[0]).toEqual({ text: 'ﬃ', hit: true })
    expect(markText('A\u030a report', ['å'])[0]).toEqual({ text: 'A\u030a', hit: true })
    expect(buildSnippet('before eﬃcient after', ['effi'])?.filter((p) => p.hit)).toEqual([
      { text: 'eﬃ', hit: true },
    ])
    expect(buildSnippet('before eﬃcient after', ['ffi'])?.filter((p) => p.hit)).toEqual([
      { text: 'ﬃ', hit: true },
    ])
    expect(excerpt('x'.repeat(100) + ' eﬃcient', ['effi'], 12)).toContain('eﬃ')
  })

  it('fuses overlapping bigram hits into one run', () => {
    const parts = buildSnippet('\u2026\u5927\u6a21\u578b\u6280\u672f\u6f14\u8fdb', [
      '\u5927\u6a21',
      '\u6a21\u578b',
      '\u578b\u6280',
      '\u6280\u672f',
    ])!
    expect(parts.filter((p) => p.hit).map((p) => p.text)).toEqual([
      '\u5927\u6a21\u578b\u6280\u672f',
    ])
  })

  it('lights Latin terms only at a word start', () => {
    const parts = buildSnippet('Minneapolis is an island', ['is'])!
    expect(parts.filter((p) => p.hit).map((p) => p.text)).toEqual(['is', 'is'])
    expect(parts.map((p) => p.text).join('')).toBe('Minneapolis is an island')
    expect(containsAny('Minneapolis', ['is'])).toBe(false)
    expect(containsAny('island', ['is'])).toBe(true)
  })

  it('windows around the first hit and flags every term inside', () => {
    const text = 'a'.repeat(300) + ' \u8c03\u7814\u62a5\u544a 2026 report ' + 'b'.repeat(300)
    const parts = buildSnippet(text, ['\u62a5\u544a', 'rep'])!
    expect(parts[0]).toEqual({ text: '…', hit: false })
    expect(parts.filter((p) => p.hit).map((p) => p.text)).toEqual(['\u62a5\u544a', 'rep'])
    expect(parts[parts.length - 1]).toEqual({ text: '…', hit: false })
    expect(parts.map((p) => p.text).join('').length).toBeLessThanOrEqual(160)
  })

  it('matches case-insensitively and returns null without a hit', () => {
    expect(buildSnippet('Quarterly REPORT', ['report'])![1]).toEqual({ text: 'REPORT', hit: true })
    expect(buildSnippet('nothing here', ['report'])).toBeNull()
  })
})
