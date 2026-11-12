/**
 * Word breaks no line at a space beside U+00B7 or before ":" when the East
 * Asian language is zh-CN or absent (probe 2026-11-08).
 */
import { describe, expect, it } from 'vitest'
import { inlineRunsOf, inlineSegments } from '../src/renderer/text-layout/inline-segments'

const DOT = '·'
type TestRun = { text: string; eastAsiaLang?: string; image?: unknown }

const noBreaks = (runs: TestRun[], eaLang: string | null = null) => {
  const text = runs.map((r) => r.text + (r.image ? '￼' : '')).join('')
  return inlineSegments({
    runs: inlineRunsOf(runs).runs,
    autospace: null,
    inheritedAscii: null,
    inheritedEa: null,
    wordWrap: true,
    balancedSpaces: false,
    justified: false,
    simsunGapLineHeight: null,
    eaLang,
  })
    .segments.filter((s) => s.cls === 'doc-nobreak')
    .map((s) => text.slice(s.from, s.to))
}

describe('spaces beside a middle dot', () => {
  it('glues both sides without an East Asian language or under zh-CN', () => {
    expect(noBreaks([{ text: `mmmm ${DOT} nnnn` }])).toEqual([`m ${DOT} `])
    expect(noBreaks([{ text: `mmmm ${DOT} nnnn` }], 'zh-CN')).toEqual([`m ${DOT} `])
    expect(noBreaks([{ text: `mmmm  ${DOT}  nnnn` }])).toEqual([`m  ${DOT}  `])
    expect(noBreaks([{ text: `mmmm${DOT} nnnn` }])).toEqual([`${DOT} `])
    expect(noBreaks([{ text: `mmmm ${DOT}nnnn` }])).toEqual(['m '])
  })

  it('glues only the side before the dot under zh-TW and neither side in other languages', () => {
    expect(noBreaks([{ text: `mmmm ${DOT} nnnn` }], 'zh-TW')).toEqual(['m '])
    expect(noBreaks([{ text: `mmmm ${DOT} nnnn` }], 'pt-BR')).toEqual([])
    expect(noBreaks([{ text: `mmmm ${DOT} nnnn`, eastAsiaLang: 'ja-JP' }])).toEqual([])
  })

  it('glues the space before a colon, never the one after', () => {
    expect(noBreaks([{ text: 'mmmm : nnnn' }])).toEqual(['m '])
    expect(noBreaks([{ text: 'mmmm : nnnn' }], 'en-US')).toEqual([])
  })

  it('leaves bullets, dashes and trailing spaces alone', () => {
    expect(noBreaks([{ text: 'mmmm • nnnn — oooo' }])).toEqual([])
    expect(noBreaks([{ text: `mmmm ${DOT} ` }])).toEqual(['m '])
  })

  it('keeps the span inside one run and never opens it on a space', () => {
    expect(noBreaks([{ text: `mmmm ${DOT} ` }, { text: 'nnnn' }])).toEqual([`m ${DOT} `])
    expect(noBreaks([{ text: 'mmmm' }, { text: ` ${DOT} nnnn` }])).toEqual([`${DOT} `])
    expect(noBreaks([{ text: 'mmmm ' }, { text: `${DOT} nnnn` }])).toEqual(['m ', `${DOT} `])
    expect(noBreaks([{ text: `mmmm ${DOT} `, image: {} }, { text: 'nnnn' }])).toEqual(['m '])
  })

  it('hands the wrap mode to a balanced-space span inside the glued range', () => {
    const text = `\u4e2d\u6587 ${DOT} nnnn`
    const segs = inlineSegments({
      runs: inlineRunsOf([{ text }]).runs,
      autospace: null,
      inheritedAscii: null,
      inheritedEa: null,
      wordWrap: true,
      balancedSpaces: true,
      justified: false,
      simsunGapLineHeight: null,
      eaLang: null,
    }).segments
    const glued = segs
      .filter((s) => s.cls?.split(' ').includes('doc-nobreak'))
      .sort((a, b) => a.from - b.from)
    expect(glued.map((s) => text.slice(s.from, s.to)).join('')).toBe(`\u6587 ${DOT} `)
    expect(glued.some((s) => s.cls?.includes('doc-balanced-space'))).toBe(true)
    for (let i = 1; i < glued.length; i++) expect(glued[i].from).toBe(glued[i - 1].to)
  })
})
