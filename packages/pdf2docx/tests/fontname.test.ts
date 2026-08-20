/** PS-name → family-name heuristic unit tests (P5): pure string mapping, no wasm. */
import { describe, expect, it } from 'vitest'
import {
  familyFromPsName,
  italicFromPsName,
  stripTrailingStyleWords,
  weightFromPsName,
} from '../src/extract/fontname'

describe('familyFromPsName', () => {
  it('splits camelCase and drops the style suffix', () => {
    expect(familyFromPsName('NotoSansSC-Regular')).toBe('Noto Sans SC')
    expect(familyFromPsName('NotoSansSC-Bold')).toBe('Noto Sans SC')
    expect(familyFromPsName('SourceHanSerifCN-Medium')).toBe('Source Han Serif CN')
  })

  it('strips the embedded-subset prefix', () => {
    expect(familyFromPsName('ABCDEF+NotoSansSC-Regular')).toBe('Noto Sans SC')
  })

  it('handles legacy foundry tags (MT/PS)', () => {
    expect(familyFromPsName('ArialMT')).toBe('Arial')
    expect(familyFromPsName('Arial-BoldMT')).toBe('Arial')
    expect(familyFromPsName('TimesNewRomanPSMT')).toBe('Times New Roman')
    expect(familyFromPsName('TimesNewRomanPS-BoldItalicMT')).toBe('Times New Roman')
  })

  it('handles comma style suffixes and multi-word styles', () => {
    expect(familyFromPsName('Helvetica,Bold')).toBe('Helvetica')
    expect(familyFromPsName('Times-Roman')).toBe('Times')
    expect(familyFromPsName('Georgia-BoldItalic')).toBe('Georgia')
  })

  it('keeps suffix segments that are not pure style words', () => {
    expect(familyFromPsName('Neue-Haas')).toBe('Neue Haas')
  })

  it('keeps camelCase brand words intact (PingFang)', () => {
    expect(familyFromPsName('PingFangSC-Regular')).toBe('PingFang SC')
    expect(familyFromPsName('PingFangSC-Semibold')).toBe('PingFang SC')
  })

  it('passes names that already contain spaces through untouched', () => {
    expect(familyFromPsName('Noto Sans SC')).toBe('Noto Sans SC')
    expect(familyFromPsName('PingFang SC')).toBe('PingFang SC')
  })

  it('returns empty for empty input', () => {
    expect(familyFromPsName('')).toBe('')
  })
})

describe('weightFromPsName (P21 A)', () => {
  it('trusts an explicit Roman/Regular suffix over any descriptor weight', () => {
    expect(weightFromPsName('AEXPIG+Times-Roman')).toBe('regular')
    expect(weightFromPsName('HelveticaNeueLTStd-Roman')).toBe('regular')
    expect(weightFromPsName('HelveticaWorld-Regular')).toBe('regular')
    expect(weightFromPsName('Frutiger-Light')).toBe('regular')
  })

  it('recognizes explicit bold-ish suffixes including foundry abbreviations', () => {
    expect(weightFromPsName('FLRKVV+Times-Bold')).toBe('bold')
    expect(weightFromPsName('Helvetica,Bold')).toBe('bold')
    expect(weightFromPsName('HelveticaNeueLTStd-Bd')).toBe('bold')
    expect(weightFromPsName('HelveticaLTStd-Blk')).toBe('bold')
    expect(weightFromPsName('ITCFranklinGothicStd-Demi')).toBe('bold')
    expect(weightFromPsName('Arial-BoldItalicMT')).toBe('bold')
  })

  it('counts a style word glued onto the base for bold only', () => {
    expect(weightFromPsName('ArialBold')).toBe('bold')
    // 'Roman' inside an unsuffixed base is family vocabulary, not a weight claim
    expect(weightFromPsName('TimesNewRomanPSMT')).toBe(null)
    expect(weightFromPsName('Times New Roman')).toBe(null)
  })

  it('declares nothing for style-free names', () => {
    expect(weightFromPsName('AdvP4C4E74')).toBe(null)
    expect(weightFromPsName('Blackadder ITC')).toBe(null)
    expect(weightFromPsName('Helvetica')).toBe(null)
  })
})

describe('italicFromPsName (P21 A)', () => {
  it('recognizes italic suffixes including abbreviations', () => {
    expect(italicFromPsName('HelveticaNeueLTStd-It')).toBe(true)
    expect(italicFromPsName('Arial-BoldItalicMT')).toBe(true)
    expect(italicFromPsName('Helvetica-Oblique')).toBe(true)
    expect(italicFromPsName('HelveticaWorld-Regular')).toBe(false)
    expect(italicFromPsName('Palatino-Roman')).toBe(false)
  })
})

describe('stripTrailingStyleWords (P21 A)', () => {
  it('drops trailing pure-style words from spaced family names', () => {
    expect(stripTrailingStyleWords('Helvetica Neue LTStd It')).toBe('Helvetica Neue LTStd')
    expect(stripTrailingStyleWords('Helvetica Neue LTStd Bd')).toBe('Helvetica Neue LTStd')
    expect(stripTrailingStyleWords('ITCFranklin Gothic Std Demi')).toBe('ITCFranklin Gothic Std')
    expect(stripTrailingStyleWords('Arial Bold')).toBe('Arial')
  })

  it('keeps names whose last word is not a style word, and never empties a name', () => {
    expect(stripTrailingStyleWords('Helvetica World')).toBe('Helvetica World')
    expect(stripTrailingStyleWords('Bold')).toBe('Bold')
    expect(stripTrailingStyleWords('Noto Sans SC')).toBe('Noto Sans SC')
  })
})
