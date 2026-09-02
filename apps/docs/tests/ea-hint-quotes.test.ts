import { describe, expect, it } from 'vitest'
import { eaHintQuoteRanges } from '../src/renderer/editor/decoration-extensions'

const hinted = '<w:rPr><w:rFonts w:hint="eastAsia"/><w:color w:val="000000"/></w:rPr>'

describe('eaHintQuoteRanges', () => {
  it('curly quotes in a hinted run are decorated', () => {
    expect(eaHintQuoteRanges(hinted, '停止“面子游戏”真正')).toEqual([
      { from: 2, to: 3 },
      { from: 7, to: 8 },
    ])
  })

  it('hint wins regardless of surrounding script (Word probe 2026-09-01)', () => {
    expect(eaHintQuoteRanges(hinted, 'it’s a “quote”')).toEqual([
      { from: 2, to: 3 },
      { from: 7, to: 8 },
      { from: 13, to: 14 },
    ])
  })

  it('single quotes and adjacent quotes group into one range', () => {
    expect(eaHintQuoteRanges(hinted, '‘时’“”')).toEqual([
      { from: 0, to: 1 },
      { from: 2, to: 5 },
    ])
  })

  it('runs without the hint keep the Latin form', () => {
    expect(eaHintQuoteRanges('<w:rPr><w:rFonts w:ascii="Calibri"/></w:rPr>', '叫“时间”')).toEqual(
      [],
    )
    expect(eaHintQuoteRanges(null, '叫“时间”')).toEqual([])
  })

  it('other hint values do not count', () => {
    expect(eaHintQuoteRanges('<w:rPr><w:rFonts w:hint="default"/></w:rPr>', '叫“时间”')).toEqual([])
  })

  it('a hint inside the w:rPrChange revision snapshot does not count', () => {
    const raw =
      '<w:rPr><w:rFonts w:ascii="Calibri"/><w:rPrChange w:id="1" w:author="a">' +
      '<w:rPr><w:rFonts w:hint="eastAsia"/></w:rPr></w:rPrChange></w:rPr>'
    expect(eaHintQuoteRanges(raw, '叫“时间”')).toEqual([])
  })

  it('quote-free text yields nothing', () => {
    expect(eaHintQuoteRanges(hinted, '停止面子游戏')).toEqual([])
  })
})
