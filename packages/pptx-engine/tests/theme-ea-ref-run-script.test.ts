import { describe, expect, it } from 'vitest'
import { parseSlide } from '../src/parse'
import { parseTheme } from '../src/theme'
import type { TextElement } from '../src/types'

// Theme with an empty <a:ea/> and a Jpan script font (Japanese government/university decks)
const THEME = parseTheme(
  '<?xml version="1.0"?><a:theme xmlns:a="a"><a:themeElements>' +
    '<a:clrScheme name="x"><a:dk1><a:srgbClr val="000000"/></a:dk1></a:clrScheme>' +
    '<a:fontScheme name="x"><a:majorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/>' +
    '<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/></a:majorFont>' +
    '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>' +
    '</a:fontScheme><a:fmtScheme name="x"/></a:themeElements></a:theme>',
)

const parseRuns = (runs: string) => {
  const slide = parseSlide({
    path: 'ppt/slides/slide1.xml',
    slideXml:
      '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>' +
      '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/>' +
      // the level default carries the theme ref, as a master titleStyle does
      '<a:lstStyle><a:lvl1pPr><a:defRPr><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></a:lstStyle>' +
      `<a:p>${runs}</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`,
    ctx: { theme: THEME },
  })
  return (slide.elements[0] as TextElement).text!.paragraphs[0]!.runs
}

describe('inherited +mj-ea re-resolves with the run script', () => {
  it('a ja-JP run without its own <a:ea> gets the Jpan script font, not the Latin fallback', () => {
    const [r] = parseRuns(
      '<a:r><a:rPr lang="ja-JP" altLang="en-US"/><a:t>プログラミング</a:t></a:r>',
    )
    expect(r!.fontFamily).toBe('ＭＳ Ｐゴシック')
  })

  it('a Latin run keeps the Latin resolution of the same ref', () => {
    const [r] = parseRuns('<a:r><a:rPr lang="en-US"/><a:t>Hello</a:t></a:r>')
    expect(r!.fontFamily).toBe('Calibri')
  })

  it('a paragraph defRPr with a concrete <a:ea> replaces the inherited theme ref', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml:
        '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>' +
        '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/>' +
        '<a:lstStyle><a:lvl1pPr><a:defRPr><a:ea typeface="+mj-ea"/></a:defRPr></a:lvl1pPr></a:lstStyle>' +
        '<a:p><a:pPr><a:defRPr><a:ea typeface="Meiryo"/></a:defRPr></a:pPr>' +
        '<a:r><a:rPr lang="ja-JP"/><a:t>\u30d7\u30ed</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
      ctx: { theme: THEME },
    })
    const [r] = (slide.elements[0] as TextElement).text!.paragraphs[0]!.runs
    expect(r!.fontFamily).toBe('Meiryo')
  })

  it('an explicit run <a:ea> still wins over the inherited ref', () => {
    const [r] = parseRuns(
      '<a:r><a:rPr lang="ja-JP"><a:ea typeface="Meiryo"/></a:rPr><a:t>プロ</a:t></a:r>',
    )
    expect(r!.fontFamily).toBe('Meiryo')
  })
})
