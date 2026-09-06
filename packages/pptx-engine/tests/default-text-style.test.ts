import { describe, it, expect } from 'vitest'
import { parseDecorations, parseSlide } from '../src/parse'
import {
  parseDefaultTextStyle,
  parsePlaceholderMap,
  resolvePlaceholderFillSpPr,
} from '../src/placeholder'

const slideWith = (bodyShapes: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${bodyShapes}</p:spTree></p:cSld></p:sld>`

const PRES =
  '<?xml version="1.0"?><p:presentation xmlns:p="p" xmlns:a="a">' +
  '<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' +
  '<a:lvl1pPr algn="ctr"><a:defRPr sz="1200"><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr>' +
  '</p:defaultTextStyle></p:presentation>'

describe('presentation defaultTextStyle (napierone 0042)', () => {
  it('parses lvl1 size and font', () => {
    const st = parseDefaultTextStyle(PRES)!
    expect(st.levels[0]?.fontSize).toBe(12)
    expect(st.levels[0]?.latinFont).toBe('Arial')
  })

  it('applies to a non-placeholder TextBox run without explicit size', () => {
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-GB"/><a:t>plain</a:t></a:r></a:p></p:txBody></p:sp>'
    const ctx = { defaultTextStyle: parseDefaultTextStyle(PRES) }
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx })
    const el = slide.elements[0] as any
    expect(el.text.paragraphs[0].runs[0].fontSize).toBe(12)
  })

  it('does not override an explicit run size', () => {
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="5" name="TextBox 4"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-GB" sz="2400"/><a:t>big</a:t></a:r></a:p></p:txBody></p:sp>'
    const ctx = { defaultTextStyle: parseDefaultTextStyle(PRES) }
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx })
    expect((slide.elements[0] as any).text.paragraphs[0].runs[0].fontSize).toBe(24)
  })
})

describe('presentation defaultTextStyle scope', () => {
  // PowerPoint probe (defaultTextStyle 14pt vs master otherStyle 28pt): rect and text box,
  // on the slide, layout and master, all render at 14pt — otherStyle is never consulted
  const RECT =
    '<p:sp><p:nvSpPr><p:cNvPr id="6" name="Shape 5"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>' +
    '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-GB"/><a:t>shape text</a:t></a:r></a:p></p:txBody></p:sp>'

  it('applies to plain autoshapes as well as text boxes', () => {
    const ctx = { defaultTextStyle: parseDefaultTextStyle(PRES) }
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(RECT), ctx })
    expect((slide.elements[0] as any).text.paragraphs[0].runs[0].fontSize).toBe(12)
  })

  it('applies to layout/master decoration shapes', () => {
    const layout =
      '<?xml version="1.0"?><p:sldLayout xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
      `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${RECT}</p:spTree></p:cSld></p:sldLayout>`
    const ctx = { defaultTextStyle: parseDefaultTextStyle(PRES) }
    const els = parseDecorations(layout, ctx)
    expect((els[0] as any).text.paragraphs[0].runs[0].fontSize).toBe(12)
  })
})

describe('placeholder fill donor matching (fed deck body ph)', () => {
  const LAYOUT =
    '<?xml version="1.0"?><p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>' +
    // the true match: body idx 10, has geometry but NO fill
    '<p:sp><p:nvSpPr><p:cNvPr id="18" name="Text Placeholder 12"/><p:nvPr><p:ph type="body" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="1106182" y="3257551"/><a:ext cx="5447018" cy="685799"/></a:xfrm></p:spPr></p:sp>' +
    // a decorative content placeholder (no type -> normalized body) WITH a dark fill
    '<p:sp><p:nvSpPr><p:cNvPr id="20" name="Content Placeholder 20"/><p:nvPr><p:ph idx="13"/></p:nvPr></p:nvSpPr>' +
    '<p:spPr><a:xfrm><a:off x="838199" y="637032"/><a:ext cx="758952" cy="201168"/></a:xfrm>' +
    '<a:solidFill><a:srgbClr val="3C718F"/></a:solidFill></p:spPr></p:sp>' +
    '</p:spTree></p:cSld></p:sldLayout>'

  it('an exact placeholder match without a fill does not steal a sibling fill', () => {
    const map = parsePlaceholderMap(LAYOUT)
    expect(resolvePlaceholderFillSpPr(map, undefined, 'body', '10')).toBeUndefined()
  })

  it('the matched placeholder with a fill still donates it', () => {
    const map = parsePlaceholderMap(LAYOUT)
    const hit = resolvePlaceholderFillSpPr(map, undefined, 'body', '13')
    expect(hit).toBeTruthy()
  })
})

describe('mixed-script runs keep the a:latin face for Latin characters', () => {
  const run = (text: string, lang = 'ko-KR') => {
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="7" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
      `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="${lang}"><a:latin typeface="NanumSquareEB"/>` +
      `<a:ea typeface="Malgun Gothic"/></a:rPr><a:t>${text}</a:t></a:r></a:p></p:txBody></p:sp>`
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx: {} })
    return (slide.elements[0] as any).text.paragraphs[0].runs[0]
  }
  it('a Hangul run with Latin digits picks the ea face and remembers the latin face', () => {
    const r = run('ISO 45001 안전')
    expect(r.fontFamily).toBe('Malgun Gothic')
    expect(r.latinFamily).toBe('NanumSquareEB')
    expect(r.fontScriptHint).toBe('ko')
  })
  it('a complex-script run picks a:cs and records no latin alternate', () => {
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="8" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="ar-SA"><a:latin typeface="Calibri"/>' +
      '<a:cs typeface="Arial"/></a:rPr><a:t>مرحبا 2024</a:t></a:r></a:p></p:txBody></p:sp>'
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx: {} })
    const r = (slide.elements[0] as any).text.paragraphs[0].runs[0]
    expect(r.fontFamily).toBe('Arial')
    expect(r.latinFamily).toBeUndefined()
  })
  it('a pure-Latin run draws with the latin face and carries no alternate', () => {
    const r = run('Global Trend')
    expect(r.fontFamily).toBe('NanumSquareEB')
    expect(r.latinFamily).toBeUndefined()
  })
})

describe('Latin-only runs: declared charset steers the substitute, lang alone does not', () => {
  const run = (latinAttrs: string, lang = 'en-US', altLang = 'ko-KR') => {
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="9" name="T"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
      `<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="${lang}" altLang="${altLang}"><a:latin ${latinAttrs}/>` +
      '</a:rPr><a:t>ISO 45001</a:t></a:r></a:p></p:txBody></p:sp>'
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx: {} })
    return (slide.elements[0] as any).text.paragraphs[0].runs[0]
  }
  it('charset=-127 (Hangul) on the latin face hints ko even for Latin text (prod_029)', () => {
    const r = run('typeface="LG스마트체 Regular" charset="-127"')
    expect(r.fontFamily).toBe('LG스마트체 Regular')
    expect(r.fontScriptHint).toBe('ko')
  })
  it('no charset: Latin text carries no hint despite altLang ko-KR (prod_026)', () => {
    expect(run('typeface="NanumSquareExtraBold"').fontScriptHint).toBeUndefined()
  })
})
