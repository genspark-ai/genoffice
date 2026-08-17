import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
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
  it('does not apply to plain autoshapes (only txBox text boxes)', () => {
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="6" name="Shape 5"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"/></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-GB"/><a:t>shape text</a:t></a:r></a:p></p:txBody></p:sp>'
    const ctx = { defaultTextStyle: parseDefaultTextStyle(PRES) }
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx })
    expect((slide.elements[0] as any).text.paragraphs[0].runs[0].fontSize).not.toBe(12)
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
