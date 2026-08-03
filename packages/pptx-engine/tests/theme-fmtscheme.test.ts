/**
 * Theme fmtScheme template inheritance: p:style's fillRef/lnRef/effectRef take the
 * fillStyleLst/lnStyleLst/effectStyleLst template by idx and substitute phClr.
 */
import { describe, it, expect } from 'vitest'
import { parseTheme } from '../src/theme'
import { parseSlide } from '../src/parse'
import { applyColorMods } from '../src/color'
import type { TextElement } from '../src/types'

// Trimmed-down Office theme: fillStyleLst = [solid, linear gradient, radial gradient],
// lnStyleLst has three line widths (6350/12700/19050, second entry dashed), effectStyleLst third entry has shadow + glow
const THEME_XML =
  '<?xml version="1.0"?><a:theme xmlns:a="a" name="T"><a:themeElements>' +
  '<a:clrScheme name="c"><a:dk1><a:srgbClr val="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>' +
  '<a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>' +
  '<a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>' +
  '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4>' +
  '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6>' +
  '<a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme>' +
  '<a:fontScheme name="f"><a:majorFont><a:latin typeface="Calibri Light"/></a:majorFont>' +
  '<a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>' +
  '<a:fmtScheme name="Office">' +
  '<a:fillStyleLst>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:gradFill rotWithShape="1"><a:gsLst>' +
  '<a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs>' +
  '<a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs>' +
  '</a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
  '<a:gradFill rotWithShape="1"><a:gsLst>' +
  '<a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/></a:schemeClr></a:gs>' +
  '<a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="78000"/></a:schemeClr></a:gs>' +
  '</a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill>' +
  '</a:fillStyleLst>' +
  '<a:lnStyleLst>' +
  '<a:ln w="6350" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
  '<a:ln w="12700" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="dash"/></a:ln>' +
  '<a:ln w="19050" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>' +
  '</a:lnStyleLst>' +
  '<a:effectStyleLst>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle>' +
  '<a:effectStyle><a:effectLst/></a:effectStyle>' +
  '<a:effectStyle><a:effectLst>' +
  '<a:glow rad="63500"><a:schemeClr val="phClr"><a:alpha val="40000"/></a:schemeClr></a:glow>' +
  '<a:outerShdw blurRad="57150" dist="19050" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw>' +
  '</a:effectLst></a:effectStyle>' +
  '</a:effectStyleLst>' +
  '<a:bgFillStyleLst>' +
  '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
  '<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill>' +
  '<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="60000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="0" scaled="1"/></a:gradFill>' +
  '</a:bgFillStyleLst>' +
  '</a:fmtScheme></a:themeElements></a:theme>'

const theme = parseTheme(THEME_XML)

const slideWith = (body: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`

const shapeWithStyle = (style: string, spPrExtra = '') =>
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm>' +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${spPrExtra}</p:spPr>` +
  `<p:style>${style}</p:style>` +
  '<p:txBody><a:bodyPr/><a:p><a:r><a:t>x</a:t></a:r></a:p></p:txBody></p:sp>'

const STYLE_REFS =
  '<a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef>' +
  '<a:fillRef idx="2"><a:schemeClr val="accent2"/></a:fillRef>' +
  '<a:effectRef idx="3"><a:schemeClr val="accent1"/></a:effectRef>' +
  '<a:fontRef idx="minor"><a:schemeClr val="dk1"/></a:fontRef>'

describe('parseTheme fmtScheme', () => {
  it('fillStyleLst/lnStyleLst/effectStyleLst/bgFillStyleLst parse in order', () => {
    expect(theme.fillStyles).toHaveLength(3)
    expect(theme.fillStyles![0]['a:solidFill']).toBeTruthy()
    expect(theme.fillStyles![1]['a:gradFill']).toBeTruthy()
    expect(theme.lnStyles).toHaveLength(3)
    expect((theme.lnStyles![1]!['a:ln'] as Record<string, unknown>)['@_w']).toBe('12700')
    expect(theme.effectStyles).toHaveLength(3)
    expect(theme.bgFillStyles).toHaveLength(3)
  })
})

describe('p:style references -> fmtScheme templates', () => {
  const parse = () =>
    parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(shapeWithStyle(STYLE_REFS)),
      ctx: { theme },
    })

  it('fillRef idx=2 -> linear gradient template, phClr substituted with accent2 and modifiers applied', () => {
    const el = parse().elements[0] as TextElement
    expect(el.fill?.type).toBe('gradient')
    const grad = el.fill as Extract<TextElement['fill'], { type: 'gradient' }>
    expect(grad.stops).toHaveLength(2)
    expect(grad.angle).toBe(5400000)
    // First stop = accent2 + lumMod 110% + satMod 105% + tint 67%
    const expected = applyColorMods('#ED7D31', {
      'a:lumMod': { '@_val': '110000' },
      'a:satMod': { '@_val': '105000' },
      'a:tint': { '@_val': '67000' },
    })
    expect(grad.stops[0]!.color).toBe(expected)
    expect(grad.stops[0]!.color).not.toBe('#ED7D31')
  })

  it('lnRef idx=2 -> template line width 12700 + dash (no longer hardcoded 1pt solid)', () => {
    const el = parse().elements[0] as TextElement
    expect(el.stroke?.width).toBe(12700)
    expect(el.stroke?.dash).toBe('dash')
    expect(el.stroke?.fill).toEqual({ type: 'solid', color: '#4472C4' })
  })

  it('effectRef idx=3 -> shadow + glow (phClr substituted with the reference color)', () => {
    const el = parse().elements[0] as TextElement
    expect(el.shadow).toBeTruthy()
    expect(el.shadow!.blurRad).toBe(57150)
    expect(el.shadow!.dist).toBe(19050)
    expect(el.shadow!.color.startsWith('#000000')).toBe(true)
    expect(el.glow).toBeTruthy()
    expect(el.glow!.radius).toBe(63500)
    expect(el.glow!.color.startsWith('#4472C4')).toBe(true)
  })

  it('fillRef idx=1001 -> first entry of bgFillStyleLst', () => {
    const style =
      '<a:lnRef idx="0"/><a:fillRef idx="1001"><a:schemeClr val="accent3"/></a:fillRef>' +
      '<a:effectRef idx="0"/><a:fontRef idx="minor"/>'
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(shapeWithStyle(style)),
      ctx: { theme },
    })
    expect((slide.elements[0] as TextElement).fill).toEqual({ type: 'solid', color: '#A5A5A5' })
  })

  it('explicit spPr fill/stroke wins over style references (not overridden by templates)', () => {
    const explicit =
      '<a:solidFill><a:srgbClr val="112233"/></a:solidFill><a:ln w="25400"><a:solidFill><a:srgbClr val="445566"/></a:solidFill></a:ln>'
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(shapeWithStyle(STYLE_REFS, explicit)),
      ctx: { theme },
    })
    const el = slide.elements[0] as TextElement
    expect(el.fill).toEqual({ type: 'solid', color: '#112233' })
    expect(el.stroke?.width).toBe(25400)
  })

  it('explicit <a:ln><a:noFill/> wins over lnRef: no stroke is drawn', () => {
    const noOutline = '<a:ln><a:noFill/></a:ln>'
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(shapeWithStyle(STYLE_REFS, noOutline)),
      ctx: { theme },
    })
    expect((slide.elements[0] as TextElement).stroke).toBeUndefined()
  })

  it('connector with explicit <a:ln><a:noFill/> draws no stroke; absent a:ln still falls back', () => {
    const cxn = (ln: string) =>
      '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="7" name="C"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000" cy="1000"/></a:xfrm>' +
      `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>${ln}</p:spPr></p:cxnSp>`
    const parseCxn = (ln: string) =>
      parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(cxn(ln)), ctx: { theme } })
        .elements[0] as TextElement
    expect(parseCxn('<a:ln><a:noFill/></a:ln>').stroke).toBeUndefined()
    expect(parseCxn('').stroke).toBeTruthy()
  })

  it('theme without fmtScheme (legacy behavior): falls back to reference-color solid + 1pt', () => {
    const bare: any = { colors: { accent1: '#4472C4', accent2: '#ED7D31', dk1: '#000000' } }
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(shapeWithStyle(STYLE_REFS)),
      ctx: { theme: bare },
    })
    const el = slide.elements[0] as TextElement
    expect(el.fill?.type).toBe('solid')
    expect(el.stroke?.width).toBe(12700)
  })
})
