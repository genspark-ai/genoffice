import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'

const slideWith = (bodyShapes: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${bodyShapes}</p:spTree></p:cSld></p:sld>`

const GRP_XFRM =
  '<a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="1124744"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="9144000" cy="1124744"/></a:xfrm>'

const child = (fillXml: string, style = '') =>
  '<p:sp><p:nvSpPr><p:cNvPr id="8" name="R"/></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
  `<a:prstGeom prst="rect"/>${fillXml}</p:spPr>${style}</p:sp>`

const STYLE =
  '<p:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
  '<a:fontRef idx="minor"/></p:style>'

describe('<a:grpFill/> (PlanS master banner)', () => {
  it('child with grpFill inherits the group grpSpPr fill, beating the style fillRef', () => {
    const grp =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="7" name="G"/></p:nvGrpSpPr>' +
      `<p:grpSpPr>${GRP_XFRM}<a:solidFill><a:srgbClr val="C5DCF0"/></a:solidFill></p:grpSpPr>` +
      child('<a:grpFill/>', STYLE) +
      '</p:grpSp>'
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(grp), ctx: {} })
    const g = slide.elements[0] as any
    expect(g.type).toBe('group')
    expect(g.children[0].fill).toEqual({ type: 'solid', color: '#C5DCF0' })
  })

  it('nested group with grpFill defers to the outer group fill', () => {
    const inner =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="9" name="G2"/></p:nvGrpSpPr>' +
      `<p:grpSpPr>${GRP_XFRM}<a:grpFill/></p:grpSpPr>` +
      child('<a:grpFill/>') +
      '</p:grpSp>'
    const outer =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="7" name="G"/></p:nvGrpSpPr>' +
      `<p:grpSpPr>${GRP_XFRM}<a:solidFill><a:srgbClr val="112233"/></a:solidFill></p:grpSpPr>` +
      inner +
      '</p:grpSp>'
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(outer), ctx: {} })
    const g = (slide.elements[0] as any).children[0]
    expect(g.children[0].fill).toEqual({ type: 'solid', color: '#112233' })
  })

  it('grpFill with no group fill leaves the shape on its normal fallback chain', () => {
    const grp =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="7" name="G"/></p:nvGrpSpPr>' +
      `<p:grpSpPr>${GRP_XFRM}</p:grpSpPr>` +
      child('<a:grpFill/>', STYLE) +
      '</p:grpSp>'
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(grp), ctx: {} })
    const g = slide.elements[0] as any
    // no group fill to inherit → style fillRef template applies (no theme here → undefined phClr → undefined)
    expect(g.children[0].fill).toBeUndefined()
  })

  it('pic with grpFill records the group fill as its backdrop', () => {
    const pic =
      '<p:pic><p:nvPicPr><p:cNvPr id="9" name="Bild"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId3"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
      '<a:prstGeom prst="rect"/><a:grpFill/></p:spPr></p:pic>'
    const grp =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="7" name="G"/></p:nvGrpSpPr>' +
      `<p:grpSpPr>${GRP_XFRM}<a:solidFill><a:srgbClr val="C5DCF0"/></a:solidFill></p:grpSpPr>` +
      pic +
      '</p:grpSp>'
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(grp), ctx: {} })
    const g = slide.elements[0] as any
    expect(g.children[0].type).toBe('picture')
    expect(g.children[0].fill).toEqual({ type: 'solid', color: '#C5DCF0' })
  })
})
