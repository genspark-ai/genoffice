import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'

const slideWith = (bodyShapes: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${bodyShapes}</p:spTree></p:cSld></p:sld>`

const pic = (blipInner: string, spPrFill = '') =>
  '<p:pic><p:nvPicPr><p:cNvPr id="14" name="Picture 13"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
  `<p:blipFill><a:blip r:embed="rId2">${blipInner}</a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
  `<a:prstGeom prst="rect"/>${spPrFill}</p:spPr></p:pic>`

describe('p:pic spPr fill backdrop (napierone 0027 darkened map)', () => {
  it('parses the pic solidFill alongside a translucent blip', () => {
    const xml = slideWith(
      pic('<a:alphaModFix amt="26000"/>', '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'),
    )
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: xml, ctx: {} })
    const el = slide.elements[0] as any
    expect(el.type).toBe('picture')
    expect(el.opacity).toBeCloseTo(0.26, 5)
    expect(el.fill).toEqual({ type: 'solid', color: '#000000' })
  })
})

describe('p:pic blip clrChange/duotone (napierone 0034 EMF silhouette)', () => {
  it('parses clrChange color-to-transparent and duotone on a pic blip', () => {
    const xml = slideWith(
      pic(
        '<a:clrChange><a:clrFrom><a:srgbClr val="000000"/></a:clrFrom>' +
          '<a:clrTo><a:srgbClr val="000000"><a:alpha val="0"/></a:srgbClr></a:clrTo></a:clrChange>' +
          '<a:duotone><a:srgbClr val="7C1E4F"/><a:prstClr val="white"/></a:duotone>',
      ),
    )
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: xml, ctx: {} })
    const el = slide.elements[0] as any
    expect(el.clrChange).toEqual({ from: '#000000', to: '#00000000' })
    expect(el.duotone).toEqual(['#7C1E4F', '#FFFFFF'])
  })

  it('parses clrChange on a shape blipFill too', () => {
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm><a:prstGeom prst="rect"/>' +
      '<a:blipFill><a:blip r:embed="rId2">' +
      '<a:clrChange><a:clrFrom><a:srgbClr val="FFFFFF"/></a:clrFrom>' +
      '<a:clrTo><a:srgbClr val="FF0000"/></a:clrTo></a:clrChange>' +
      '</a:blip><a:stretch><a:fillRect/></a:stretch></a:blipFill></p:spPr></p:sp>'
    const ctx = { mediaRels: new Map([['rId2', 'ppt/media/image1.png']]) }
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx })
    const el = slide.elements[0] as any
    expect(el.fill.type).toBe('image')
    expect(el.fill.clrChange).toEqual({ from: '#FFFFFF', to: '#FF0000' })
  })
})

describe('buAutoNum startAt (PlanS numbered headings)', () => {
  it('parses startAt into the bullet', () => {
    const sp =
      '<p:sp><p:nvSpPr><p:cNvPr id="5" name="T"/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
      '<p:txBody><a:bodyPr/><a:p><a:pPr marL="457200" indent="-457200">' +
      '<a:buAutoNum type="arabicPeriod" startAt="3"/></a:pPr>' +
      '<a:r><a:t>10 Principles</a:t></a:r></a:p></p:txBody></p:sp>'
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx: {} })
    const el = slide.elements[0] as any
    expect(el.text.paragraphs[0].bullet).toMatchObject({
      type: 'number',
      numType: 'arabicPeriod',
      startAt: 3,
    })
  })
})
