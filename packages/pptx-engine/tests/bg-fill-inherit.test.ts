import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { parsePlaceholderMap } from '../src/placeholder'

const slideWith = (bodyShapes: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${bodyShapes}</p:spTree></p:cSld></p:sld>`

describe('useBgFill (tdf93868)', () => {
  it('flags <p:sp useBgFill="1"> so the render layer substitutes the slide background', () => {
    const sp =
      '<p:sp useBgFill="1"><p:nvSpPr><p:cNvPr id="2" name="R"/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
      '<a:prstGeom prst="roundRect"/><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:spPr></p:sp>'
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx: {} })
    const el = slide.elements[0] as any
    expect(el.useBgFill).toBe(true)
    // spPr fill is kept only as a fallback
    expect(el.fill).toEqual({ type: 'solid', color: '#FFFFFF' })
  })

  it('leaves plain shapes unflagged', () => {
    const sp =
      '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
      '<a:prstGeom prst="rect"/><a:noFill/></p:spPr></p:sp>'
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx: {} })
    expect((slide.elements[0] as any).useBgFill).toBeUndefined()
  })
})

describe('p:pic placeholder geometry backfill (customshape-bitmapfill-srcrect)', () => {
  it('inherits layout xfrm when a picture placeholder omits <a:xfrm>', () => {
    const layoutXml =
      '<p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>' +
      '<p:sp><p:nvSpPr><p:cNvPr id="2"/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="6192000" y="1332000"/><a:ext cx="5493600" cy="4012789"/></a:xfrm></p:spPr></p:sp>' +
      '</p:spTree></p:cSld></p:sldLayout>'
    const pic =
      '<p:pic><p:nvPicPr><p:cNvPr id="6" name="Content Placeholder 5"/><p:cNvPicPr/>' +
      '<p:nvPr><p:ph idx="1"/></p:nvPr></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId2"/><a:srcRect l="4393" r="4393"/><a:stretch/></p:blipFill>' +
      '<p:spPr/></p:pic>'
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slideWith(pic),
      ctx: { layoutPlaceholders: parsePlaceholderMap(layoutXml) },
    })
    const el = slide.elements[0] as any
    expect(el.type).toBe('picture')
    expect(el.transform.offset).toEqual({ x: 6192000, y: 1332000, cx: 5493600, cy: 4012789 })
  })

  it('keeps an explicit pic xfrm untouched', () => {
    const pic =
      '<p:pic><p:nvPicPr><p:cNvPr id="6"/><p:cNvPicPr/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvPicPr>' +
      '<p:blipFill><a:blip r:embed="rId2"/><a:stretch/></p:blipFill>' +
      '<p:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="30" cy="40"/></a:xfrm></p:spPr></p:pic>'
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(pic), ctx: {} })
    expect((slide.elements[0] as any).transform.offset).toEqual({ x: 10, y: 20, cx: 30, cy: 40 })
  })
})
