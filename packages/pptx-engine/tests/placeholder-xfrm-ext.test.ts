/**
 * A placeholder whose <a:xfrm> carries only <a:off> (no <a:ext>) keeps its own offset
 * but takes the size from the layout/master placeholder; a zero-size box is never
 * what the authoring app meant.
 */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { parsePlaceholderMap } from '../src/placeholder'

const layoutXml =
  '<?xml version="1.0"?><p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>' +
  '<p:nvGrpSpPr/><p:grpSpPr/>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/>' +
  '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="8000000" cy="1000000"/></a:xfrm></p:spPr>' +
  '<p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="5" name="Picture Placeholder 2"/><p:cNvSpPr/>' +
  '<p:nvPr><p:ph type="pic" idx="12"/></p:nvPr></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="300" y="400"/><a:ext cx="5000000" cy="3000000"/></a:xfrm></p:spPr>' +
  '<p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>' +
  '</p:spTree></p:cSld></p:sldLayout>'

const slideXml =
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>' +
  '<p:nvGrpSpPr/><p:grpSpPr/>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/>' +
  '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>' +
  '<p:spPr><a:xfrm><a:off x="1234" y="5678"/></a:xfrm></p:spPr>' +
  '<p:txBody><a:bodyPr/><a:p><a:r><a:t>Hi</a:t></a:r></a:p></p:txBody></p:sp>' +
  '<p:pic><p:nvPicPr><p:cNvPr id="27" name="Picture 26"/><p:cNvPicPr/>' +
  '<p:nvPr><p:ph type="pic" idx="12"/></p:nvPr></p:nvPicPr>' +
  '<p:blipFill><a:blip r:embed="rId3"/><a:stretch/></p:blipFill>' +
  '<p:spPr><a:xfrm><a:off x="42" y="43"/></a:xfrm></p:spPr></p:pic>' +
  '</p:spTree></p:cSld></p:sld>'

const parse = (xml: string) =>
  parseSlide({
    path: 'ppt/slides/slide1.xml',
    slideXml: xml,
    ctx: { layoutPlaceholders: parsePlaceholderMap(layoutXml) },
  })

describe('placeholder xfrm without a:ext', () => {
  it('shape keeps its own offset and inherits the layout size', () => {
    const el = parse(slideXml).elements[0] as any
    expect(el.transform.offset).toEqual({ x: 1234, y: 5678, cx: 8000000, cy: 1000000 })
  })

  it('picture keeps its own offset and inherits the layout size', () => {
    const el = parse(slideXml).elements[1] as any
    expect(el.type).toBe('picture')
    expect(el.transform.offset).toEqual({ x: 42, y: 43, cx: 5000000, cy: 3000000 })
  })

  it('a full own xfrm is left alone', () => {
    const own = slideXml.replace(
      '<a:off x="1234" y="5678"/>',
      '<a:off x="1234" y="5678"/><a:ext cx="10" cy="20"/>',
    )
    const el = parse(own).elements[0] as any
    expect(el.transform.offset).toEqual({ x: 1234, y: 5678, cx: 10, cy: 20 })
  })

  it('a non-placeholder shape without a:ext stays zero-size', () => {
    const noPh = slideXml.replace('<p:nvPr><p:ph type="title"/></p:nvPr>', '<p:nvPr/>')
    const el = parse(noPh).elements[0] as any
    expect(el.transform.offset).toEqual({ x: 1234, y: 5678, cx: 0, cy: 0 })
  })
})
