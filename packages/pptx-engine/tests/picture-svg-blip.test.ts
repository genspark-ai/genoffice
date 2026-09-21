/** SVG-only pictures: <a:blip> without r:embed whose image is the asvg:svgBlip extension. */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import type { PictureElement, TextElement } from '../src/types'

const SVG_EXT =
  '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">' +
  '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="rId4"/>' +
  '</a:ext></a:extLst>'

const slideWith = (inner: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${inner}</p:spTree></p:cSld></p:sld>`
const pic = (blip: string) =>
  '<p:pic><p:nvPicPr><p:cNvPr id="7" name="P"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
  `<p:blipFill>${blip}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr></p:pic>'

const mediaRels = new Map([
  ['rId2', 'ppt/media/image1.png'],
  ['rId4', 'ppt/media/image2.svg'],
])
const parseOne = (inner: string) =>
  parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(inner), ctx: { mediaRels } })

describe('svgBlip fallback', () => {
  it('SVG-only picture (blip without r:embed) resolves media via asvg:svgBlip', () => {
    const slide = parseOne(pic(`<a:blip>${SVG_EXT}</a:blip>`))
    const el = slide.elements[0] as PictureElement
    expect(el.mediaRef).toBe('ppt/media/image2.svg')
  })

  it('svgBlip wins over the raster r:embed fallback when both are present', () => {
    // PowerPoint 2016+ draws the vector; the raster is only for older readers (a prod
    // deck's fallback PNG is a blank white square where the SVG logo should be)
    const slide = parseOne(pic(`<a:blip r:embed="rId2">${SVG_EXT}</a:blip>`))
    const el = slide.elements[0] as PictureElement
    expect(el.mediaRef).toBe('ppt/media/image2.svg')
  })

  it('shape blipFill fill resolves media via asvg:svgBlip', () => {
    const shape =
      '<p:sp><p:nvSpPr><p:cNvPr id="8" name="S"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>' +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip>${SVG_EXT}</a:blip><a:stretch/></a:blipFill></p:spPr>` +
      '<p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>'
    const slide = parseOne(shape)
    const el = slide.elements[0] as TextElement
    expect(el.fill).toEqual({ type: 'image', mediaRef: 'ppt/media/image2.svg', mode: 'stretch' })
  })
})
