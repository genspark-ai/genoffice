/**
 * Placeholder pictures without their own <a:prstGeom> inherit the layout/master
 * placeholder's shape (with adjustment values), so the image clips like PowerPoint
 * (e.g. parallelogram picture placeholders in cover layouts).
 */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import { parsePlaceholderMap } from '../src/placeholder'

const layoutXml =
  '<?xml version="1.0"?><p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>' +
  '<p:nvGrpSpPr/><p:grpSpPr/>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="5" name="Picture Placeholder 2"/><p:cNvSpPr/>' +
  '<p:nvPr><p:ph type="pic" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr>' +
  '<p:spPr><a:xfrm flipH="1"><a:off x="3918353" y="-12"/><a:ext cx="17689293" cy="335935"/></a:xfrm>' +
  '<a:prstGeom prst="parallelogram"><a:avLst><a:gd name="adj" fmla="val 88235"/></a:avLst></a:prstGeom>' +
  '</p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>' +
  '</p:spTree></p:cSld></p:sldLayout>'

const slideXml =
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>' +
  '<p:nvGrpSpPr/><p:grpSpPr/>' +
  '<p:pic><p:nvPicPr><p:cNvPr id="27" name="Picture Placeholder 26"/><p:cNvPicPr/>' +
  '<p:nvPr><p:ph type="pic" sz="quarter" idx="12"/></p:nvPr></p:nvPicPr>' +
  '<p:blipFill><a:blip r:embed="rId3"/><a:stretch/></p:blipFill>' +
  '<p:spPr><a:xfrm flipH="1"><a:off x="3918353" y="-12"/><a:ext cx="17689293" cy="10286993"/></a:xfrm></p:spPr></p:pic>' +
  '</p:spTree></p:cSld></p:sld>'

describe('placeholder picture geometry inheritance', () => {
  it('pic without own prstGeom inherits the layout placeholder shape and adjust values', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml,
      ctx: { layoutPlaceholders: parsePlaceholderMap(layoutXml) },
    })
    const el = slide.elements[0] as any
    expect(el.type).toBe('picture')
    expect(el.presetGeometry).toBe('parallelogram')
    expect(el.adjust).toEqual({ adj: 88235 })
  })

  it('pic with its own geometry keeps it (no inheritance override)', () => {
    const own = slideXml.replace(
      '</a:xfrm></p:spPr></p:pic>',
      '</a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></p:spPr></p:pic>',
    )
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: own,
      ctx: { layoutPlaceholders: parsePlaceholderMap(layoutXml) },
    })
    expect((slide.elements[0] as any).presetGeometry).toBe('ellipse')
  })

  it('non-placeholder pic does not inherit', () => {
    const noPh = slideXml.replace('<p:ph type="pic" sz="quarter" idx="12"/>', '')
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: noPh,
      ctx: { layoutPlaceholders: parsePlaceholderMap(layoutXml) },
    })
    expect((slide.elements[0] as any).presetGeometry).toBeUndefined()
  })
})
