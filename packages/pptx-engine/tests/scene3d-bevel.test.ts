import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'

const slide = (sp3d: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>' +
  '<p:sp><p:nvSpPr><p:cNvPr id="2" name="R"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>' +
  '<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>' +
  '<a:solidFill><a:srgbClr val="5B9BD5"/></a:solidFill>' +
  '<a:scene3d><a:camera prst="orthographicFront"/><a:lightRig rig="chilly" dir="t"/></a:scene3d>' +
  sp3d +
  '</p:spPr></p:sp></p:spTree></p:cSld></p:sld>'

describe('sp3d top bevel', () => {
  it('records width, height and preset', () => {
    const s = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slide(
        '<a:sp3d prstMaterial="translucentPowder"><a:bevelT w="127000" h="25400" prst="softRound"/></a:sp3d>',
      ),
      ctx: {},
    })
    const el = s.elements[0] as any
    expect(el.scene3d.material).toBe('translucentPowder')
    expect(el.scene3d.bevelTop).toEqual({ wEmu: 127000, hEmu: 25400, preset: 'softRound' })
  })
  it('defaults an attribute-less bevelT to the 76200 circle preset', () => {
    const s = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slide('<a:sp3d><a:bevelT/></a:sp3d>'),
      ctx: {},
    })
    expect((s.elements[0] as any).scene3d.bevelTop).toEqual({
      wEmu: 76200,
      hEmu: 76200,
      preset: 'circle',
    })
  })
  it('no bevelT: no bevel recorded', () => {
    const s = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml: slide('<a:sp3d prstMaterial="matte"/>'),
      ctx: {},
    })
    expect((s.elements[0] as any).scene3d.bevelTop).toBeUndefined()
  })
})
