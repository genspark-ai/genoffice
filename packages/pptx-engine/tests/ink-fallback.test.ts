import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'

// Ink stroke (p:contentPart) with the picture fallback PowerPoint writes next to it: the
// fallback box is a tall strip, the ink's own p14:xfrm is the 2-point-high line it draws
const slideXml =
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r" xmlns:mc="mc" xmlns:p14="p14"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>' +
  '<mc:AlternateContent><mc:Choice Requires="p14"><p:contentPart r:id="rId12"><p14:nvContentPartPr><p14:cNvPr id="29" name="Ink 28"/><p14:cNvContentPartPr/><p14:nvPr/></p14:nvContentPartPr>' +
  '<p14:xfrm><a:off x="7419208" y="5196805"/><a:ext cx="6181200" cy="23760"/></p14:xfrm></p:contentPart></mc:Choice>' +
  '<mc:Fallback><p:pic><p:nvPicPr><p:cNvPr id="29" name="Ink 28"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId13"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  '<p:spPr><a:xfrm><a:off x="7413088" y="4792885"/><a:ext cx="6193440" cy="831600"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></mc:Fallback></mc:AlternateContent>' +
  '</p:spTree></p:cSld></p:sld>'

describe('ink contentPart fallback picture', () => {
  it('takes the ink geometry (p14:xfrm), not the fallback strip', () => {
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml, ctx: {} })
    const el = slide.elements[0] as any
    expect(el.type).toBe('picture')
    expect(el.transform.offset).toEqual({ x: 7419208, y: 5196805, cx: 6181200, cy: 23760 })
  })
})
