/** Shapes whose spPr carries neither prstGeom nor custGeom: flagged so the renderer draws text only. */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import type { TextElement } from '../src/types'

const slideWith = (body: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`

const sp = (nvPr: string, geom: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="S"/><p:cNvSpPr/>${nvPr}</p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm>${geom}` +
  '<a:solidFill><a:srgbClr val="1F497D"/></a:solidFill></p:spPr>' +
  '<p:txBody><a:bodyPr/><a:p><a:endParaRPr/></a:p></p:txBody></p:sp>'

const first = (xml: string) =>
  parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(xml), ctx: {} })
    .elements[0] as TextElement

describe('noGeometry flag', () => {
  it('is set for a non-placeholder shape without prstGeom/custGeom', () => {
    expect(first(sp('<p:nvPr/>', '')).noGeometry).toBe(true)
  })
  it('is absent with a preset geometry or on placeholders (geometry inherited from the layout)', () => {
    expect(
      first(sp('<p:nvPr/>', '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>')).noGeometry,
    ).toBeUndefined()
    expect(first(sp('<p:nvPr><p:ph type="body" idx="1"/></p:nvPr>', '')).noGeometry).toBeUndefined()
  })
})
