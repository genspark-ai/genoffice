import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import type { TextElement } from '../src/types'

const slideWith = (sp: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${sp}</p:spTree></p:cSld></p:sld>`

const parseEl = (bodyPr: string, runs: string) => {
  const sp =
    '<p:sp><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="100"/></a:xfrm></p:spPr>' +
    `<p:txBody>${bodyPr}<a:p>${runs}</a:p></p:txBody></p:sp>`
  const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: slideWith(sp), ctx: {} })
  return slide.elements[0] as TextElement
}

describe('WordArt run effects', () => {
  it('parses gradient text fill with a mid-stop fallback color', () => {
    const el = parseEl(
      '<a:bodyPr/>',
      '<a:r><a:rPr lang="en"><a:gradFill><a:gsLst>' +
        '<a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
        '<a:gs pos="50000"><a:srgbClr val="00FF00"/></a:gs>' +
        '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs>' +
        '</a:gsLst><a:lin ang="5400000"/></a:gradFill></a:rPr><a:t>W</a:t></a:r>',
    )
    const r = el.text!.paragraphs[0]!.runs[0]!
    expect(r.gradient?.stops.map((s) => s.color)).toEqual(['#FF0000', '#00FF00', '#0000FF'])
    expect(r.gradient?.angle).toBe(5400000)
    expect(r.color).toBe('#00FF00')
  })

  it('parses run glow and reflection', () => {
    const el = parseEl(
      '<a:bodyPr/>',
      '<a:r><a:rPr lang="en"><a:effectLst>' +
        '<a:glow rad="53100"><a:srgbClr val="ED7D31"/></a:glow>' +
        '<a:reflection blurRad="12700" stA="50000"/>' +
        '</a:effectLst></a:rPr><a:t>W</a:t></a:r>',
    )
    const r = el.text!.paragraphs[0]!.runs[0]!
    expect(r.glow).toEqual({ color: '#ED7D31', radius: 53100 })
    expect(r.reflection).toBe(true)
  })

  it('parses bodyPr text extrusion (sp3d + camera tilt)', () => {
    const el = parseEl(
      '<a:bodyPr><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="1800000" lon="0" rev="0"/></a:camera>' +
        '<a:lightRig rig="threePt" dir="t"/></a:scene3d>' +
        '<a:sp3d extrusionH="127000"><a:extrusionClr><a:srgbClr val="C0504D"/></a:extrusionClr></a:sp3d></a:bodyPr>',
      '<a:r><a:rPr lang="en"/><a:t>W</a:t></a:r>',
    )
    expect(el.text!.extrusion3d).toEqual({
      color: '#C0504D',
      depthEmu: 127000,
      latDeg: 30,
      lonDeg: 0,
    })
  })

  it('no effects → fields absent', () => {
    const el = parseEl('<a:bodyPr/>', '<a:r><a:rPr lang="en"/><a:t>W</a:t></a:r>')
    const r = el.text!.paragraphs[0]!.runs[0]!
    expect(r.gradient).toBeUndefined()
    expect(r.glow).toBeUndefined()
    expect(r.reflection).toBeUndefined()
    expect(el.text!.extrusion3d).toBeUndefined()
  })
})
