/** SmartArt drawing text: raw "\n" inside one a:t is a paragraph break, <a:br/> stays a soft break. */
import { describe, it, expect } from 'vitest'
import { parseSlide } from '../src/parse'
import type { PassthroughElement, TextElement } from '../src/types'

const slideWith = (body: string) =>
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld>' +
  `<p:spTree><p:nvGrpSpPr/><p:grpSpPr/>${body}</p:spTree></p:cSld></p:sld>`

const FRAME =
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Diagram 1"/></p:nvGraphicFramePr>' +
  '<p:xfrm><a:off x="100" y="200"/><a:ext cx="4000" cy="3000"/></p:xfrm>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
  '<dgm:relIds xmlns:dgm="dgm" r:dm="rId1" r:lo="rId2" r:qs="rId3" r:cs="rId4"/>' +
  '</a:graphicData></a:graphic></p:graphicFrame>'

const drawing = (runs: string) =>
  '<dsp:drawing xmlns:dsp="dsp" xmlns:a="a"><dsp:spTree><dsp:nvGrpSpPr/><dsp:grpSpPr/>' +
  '<dsp:sp modelId="{X}"><dsp:nvSpPr><dsp:cNvPr id="1" name=""/><dsp:cNvSpPr/></dsp:nvSpPr>' +
  '<dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000" cy="1000"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></dsp:spPr>' +
  '<dsp:txBody><a:bodyPr/><a:p><a:pPr marL="228600" indent="-228600"><a:buChar char="&#8226;"/></a:pPr>' +
  `${runs}</a:p></dsp:txBody></dsp:sp></dsp:spTree></dsp:drawing>`

const paragraphsOf = (dsp: string) => {
  const slide = parseSlide({
    path: 'ppt/slides/slide1.xml',
    slideXml: slideWith(FRAME),
    ctx: { diagramDrawings: new Map([['rId1', dsp]]) },
  })
  const el = slide.elements[0] as PassthroughElement
  return (el.previewShapes![0] as TextElement).text!.paragraphs
}

describe('SmartArt hard returns', () => {
  it('splits a run with raw newlines into bulleted paragraphs (PowerPoint regenerates from the data model)', () => {
    const paras = paragraphsOf(
      drawing('<a:r><a:rPr lang="en-US" sz="2000"/><a:t>One.\nTwo.\nThree.</a:t></a:r>'),
    )
    expect(paras.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['One.', 'Two.', 'Three.'])
    expect(paras.every((p) => p.bullet?.char === '•' && p.marL === 228600)).toBe(true)
    expect(paras.every((p) => p.runs[0]!.fontSize === 20)).toBe(true)
  })

  it('also splits an a:t carrying attributes (xml:space="preserve")', () => {
    const paras = paragraphsOf(
      drawing('<a:r><a:rPr lang="en-US"/><a:t xml:space="preserve">One. \nTwo.</a:t></a:r>'),
    )
    expect(paras.map((p) => p.runs.map((r) => r.text).join(''))).toEqual(['One. ', 'Two.'])
  })

  it('leaves separate runs of one paragraph alone', () => {
    const paras = paragraphsOf(drawing('<a:r><a:t>One</a:t></a:r><a:br/><a:r><a:t>Two</a:t></a:r>'))
    expect(paras).toHaveLength(1)
    expect(paras[0]!.runs.map((r) => r.text)).toEqual(['One', 'Two'])
  })
})
