import { describe, it, expect } from 'vitest'
import { parseSlide, parseDecorations } from '../src/parse'
import type { TextElement, GroupElement } from '../src/types'

const wrap = (spTree: string) =>
  `<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r"><p:cSld><p:spTree>` +
  `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>` +
  spTree +
  `</p:spTree></p:cSld></p:sld>`

const CXN = `<p:cxnSp>
  <p:nvCxnSpPr><p:cNvPr id="4" name="Straight Arrow Connector 3"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
  <p:spPr>
    <a:xfrm flipV="1"><a:off x="100" y="200"/><a:ext cx="914400" cy="457200"/></a:xfrm>
    <a:prstGeom prst="bentConnector3"><a:avLst><a:gd name="adj1" fmla="val 75000"/></a:avLst></a:prstGeom>
    <a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/><a:tailEnd type="triangle"/></a:ln>
  </p:spPr>
</p:cxnSp>`

describe('p:cxnSp connector parsing', () => {
  it('parses connector as stroked shape with geometry/adjust/arrow', () => {
    const slide = parseSlide({ path: 'ppt/slides/slide1.xml', slideXml: wrap(CXN), ctx: {} })
    expect(slide.elements).toHaveLength(1)
    const el = slide.elements[0] as TextElement
    expect(el.type).toBe('shape')
    expect(el.presetGeometry).toBe('bentConnector3')
    expect(el.adjust?.adj1).toBe(75000)
    expect(el.transform.flipV).toBe(true)
    expect(el.fill).toEqual({ type: 'none' })
    expect(el.stroke?.width).toBe(19050)
    expect(el.stroke?.fill).toEqual({ type: 'solid', color: '#FF0000' })
    expect(el.stroke?.dash).toBe('dash')
    expect(el.stroke?.tailEnd).toEqual({ type: 'triangle' })
    expect(el.stroke?.headEnd).toBeUndefined()
  })

  it('connector without explicit ln fill falls back to visible stroke', () => {
    const bare = `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="5" name="c"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>
      <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="0"/></a:xfrm>
      <a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>`
    const slide = parseSlide({ path: 'p', slideXml: wrap(bare), ctx: {} })
    const el = slide.elements[0] as TextElement
    expect(el.stroke).toBeTruthy()
    expect(el.stroke!.fill.type).toBe('solid')
  })

  it('parses connector inside a group', () => {
    const grp = `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="6" name="g"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="914400" cy="914400"/></a:xfrm></p:grpSpPr>
      ${CXN}</p:grpSp>`
    const slide = parseSlide({ path: 'p', slideXml: wrap(grp), ctx: {} })
    const g = slide.elements[0] as GroupElement
    expect(g.type).toBe('group')
    expect(g.children).toHaveLength(1)
    expect((g.children[0] as TextElement).presetGeometry).toBe('bentConnector3')
  })
})

const LAYOUT = `<?xml version="1.0"?><p:sldLayout xmlns:p="p" xmlns:a="a"><p:cSld><p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>
<p:sp>
  <p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm></p:spPr>
  <p:txBody><a:bodyPr/><a:p><a:r><a:t>Click to edit title</a:t></a:r></a:p></p:txBody>
</p:sp>
<p:sp>
  <p:nvSpPr><p:cNvPr id="3" name="logo color bar"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="12192000" cy="100000"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
    <a:solidFill><a:srgbClr val="C43E1C"/></a:solidFill></p:spPr>
</p:sp>
<p:sp>
  <p:nvSpPr><p:cNvPr id="4" name="Slide Number"/><p:cNvSpPr/><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="11000000" y="6500000"/><a:ext cx="900000" cy="300000"/></a:xfrm></p:spPr>
  <p:txBody><a:bodyPr/><a:p><a:fld id="{X}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld></a:p></p:txBody>
</p:sp>
<p:sp>
  <p:nvSpPr><p:cNvPr id="5" name="Footer"/><p:cNvSpPr/><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="4000000" y="6500000"/><a:ext cx="4000000" cy="300000"/></a:xfrm></p:spPr>
  <p:txBody><a:bodyPr/><a:p><a:r><a:t>Company Confidential</a:t></a:r></a:p></p:txBody>
</p:sp>
</p:spTree></p:cSld></p:sldLayout>`

describe('parseDecorations (master/layout decoration layer)', () => {
  it('keeps non-placeholder shapes, drops content placeholders', () => {
    const decos = parseDecorations(LAYOUT, {}, { hfTypes: new Set() })
    expect(decos).toHaveLength(1)
    expect(decos[0]!.name).toBe('logo color bar')
    expect((decos[0] as TextElement).fill).toEqual({ type: 'solid', color: '#C43E1C' })
  })

  it('includes enabled footer-family placeholders and substitutes slide number', () => {
    const decos = parseDecorations(LAYOUT, {}, { hfTypes: new Set(['ftr', 'sldNum']), slideNum: 7 })
    expect(decos).toHaveLength(3)
    const sldNum = decos.find((d) => (d as TextElement).placeholder === 'sldNum') as TextElement
    expect(sldNum).toBeTruthy()
    const runs = sldNum.text!.paragraphs.flatMap((p) => p.runs)
    expect(runs.some((r) => r.field === 'slidenum' && r.text === '7')).toBe(true)
    const ftr = decos.find((d) => (d as TextElement).placeholder === 'ftr') as TextElement
    expect(ftr.text!.paragraphs[0]!.runs[0]!.text).toBe('Company Confidential')
  })

  it('subset of hfTypes filters accordingly', () => {
    const decos = parseDecorations(LAYOUT, {}, { hfTypes: new Set(['ftr']) })
    expect(decos.map((d) => (d as TextElement).placeholder ?? 'none').sort()).toEqual(
      ['ftr', 'none'].sort(),
    )
  })
})
