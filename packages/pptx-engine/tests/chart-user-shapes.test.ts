import { describe, expect, it } from 'vitest'
import { parseSlide } from '../src/parse'
import type { ChartElement } from '../src/types'

const chartXml =
  '<?xml version="1.0"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<c:chart><c:plotArea><c:barChart><c:grouping val="clustered"/>' +
  '<c:ser><c:idx val="0"/><c:order val="0"/><c:val><c:numRef><c:numCache><c:ptCount val="2"/>' +
  '<c:pt idx="0"><c:v>1</c:v></c:pt><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>' +
  '</c:barChart><c:catAx><c:axId val="1"/></c:catAx><c:valAx><c:axId val="2"/></c:valAx></c:plotArea></c:chart></c:chartSpace>'

const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
const userShapes =
  '<c:userShapes xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:cdr="http://schemas.openxmlformats.org/drawingml/2006/chartDrawing">' +
  '<cdr:relSizeAnchor><cdr:from><cdr:x>0.25</cdr:x><cdr:y>0.5</cdr:y></cdr:from><cdr:to><cdr:x>0.75</cdr:x><cdr:y>1</cdr:y></cdr:to>' +
  '<cdr:sp macro="" textlink=""><cdr:nvSpPr><cdr:cNvPr id="2" name="Note"/><cdr:cNvSpPr txBox="1"/></cdr:nvSpPr>' +
  `<cdr:spPr><a:xfrm ${A}><a:off x="1" y="1"/><a:ext cx="1" cy="1"/></a:xfrm><a:prstGeom ${A} prst="rect"><a:avLst/></a:prstGeom><a:solidFill ${A}><a:srgbClr val="FF0000"/></a:solidFill></cdr:spPr>` +
  `<cdr:txBody><a:bodyPr ${A}/><a:p ${A}><a:r><a:rPr lang="en-US" sz="1100"/><a:t>Current Month</a:t></a:r></a:p></cdr:txBody></cdr:sp></cdr:relSizeAnchor>` +
  '<cdr:relSizeAnchor><cdr:from><cdr:x>0</cdr:x><cdr:y>0</cdr:y></cdr:from><cdr:to><cdr:x>1</cdr:x><cdr:y>0</cdr:y></cdr:to>' +
  `<cdr:sp macro="" textlink=""><cdr:nvSpPr><cdr:cNvPr id="3" name="Line"/><cdr:cNvSpPr/></cdr:nvSpPr><cdr:spPr><a:prstGeom ${A} prst="line"><a:avLst/></a:prstGeom><a:ln ${A}><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:ln></cdr:spPr></cdr:sp></cdr:relSizeAnchor>` +
  '<cdr:absSizeAnchor><cdr:from><cdr:x>0.5</cdr:x><cdr:y>0.5</cdr:y></cdr:from><cdr:ext cx="200000" cy="100000"/>' +
  `<cdr:sp macro="" textlink=""><cdr:nvSpPr><cdr:cNvPr id="4" name="Abs"/><cdr:cNvSpPr/></cdr:nvSpPr><cdr:spPr><a:prstGeom ${A} prst="rect"><a:avLst/></a:prstGeom></cdr:spPr>` +
  `<cdr:txBody><a:bodyPr ${A}/><a:p ${A}><a:r><a:t>abs</a:t></a:r></a:p></cdr:txBody></cdr:sp></cdr:absSizeAnchor>` +
  '</c:userShapes>'

const slideXml =
  '<?xml version="1.0"?><p:sld xmlns:p="p" xmlns:a="a" xmlns:r="r" xmlns:c="c"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/>' +
  '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Chart 1"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>' +
  '<p:xfrm><a:off x="1000000" y="2000000"/><a:ext cx="4000000" cy="2000000"/></p:xfrm>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rId1"/></a:graphicData></a:graphic>' +
  '</p:graphicFrame></p:spTree></p:cSld></p:sld>'

describe('chart userShapes overlays', () => {
  it('text boxes map frame fractions / EMU anchors into slide coordinates; lines stay userLines', () => {
    const slide = parseSlide({
      path: 'ppt/slides/slide1.xml',
      slideXml,
      ctx: {
        chartXmls: new Map([['rId1', chartXml]]),
        chartUserShapes: new Map([['rId1', userShapes]]),
      },
    })
    const chart = slide.elements[0] as ChartElement
    expect(chart.type).toBe('chart')
    expect(chart.chart.userLines).toHaveLength(1)
    const us = chart.userShapes!
    expect(us).toHaveLength(2)
    expect(us[0]!.transform.offset).toEqual({ x: 2000000, y: 3000000, cx: 2000000, cy: 1000000 })
    expect(us[0]!.fill).toEqual({ type: 'solid', color: '#FF0000' })
    expect(us[0]!.text!.paragraphs[0]!.runs[0]!.text).toBe('Current Month')
    expect(us[1]!.transform.offset).toEqual({ x: 3000000, y: 3000000, cx: 200000, cy: 100000 })
    // overlays are render-only: the slide's own element list is untouched
    expect(slide.elements).toHaveLength(1)
  })
})
