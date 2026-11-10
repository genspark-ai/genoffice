import { describe, expect, it } from 'vitest'
import { PAGE_MARK, parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * Footers anchoring a DrawingML text box (wps:wsp with w:txbxContent inside a
 * wp:anchor) next to their own runs: the box paragraphs surface after the
 * anchor paragraph, carrying the box placement and the anchor's index, instead
 * of being dropped. Page-aligned boxes record the align's relativeFrom.
 */

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml"'

const box = (positionXml: string, wrapXml: string, text: string, bodyWrap = '') =>
  '<w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  `<wp:simplePos x="0" y="0"/>${positionXml}<wp:extent cx="2736850" cy="314325"/>${wrapXml}` +
  '<wp:docPr id="7" name="Text Box 5"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2736850" cy="314325"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></wps:spPr>' +
  `<wps:txbx><w:txbxContent><w:p><w:r><w:rPr><w:sz w:val="16"/></w:rPr><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
  `<wps:bodyPr${bodyWrap} lIns="254000" tIns="0" rIns="0" bIns="190500" anchor="b"><a:spAutoFit/></wps:bodyPr>` +
  '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
  // Word's VML twin of the same box: must not double up as strip text
  `<mc:Fallback><w:pict><v:shape style="position:absolute"><v:textbox><w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback>` +
  '</mc:AlternateContent></w:r>'

const PAGE_BOTTOM_BOX = box(
  '<wp:positionH relativeFrom="page"><wp:align>left</wp:align></wp:positionH>' +
    '<wp:positionV relativeFrom="page"><wp:align>bottom</wp:align></wp:positionV>',
  '<wp:wrapNone/>',
  'Classified as Confidential.',
  ' wrap="none"',
)

const PARAGRAPH_BOX = box(
  '<wp:positionH relativeFrom="column"><wp:posOffset>5166360</wp:posOffset></wp:positionH>' +
    '<wp:positionV relativeFrom="paragraph"><wp:posOffset>304800</wp:posOffset></wp:positionV>',
  '<wp:wrapSquare wrapText="bothSides"/>',
  'Prepared by: A. Author',
)

/** footer: page-number paragraph, then a paragraph carrying the box next to a text run */
const FOOTER_BOX_WITH_TEXT =
  XML_DECL +
  `<w:ftr ${NS}>` +
  `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:fldSimple w:instr=" PAGE "><w:r><w:t>2</w:t></w:r></w:fldSimple></w:p>` +
  `<w:p>${PARAGRAPH_BOX}<w:r><w:tab/><w:t>draft</w:t></w:r></w:p>` +
  '</w:ftr>'

/** footer whose only paragraph anchors the box (and nothing else); the
 *  anchor paragraph collapses its own line like the tender headers do */
const FOOTER_BOX_ONLY =
  XML_DECL +
  `<w:ftr ${NS}><w:p><w:pPr><w:spacing w:line="14" w:lineRule="auto"/></w:pPr>${PAGE_BOTTOM_BOX}</w:p></w:ftr>`

async function parseWithFooter(footerXml: string, extraStylesXml?: string) {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    ...(extraStylesXml ? { extraStylesXml } : {}),
    extraRels:
      '<Relationship Id="rId70" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
    extraParts: [
      {
        path: 'word/footer1.xml',
        xml: footerXml,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
      },
    ],
    sectPrExtra: '<w:footerReference w:type="default" r:id="rId70"/>',
  })
  return parseDocx(bytes)
}

describe('anchored text boxes sharing a footer paragraph with text', () => {
  it('surfaces the box after its anchor paragraph, anchored to it', async () => {
    const parsed = await parseWithFooter(FOOTER_BOX_WITH_TEXT)
    const paras = parsed.footerParas!
    expect(paras.map((p) => p.runs.map((r) => r.text).join(''))).toEqual([
      PAGE_MARK,
      '\tdraft',
      'Prepared by: A. Author',
    ])
    expect(paras[1].box).toBeUndefined()
    const b = paras[2].box!
    expect(paras[2].boxAnchored).toBe(true)
    expect(b.anchorPara).toBe(1)
    expect(b).toMatchObject({
      wrap: 'square',
      posHRel: 'margin',
      posXPx: 542,
      posVRel: 'paragraph',
      posYPx: 32,
      widthPx: 287,
      insets: [26.67, 0, 0, 20],
      vAlign: 'bottom',
    })
    expect(b.nowrap).toBeUndefined()
    // the multi-section lookup path carries the same paragraphs
    expect(parsed.hfParts?.rId70?.paras.map((p) => p.box?.anchorPara)).toEqual([
      undefined,
      undefined,
      1,
    ])
  })

  it('the anchor-only line takes its line spacing from the Footer style', async () => {
    const footer =
      XML_DECL +
      `<w:ftr ${NS}><w:p><w:pPr><w:pStyle w:val="Footer"/></w:pPr>${PAGE_BOTTOM_BOX}</w:p></w:ftr>`
    const style =
      '<w:style w:type="paragraph" w:styleId="Footer"><w:name w:val="footer"/>' +
      '<w:pPr><w:spacing w:line="14" w:lineRule="auto"/><w:jc w:val="right"/></w:pPr></w:style>'
    const parsed = await parseWithFooter(footer, style)
    const paras = parsed.footerParas!
    expect(paras[0].runs).toEqual([])
    expect(paras[0]).toMatchObject({ lineRule: 'auto', lineRawTwips: 14, align: 'right' })
    expect(paras[1].box!.anchorPara).toBe(0)
  })

  it('an inline drawing inside the box content does not suppress the anchor line', async () => {
    const inlinePic =
      '<w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="914400"/>' +
      '<wp:docPr id="9" name="pic"/><a:graphic><a:graphicData uri="x"/></a:graphic></wp:inline></w:drawing></w:r>'
    const box = PAGE_BOTTOM_BOX.replace('<w:txbxContent>', `<w:txbxContent><w:p>${inlinePic}</w:p>`)
    expect(box).not.toBe(PAGE_BOTTOM_BOX)
    const footer =
      XML_DECL +
      `<w:ftr ${NS}><w:p><w:pPr><w:spacing w:line="14" w:lineRule="auto"/></w:pPr>${box}</w:p></w:ftr>`
    const parsed = await parseWithFooter(footer)
    const paras = parsed.footerParas!
    expect(paras[0].runs).toEqual([])
    expect(paras[0]).toMatchObject({ lineRule: 'auto', lineRawTwips: 14 })
    expect(paras.slice(1).every((p) => p.box?.anchorPara === 0)).toBe(true)
  })

  it('records the relativeFrom of aligned placements and bodyPr wrap="none"', async () => {
    const parsed = await parseWithFooter(FOOTER_BOX_ONLY)
    const paras = parsed.footerParas!
    // the anchor paragraph keeps its own (collapsed) line ahead of the box
    // content; the VML fallback twin does not add a third paragraph
    expect(paras).toHaveLength(2)
    expect(paras[0].runs).toEqual([])
    expect(paras[0].box).toBeUndefined()
    expect(paras[0]).toMatchObject({ lineRule: 'auto', lineRawTwips: 14 })
    const b = paras[1].box!
    expect(b.anchorPara).toBe(0)
    expect(b).toMatchObject({
      posH: 'left',
      posHRel: 'page',
      posV: 'bottom',
      posVRel: 'page',
      wrap: 'none',
      nowrap: true,
      vAlign: 'bottom',
      autofit: true,
    })
  })

  it('leaves a box without spAutoFit fixed-size (Word clips its overflow)', async () => {
    const parsed = await parseWithFooter(
      FOOTER_BOX_ONLY.replace('<a:spAutoFit/>', '<a:noAutofit/>'),
    )
    const b = parsed.footerParas![1].box!
    expect(b.heightPx).toBeGreaterThan(0)
    expect(b.autofit).toBeUndefined()
  })
})
