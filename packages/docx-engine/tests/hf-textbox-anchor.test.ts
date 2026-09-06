import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const HEADER_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml'
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:w10="urn:schemas-microsoft-com:office:word"'

/** PDF-converter header: a near-zero anchor paragraph carrying page-anchored textboxes */
const anchoredBox = (x: number, y: number, text: string) =>
  '<w:r><w:drawing>' +
  '<wp:anchor behindDoc="1" relativeHeight="1" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  `<wp:positionH relativeFrom="page"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="page"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
  '<wp:extent cx="1855470" cy="154940"/><wp:wrapNone/><wp:docPr id="1" name="Textbox 1"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1855470" cy="154940"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>' +
  `<wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:before="18"/></w:pPr><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent></wps:txbx>` +
  '<wps:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"><a:noAutofit/></wps:bodyPr>' +
  '</wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'

async function parseHeader(headerBody: string) {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraRels:
      '<Relationship Id="rId61" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    extraParts: [
      {
        path: 'word/header1.xml',
        xml: `<w:hdr ${NS}>${headerBody}</w:hdr>`,
        contentType: HEADER_TYPE,
      },
    ],
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId61"/>',
  })
  return (await parseDocx(bytes)).headerParas ?? []
}

describe('header/footer floating textboxes carry their anchor geometry', () => {
  it('page-anchored wps textboxes: position, extent, wrap, insets and vertical anchor', async () => {
    const paras = await parseHeader(
      '<w:p><w:pPr><w:spacing w:line="14" w:lineRule="auto"/></w:pPr>' +
        anchoredBox(635297, 318638, 'Municipal Corporation') +
        anchoredBox(4880000, 318638, 'Tender No.') +
        '</w:p>',
    )
    // the anchor paragraph keeps its own line (collapsed to w:line=14 here):
    // Word starts the body at headerDist + that line, not at a bare top margin
    expect(paras).toHaveLength(3)
    expect(paras[0]).toMatchObject({ runs: [], lineRule: 'auto', lineRawTwips: 14 })
    expect(paras[0].box).toBeUndefined()
    expect(paras[1].boxAnchored).toBe(true)
    expect(paras[1].box).toMatchObject({
      behind: true,
      wrap: 'none',
      widthPx: 195,
      heightPx: 16,
      posXPx: 67,
      posHRel: 'page',
      posYPx: 33,
      posVRel: 'page',
      insets: [0, 0, 0, 0],
      vAlign: 'center',
      anchorPara: 0,
    })
    expect(paras[2].box?.posXPx).toBe(512)
    expect(paras[2].box?.id).not.toBe(paras[1].box?.id)
  })

  it('an absolutely positioned VML shape reads its style geometry; inline textboxes carry no box', async () => {
    const paras = await parseHeader(
      '<w:p><w:r><w:pict>' +
        '<v:shape id="s1" type="#_x0000_t202" style="position:absolute;margin-left:50pt;margin-top:25.1pt;width:146.1pt;height:12.2pt;z-index:-1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page">' +
        '<v:textbox inset="0,0,0,0"><w:txbxContent><w:p><w:r><w:t>VML box</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
        '<w10:wrap type="none"/></v:shape></w:pict></w:r></w:p>' +
        '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="914400" cy="457200"/>' +
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><wps:wsp>' +
        '<wps:txbx><w:txbxContent><w:p><w:r><w:t>inline box</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
        '</wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>',
    )
    expect(paras).toHaveLength(3)
    expect(paras[0].runs).toEqual([])
    expect(paras[1].box).toMatchObject({
      behind: true,
      wrap: 'none',
      widthPx: 195,
      heightPx: 16,
      posXPx: 67,
      posHRel: 'page',
      posYPx: 33,
      posVRel: 'page',
      insets: [0, 0, 0, 0],
    })
    // the inline box flows in the strip: no anchor line is added for it
    expect(paras[2].boxAnchored).toBeUndefined()
    expect(paras[2].box).toBeUndefined()
  })

  it('VML lengths: unitless zero (pt) and in/cm/mm/px/pc units all place the box', async () => {
    const paras = await parseHeader(
      '<w:p><w:r><w:pict>' +
        '<v:shape id="s2" type="#_x0000_t202" style="position:absolute;margin-left:0;margin-top:0;width:2in;height:36px;mso-position-horizontal:absolute;mso-position-vertical:absolute;mso-position-horizontal-relative:page;mso-position-vertical-relative:page">' +
        '<v:textbox inset="1pc,0,1cm,2mm"><w:txbxContent><w:p><w:r><w:t>Word-style VML box</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
        '</v:shape></w:pict></w:r></w:p>',
    )
    expect(paras).toHaveLength(2)
    expect(paras[1].box).toMatchObject({
      posXPx: 0,
      posHRel: 'page',
      posYPx: 0,
      posVRel: 'page',
      widthPx: 192,
      heightPx: 36,
      insets: [16, 0, 37.8, 7.56],
    })
    expect(paras[0].box?.posH).toBeUndefined()
  })
})
