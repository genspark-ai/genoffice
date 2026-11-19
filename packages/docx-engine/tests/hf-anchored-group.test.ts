import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const NS =
  ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'

const pic = (off: string, srcRect = '') =>
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Logo"/><pic:cNvPicPr/></pic:nvPicPr>' +
  `<pic:blipFill><a:blip r:embed="rId1"/>${srcRect}<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
  `<pic:spPr><a:xfrm>${off}<a:ext cx="1828800" cy="1828800"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'

/** header: one paragraph carrying only an anchored two-picture group (child space scaled 1:2) */
const HEADER_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:hdr${NS}><w:p><w:pPr><w:pStyle w:val="Header"/></w:pPr><w:r><w:drawing>` +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>914400</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>-457200</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1828800" cy="914400"/><wp:wrapNone/><wp:docPr id="7" name="Group 7"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
  '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/>' +
  '<a:chOff x="0" y="0"/><a:chExt cx="3657600" cy="1828800"/></a:xfrm></wpg:grpSpPr>' +
  pic('<a:off x="0" y="0"/>', '<a:srcRect r="50000"/>') +
  pic('<a:off x="1828800" y="0"/>') +
  '</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:hdr>'

const HEADER_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
  '</Relationships>'

const HEADER_STYLE =
  '<w:style w:type="paragraph" w:styleId="Header"><w:name w:val="header"/><w:basedOn w:val="Normal"/>' +
  '<w:pPr><w:spacing w:line="240" w:lineRule="auto"/></w:pPr></w:style>'

async function build(): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
    withImage: true,
    extraStylesXml: HEADER_STYLE,
    extraRels:
      '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
    extraParts: [
      {
        path: 'word/header1.xml',
        xml: HEADER_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
      },
      {
        path: 'word/_rels/header1.xml.rels',
        xml: HEADER_RELS,
        contentType: 'application/vnd.openxmlformats-package.relationships+xml',
      },
    ],
  })
}

describe('anchored picture group in a header', () => {
  it('splits into one image per child picture at its mapped offset, size and crop', async () => {
    const doc = await parseDocx(await build())
    const imgs = doc.headerImages ?? []
    expect(imgs).toHaveLength(2)
    // child space 3657600x1828800 maps onto the 1828800x914400 group: scale 1/2
    expect(imgs[0]).toMatchObject({
      floating: true,
      wrap: 'none',
      widthPx: 96,
      heightPx: 96,
      posXPx: 96,
      posYPx: -48,
      posHRel: 'margin',
      posVRel: 'paragraph',
      crop: { l: 0, t: 0, r: 0.5, b: 0 },
    })
    expect(imgs[1]).toMatchObject({ widthPx: 96, heightPx: 96, posXPx: 192, posYPx: -48 })
    expect(imgs[1].crop).toBeUndefined()
  })

  it('the anchor-only paragraph keeps its (style-sized) line so the body starts below it', async () => {
    const doc = await parseDocx(await build())
    const part = Object.values(doc.hfParts ?? {})[0]
    expect(part?.paras).toHaveLength(1)
    expect(part?.paras[0]).toMatchObject({ runs: [], lineRule: 'auto', lineRawTwips: 240 })
  })
})
