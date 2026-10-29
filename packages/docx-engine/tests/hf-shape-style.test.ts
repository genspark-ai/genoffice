import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/**
 * Header/footer floating shapes carry their display transform and outline:
 * a:xfrm@rot + a:ln@w/stroke on DrawingML boxes, rotation/strokeweight/
 * strokecolor on legacy VML ones. The strip renderer draws a turned, stroked
 * banner instead of a straight bare rectangle (Word parity); the save path
 * byte-preserves anchored paragraphs, so nothing is lost on the way back.
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

const drawingmlBox = (spPrExtra: string) =>
  '<w:r><w:drawing>' +
  '<wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="page"><wp:posOffset>100000</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="page"><wp:posOffset>100000</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="2736850" cy="314325"/><wp:wrapNone/>' +
  '<wp:docPr id="7" name="Banner"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  `<wps:wsp><wps:cNvSpPr txBox="1"/><wps:spPr><a:xfrm${spPrExtra ? ' rot="2700000"' : ''}><a:off x="0" y="0"/><a:ext cx="2736850" cy="314325"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill>${spPrExtra}</wps:spPr>` +
  '<wps:txbx><w:txbxContent><w:p><w:r><w:t>Confidential</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'

const vmlBox = (styleExtra: string) =>
  `<w:r><w:pict><v:shape style="position:absolute;margin-left:20pt;margin-top:20pt;width:180pt;height:40pt;${styleExtra}">` +
  '<v:textbox><w:txbxContent><w:p><w:r><w:t>Confidential</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
  '</v:shape></w:pict></w:r>'

async function parseWithFooter(paragraphXml: string) {
  const bytes = await buildDocx({
    bodyXml: '<w:p><w:r><w:t>body</w:t></w:r></w:p>',
    extraRels:
      '<Relationship Id="rId70" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>',
    extraParts: [
      {
        path: 'word/footer1.xml',
        xml: XML_DECL + `<w:ftr ${NS}><w:p>${paragraphXml}</w:p></w:ftr>`,
        contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
      },
    ],
    sectPrExtra: '<w:footerReference w:type="default" r:id="rId70"/>',
  })
  return parseDocx(bytes)
}

describe('header/footer floating shapes keep their transform and outline', () => {
  it('parses a:xfrm rot and the a:ln stroke of a DrawingML box', async () => {
    const parsed = await parseWithFooter(
      drawingmlBox('<a:ln w="38100"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill></a:ln>'),
    )
    const box = parsed.footerParas!.find((p) => p.box)?.box!
    expect(box).toBeDefined()
    expect(box.rotDeg).toBe(45)
    expect(box.borderWidthPx).toBe(4)
    expect(box.borderColor).toBe('C00000')
  })

  it('parses rotation, strokeweight and strokecolor off a legacy VML shape', async () => {
    const parsed = await parseWithFooter(
      vmlBox('rotation:30;strokeweight:2.25pt;strokecolor:#4472C4'),
    )
    const box = parsed.footerParas!.find((p) => p.box)?.box!
    expect(box).toBeDefined()
    expect(box.rotDeg).toBe(30)
    expect(box.borderWidthPx).toBe(3)
    expect(box.borderColor).toBe('4472C4')
  })

  it('a straight unstroked box stays exactly that', async () => {
    const parsed = await parseWithFooter(drawingmlBox(''))
    const box = parsed.footerParas!.find((p) => p.box)?.box!
    expect(box).toBeDefined()
    expect(box.rotDeg).toBeUndefined()
    expect(box.borderWidthPx).toBeUndefined()
    expect(box.borderColor).toBeUndefined()
  })
})
