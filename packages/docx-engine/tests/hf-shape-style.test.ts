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

const vmlBox = (styleExtra: string, shapeAttrs = '', preChildren = '') =>
  `<w:r><w:pict><v:shape${shapeAttrs} style="position:absolute;margin-left:20pt;margin-top:20pt;width:180pt;height:40pt;${styleExtra}">` +
  `${preChildren}<v:textbox><w:txbxContent><w:p><w:r><w:t>Confidential</w:t></w:r></w:p></w:txbxContent></v:textbox>` +
  '</v:shape></w:pict></w:r>'

/** First floating box of the parsed footer strip (throws a clear failure when absent). */
function firstBox(parsed: Awaited<ReturnType<typeof parseDocx>>) {
  const box = parsed.footerParas?.find((p) => p.box)?.box
  expect(box).toBeDefined()
  return box!
}

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
    const box = firstBox(parsed)
    expect(box.rotDeg).toBe(45)
    expect(box.borderWidthPx).toBe(4)
    expect(box.borderColor).toBe('C00000')
  })

  // stroke lives on the shape's attributes (the style-attribute form is one
  // Word never writes), gated like the body path: on unless stroked="f" or a
  // v:stroke child says off, defaulting to black
  it('parses strokeweight and strokecolor off the VML shape attributes', async () => {
    const parsed = await parseWithFooter(
      vmlBox('rotation:30', ' strokecolor="#4472C4" strokeweight="2.25pt"'),
    )
    const box = firstBox(parsed)
    expect(box.rotDeg).toBe(30)
    expect(box.borderWidthPx).toBe(3)
    expect(box.borderColor).toBe('4472C4')
  })

  it('a VML box with no stroke attrs keeps the default black frame', async () => {
    const parsed = await parseWithFooter(vmlBox('rotation:30'))
    const box = firstBox(parsed)
    expect(box.rotDeg).toBe(30)
    expect(box.borderColor).toBe('000000')
    expect(box.borderWidthPx).toBeUndefined()
  })

  it('stroked="f" and a v:stroke on="f" child both leave the box bare', async () => {
    const attrOff = await parseWithFooter(vmlBox('', ' stroked="f"'))
    const offBox = firstBox(attrOff)
    expect(offBox.borderColor).toBeUndefined()
    expect(offBox.borderWidthPx).toBeUndefined()
    const childOff = await parseWithFooter(vmlBox('', '', '<v:stroke on="f"/>'))
    const strokeOffBox = firstBox(childOff)
    expect(strokeOffBox.borderColor).toBeUndefined()
    expect(strokeOffBox.borderWidthPx).toBeUndefined()
  })

  // legacy VML expresses rotation as plain degrees, as negative/360+ turns, or
  // with the fd suffix (1/65536-degree units). Header/footer boxes must go
  // through the same normalizer as body shapes, not a raw parseFloat.
  it.each([
    { style: 'rotation:90', deg: 90 },
    { style: 'rotation:-90', deg: 270 },
    { style: 'rotation:450', deg: 90 },
    { style: 'rotation:5898240fd', deg: 90 },
  ])('normalizes the legacy VML $style to $deg degrees', async ({ style, deg }) => {
    const parsed = await parseWithFooter(vmlBox(style))
    expect(firstBox(parsed).rotDeg).toBe(deg)
  })

  it('a straight unstroked box stays exactly that', async () => {
    const parsed = await parseWithFooter(drawingmlBox(''))
    const box = firstBox(parsed)
    expect(box.rotDeg).toBeUndefined()
    expect(box.borderWidthPx).toBeUndefined()
    expect(box.borderColor).toBeUndefined()
  })
})
