import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const TRIANGLE_ANCHOR =
  '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>2505075</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="line"><wp:posOffset>40640</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="589915" cy="533400"/><wp:wrapNone/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="589915" cy="533400"/></a:xfrm>' +
  '<a:prstGeom prst="triangle"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="4472C4"/></a:solidFill></wps:spPr>' +
  '<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>'

// fixed-height (noAutofit) textbox: sysClr fill, prstClr border, exact one-line spacing
const TEXTBOX_ANCHOR =
  '<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="2" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>9525</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="1000000" cy="152400"/><wp:wrapNone/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
  '<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1000000" cy="152400"/></a:xfrm>' +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>' +
  '<a:solidFill><a:sysClr val="window" lastClr="FFFFFF"/></a:solidFill>' +
  '<a:ln w="6350"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></wps:spPr>' +
  '<wps:txbx><w:txbxContent><w:p><w:pPr><w:spacing w:line="240" w:lineRule="exact"/></w:pPr>' +
  '<w:r><w:t>box text</w:t></w:r></w:p></w:txbxContent></wps:txbx>' +
  '<wps:bodyPr><a:noAutofit/></wps:bodyPr></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing>'

function tableXml(cellExtra: string): string {
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>cell text</w:t></w:r>' +
    `<w:r>${cellExtra}</w:r></w:p></w:tc></w:tr></w:tbl>`
  )
}

function multiParaTableXml(secondParaExtra: string): string {
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr/>' +
    '<w:p><w:r><w:t>first</w:t></w:r></w:p>' +
    `<w:p><w:r><w:t>second</w:t></w:r><w:r>${secondParaExtra}</w:r></w:p>` +
    '<w:p><w:r><w:t>third</w:t></w:r></w:p>' +
    '</w:tc></w:tr></w:tbl>'
  )
}

describe('anchored shapes inside table cells (tdf134277)', () => {
  it('extracts the shape as a cell display box and keeps the cell text', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: tableXml(TRIANGLE_ANCHOR) }))
    const cell = doc.blocks[0].table!.rows[0][0]
    expect(cell.anchoredBoxes).toHaveLength(1)
    expect(cell.anchoredBoxes![0].prst).toBe('triangle')
    expect(cell.paras[0]).toBe('cell text')
  })

  it('does not duplicate a shape paired with an mc:Fallback VML twin', async () => {
    const wrapped =
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      `<mc:Choice Requires="wps">${TRIANGLE_ANCHOR}</mc:Choice>` +
      '<mc:Fallback><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="position:absolute;width:46pt;height:42pt">' +
      '<v:textbox><w:txbxContent><w:p><w:r><w:t>box text</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
      '</v:rect></w:pict></mc:Fallback></mc:AlternateContent>'
    const doc = await parseDocx(await buildDocx({ bodyXml: tableXml(wrapped) }))
    const cell = doc.blocks[0].table!.rows[0][0]
    expect(cell.anchoredBoxes).toHaveLength(1)
    // the fallback's txbxContent text must not leak into the cell's plain text
    expect(cell.paras[0]).not.toContain('box text')
  })

  it('strips mc:Fallback twins whose tag carries attributes (Word 2024 output)', async () => {
    const wrapped =
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      `<mc:Choice Requires="wps">${TRIANGLE_ANCHOR}</mc:Choice>` +
      '<mc:Fallback xmlns:w16sdtfl="http://schemas.microsoft.com/office/word/2024/wordSdtFallback">' +
      '<w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" style="position:absolute;width:46pt;height:42pt">' +
      '<v:textbox><w:txbxContent><w:p><w:r><w:t>box text</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
      '</v:rect></w:pict></mc:Fallback></mc:AlternateContent>'
    const doc = await parseDocx(await buildDocx({ bodyXml: tableXml(wrapped) }))
    const cell = doc.blocks[0].table!.rows[0][0]
    expect(cell.anchoredBoxes).toHaveLength(1)
    expect(cell.paras[0]).not.toContain('box text')
  })

  it('records the anchor paragraph index of each box', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: multiParaTableXml(TEXTBOX_ANCHOR) }))
    const cell = doc.blocks[0].table!.rows[0][0]
    expect(cell.paras).toEqual(['first', 'second', 'third'])
    expect(cell.anchoredBoxes).toHaveLength(1)
    expect(cell.anchoredBoxAnchors).toEqual([1])
  })

  it('resolves a:sysClr fill and a:prstClr border colors', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: tableXml(TEXTBOX_ANCHOR) }))
    const box = doc.blocks[0].table!.rows[0][0].anchoredBoxes![0]
    expect(box.fill).toBe('FFFFFF')
    expect(box.borderColor).toBe('000000')
  })

  it('keeps the exact line rule of textbox paragraphs', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: tableXml(TEXTBOX_ANCHOR) }))
    const para = doc.blocks[0].table!.rows[0][0].anchoredBoxes![0].paras[0]
    expect(para.lineRule).toBe('exact')
    expect(para.lineRawTwips).toBe(240)
    expect(para.lineSpacing).toBeUndefined()
  })
})
