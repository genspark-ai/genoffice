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

function tableXml(cellExtra: string): string {
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr/><w:p><w:r><w:t>cell text</w:t></w:r>' +
    `<w:r>${cellExtra}</w:r></w:p></w:tc></w:tr></w:tbl>`
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
})
