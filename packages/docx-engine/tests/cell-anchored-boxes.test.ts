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

// behind-text wrapNone group (converter output: cell background fills) anchored in a header cell
const BEHIND_GROUP_ANCHOR =
  '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice Requires="wps">' +
  '<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" allowOverlap="1" layoutInCell="1" locked="0" behindDoc="1" simplePos="0" relativeHeight="487156736">' +
  '<wp:simplePos x="0" y="0"/>' +
  '<wp:positionH relativeFrom="column"><wp:posOffset>3175</wp:posOffset></wp:positionH>' +
  '<wp:positionV relativeFrom="paragraph"><wp:posOffset>217920</wp:posOffset></wp:positionV>' +
  '<wp:extent cx="4165600" cy="1835150"/><wp:wrapNone/><wp:docPr id="77" name="Group 77"/>' +
  '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
  '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
  '<wpg:wgp xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:cNvGrpSpPr/><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4165600" cy="1835150"/><a:chExt cx="4165600" cy="1835150"/></a:xfrm></wpg:grpSpPr>' +
  '<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="749300" cy="914400"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="F1F5F9"/></a:solidFill></wps:spPr><wps:bodyPr><a:noAutofit/></wps:bodyPr></wps:wsp>' +
  '<wps:wsp><wps:spPr><a:xfrm><a:off x="2000250" y="10"/><a:ext cx="2165350" cy="304800"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="EFF5FF"/></a:solidFill></wps:spPr><wps:bodyPr><a:noAutofit/></wps:bodyPr></wps:wsp>' +
  '</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent>'

function pictureAnchor(id: number, offX: number, offY: number, ext: number): string {
  return (
    `<w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" simplePos="0" relativeHeight="${id}" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    '<wp:simplePos x="0" y="0"/>' +
    `<wp:positionH relativeFrom="column"><wp:posOffset>${offX}</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${offY}</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="${ext}" cy="${ext}"/><wp:wrapSquare wrapText="bothSides"/><wp:docPr id="${id}" name="Picture ${id}"/>` +
    '<a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="Picture ${id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId10"/></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${ext}" cy="${ext}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>` +
    '</a:graphicData></a:graphic></wp:anchor></w:drawing>'
  )
}

function pictureCellXml(drawings: string[]): string {
  return (
    '<w:tbl><w:tblPr><w:tblW w:w="9000" w:type="dxa"/></w:tblPr>' +
    '<w:tblGrid><w:gridCol w:w="9000"/></w:tblGrid>' +
    '<w:tr><w:tc><w:tcPr/><w:p><w:pPr><w:jc w:val="center"/></w:pPr>' +
    drawings.map((d) => `<w:r>${d}</w:r>`).join('') +
    '</w:p></w:tc></w:tr></w:tbl>'
  )
}

describe('flow-less and picture anchors inside table cells', () => {
  it('flags a behind-text wrapNone group so the row does not grow for it', async () => {
    const doc = await parseDocx(await buildDocx({ bodyXml: tableXml(BEHIND_GROUP_ANCHOR) }))
    const cell = doc.blocks[0].table!.rows[0][0]
    expect(cell.anchoredBoxes!.length).toBeGreaterThan(0)
    for (const box of cell.anchoredBoxes!) {
      expect(box.noWrap).toBe(true)
      expect(box.behind).toBe(true)
    }
    expect(cell.paras[0]).toBe('cell text')
  })

  it('places two pictures anchored in one cell paragraph as boxes at their own offsets', async () => {
    const xml = pictureCellXml([
      pictureAnchor(1, -65405, 140335, 1082040),
      pictureAnchor(2, 918845, 178435, 1057275),
    ])
    const doc = await parseDocx(await buildDocx({ bodyXml: xml, withImage: true }))
    const cell = doc.blocks[0].table!.rows[0][0]
    expect(cell.anchoredBoxes).toHaveLength(2)
    const [a, b] = cell.anchoredBoxes!
    expect(a.fillImageDataUrl).toMatch(/^data:image\/png/)
    expect([a.offsetXEmu, b.offsetXEmu]).toEqual([-65405, 918845])
    expect([a.offsetYEmu, b.offsetYEmu]).toEqual([140335, 178435])
    expect(a.floating).toBe(true)
    expect(a.noWrap).toBeUndefined()
    expect(b.heightPx).toBe(111)
    // the pictures no longer ride the cell runs (they would float below each other)
    expect(cell.richParas![0].runs.some((r) => r.image)).toBe(false)
  })

  it('keeps a lone anchored picture on the run-image path', async () => {
    const xml = pictureCellXml([pictureAnchor(1, -65405, 140335, 1082040)])
    const doc = await parseDocx(await buildDocx({ bodyXml: xml, withImage: true }))
    const cell = doc.blocks[0].table!.rows[0][0]
    expect(cell.anchoredBoxes).toBeUndefined()
    expect(cell.richParas![0].runs.some((r) => r.image)).toBe(true)
  })
})
