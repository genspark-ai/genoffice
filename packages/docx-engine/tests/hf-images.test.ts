import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const HEADER_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<w:p><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="381000" cy="190500"/><wp:docPr id="1" name="Logo"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>' +
  '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:t>Confidential</w:t></w:r></w:p>' +
  '</w:hdr>'

const HEADER_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
  '</Relationships>'

async function buildHeaderLogoDocx(headerXml: string = HEADER_XML): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
    withImage: true,
    extraRels:
      '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
    sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
    extraParts: [
      {
        path: 'word/header1.xml',
        xml: headerXml,
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

describe('header/footer images (display-only Logo)', () => {
  it('parses header images with size, text paragraphs unaffected', async () => {
    const doc = await parseDocx(await buildHeaderLogoDocx())
    expect(doc.headerImages).toHaveLength(1)
    const img = doc.headerImages![0]
    expect(img.dataUrl.startsWith('data:image/png;base64,')).toBe(true)
    expect(img.widthPx).toBe(40)
    expect(img.heightPx).toBe(20)
    expect(img.floating).toBeUndefined()
    // text paragraphs enter the model as usual; image paragraphs produce no empty text paragraph
    expect(doc.headerText).toBe('Confidential')
    // hfParts (multi-section path) carries images too
    const part = Object.values(doc.hfParts ?? {}).find((p) => p.text === 'Confidential')
    expect(part?.images).toHaveLength(1)
  })

  it('untouched round-trip stays byte-identical', async () => {
    const bytes = await buildHeaderLogoDocx()
    const doc = await parseDocx(bytes)
    const blocks = doc.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, blocks)).toEqual(bytes)
  })

  it('inline image follows its paragraph alignment (POI headerPic: w:jc right)', async () => {
    const headerXml = HEADER_XML.replace(
      '<w:p><w:r><w:drawing>',
      '<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:r><w:drawing>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    expect(doc.headerImages![0].align).toBe('right')
  })

  it('anchored image reads behindDoc and page-relative posOffsets (picture watermark)', async () => {
    const headerXml = HEADER_XML.replace(
      '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
        '<wp:extent cx="381000" cy="190500"/>',
      '<wp:anchor behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">' +
        '<wp:positionH relativeFrom="page"><wp:posOffset>190500</wp:posOffset></wp:positionH>' +
        '<wp:positionV relativeFrom="page"><wp:posOffset>-95250</wp:posOffset></wp:positionV>' +
        '<wp:extent cx="381000" cy="190500"/>',
    ).replace('</wp:inline>', '</wp:anchor>')
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    expect(doc.headerImages).toHaveLength(1)
    const img = doc.headerImages![0]
    expect(img.floating).toBe(true)
    expect(img.behind).toBe(true)
    expect(img.posXPx).toBe(20)
    expect(img.posYPx).toBe(-10)
    expect(img.posHRel).toBe('page')
    expect(img.posVRel).toBe('page')
    expect(img.posH).toBeUndefined()
    expect(img.posV).toBeUndefined()
  })

  it('anchored image maps wp:align to the alignment fields (margin-relative)', async () => {
    const headerXml = HEADER_XML.replace(
      '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
        '<wp:extent cx="381000" cy="190500"/>',
      '<wp:anchor behindDoc="0">' +
        '<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH>' +
        '<wp:positionV relativeFrom="margin"><wp:align>bottom</wp:align></wp:positionV>' +
        '<wp:extent cx="381000" cy="190500"/>',
    ).replace('</wp:inline>', '</wp:anchor>')
    const img = (await parseDocx(await buildHeaderLogoDocx(headerXml))).headerImages![0]
    expect(img.floating).toBe(true)
    expect(img.behind).toBeUndefined()
    expect(img.posH).toBe('center')
    expect(img.posV).toBe('bottom')
    expect(img.posXPx).toBeUndefined()
    expect(img.posYPx).toBeUndefined()
  })

  it('AlternateContent picks the first blip whose media resolves (mac PDF Choice → PNG Fallback)', async () => {
    // rId9 is unresolvable (missing media part); the PNG fallback must be used
    const headerXml = HEADER_XML.replace(
      '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic>',
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
        '<mc:Choice Requires="ma"><pic:pic><pic:blipFill><a:blip r:embed="rId9"/></pic:blipFill></pic:pic></mc:Choice>' +
        '<mc:Fallback><pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic></mc:Fallback>' +
        '</mc:AlternateContent>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    expect(doc.headerImages).toHaveLength(1)
    expect(doc.headerImages![0].dataUrl.startsWith('data:image/png;base64,')).toBe(true)
  })
})

describe('layout-table cell images (header logo in a w:tbl cell)', () => {
  const INLINE_LOGO =
    '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
    '<wp:extent cx="381000" cy="190500"/><wp:docPr id="1" name="Logo"/>' +
    '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic>' +
    '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'

  const tableHeaderXml = (cell1Extra = '') =>
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
    ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
    ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
    ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
    '<w:tbl><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="6000"/></w:tblGrid><w:tr>' +
    `<w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p>${INLINE_LOGO}${cell1Extra}</w:p></w:tc>` +
    '<w:tc><w:tcPr><w:tcW w:w="6000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Title</w:t></w:r></w:p></w:tc>' +
    '</w:tr></w:tbl></w:hdr>'

  it('inline cell picture becomes a cell run image, not a part-level strip image', async () => {
    const doc = await parseDocx(await buildHeaderLogoDocx(tableHeaderXml()))
    expect(doc.headerImages ?? []).toHaveLength(0)
    const row = doc.headerParas!.find((p) => p.cells)
    expect(row).toBeDefined()
    const img = row!.cells![0].paras[0][0].image
    expect(img?.dataUrl.startsWith('data:image/png;base64,')).toBe(true)
    expect(img?.widthPx).toBe(40)
    expect(img?.heightPx).toBe(20)
    // an image-only row still enters the model
    expect(row!.cells![1].paras[0][0].text).toBe('Title')
  })

  it('anchored picture in a cell stays part-level (floating), not a cell run', async () => {
    const anchored = INLINE_LOGO.replace(
      '<wp:inline distT="0" distB="0" distL="0" distR="0">',
      '<wp:anchor behindDoc="1">' +
        '<wp:positionH relativeFrom="page"><wp:posOffset>190500</wp:posOffset></wp:positionH>' +
        '<wp:positionV relativeFrom="page"><wp:posOffset>190500</wp:posOffset></wp:positionV>',
    ).replace('</wp:inline>', '</wp:anchor>')
    const doc = await parseDocx(await buildHeaderLogoDocx(tableHeaderXml(anchored)))
    expect(doc.headerImages).toHaveLength(1)
    expect(doc.headerImages![0].floating).toBe(true)
    const row = doc.headerParas!.find((p) => p.cells)!
    // only the inline logo remains a cell run
    const cellImages = row.cells![0].paras.flat().filter((r) => r.image)
    expect(cellImages).toHaveLength(1)
    expect(cellImages[0].image!.xml.includes('<wp:inline')).toBe(true)
  })

  it('inline picture in a nested table rides its cell run (not duplicated part-level)', async () => {
    const headerXml = tableHeaderXml().replace(
      `<w:p>${INLINE_LOGO}</w:p>`,
      '<w:tbl><w:tblGrid><w:gridCol w:w="1500"/></w:tblGrid><w:tr>' +
        `<w:tc><w:tcPr><w:tcW w:w="1500" w:type="dxa"/></w:tcPr><w:p>${INLINE_LOGO}</w:p></w:tc>` +
        '</w:tr></w:tbl><w:p/>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    expect(doc.headerImages ?? []).toHaveLength(0)
    const row = doc.headerParas!.find((p) => p.cells)!
    const cellImages = row.cells![0].paras.flat().filter((r) => r.image)
    expect(cellImages).toHaveLength(1)
    expect(cellImages[0].image!.xml.includes('<wp:inline')).toBe(true)
  })

  it('run with both text and an anchored drawing keeps its text as a cell run', async () => {
    const anchored = INLINE_LOGO.replace('<w:r><w:drawing>', '<w:r><w:t>Ref</w:t><w:drawing>')
      .replace(
        '<wp:inline distT="0" distB="0" distL="0" distR="0">',
        '<wp:anchor behindDoc="0">' +
          '<wp:positionH relativeFrom="page"><wp:posOffset>190500</wp:posOffset></wp:positionH>' +
          '<wp:positionV relativeFrom="page"><wp:posOffset>190500</wp:posOffset></wp:positionV>',
      )
      .replace('</wp:inline>', '</wp:anchor>')
    const doc = await parseDocx(
      await buildHeaderLogoDocx(tableHeaderXml().replace(INLINE_LOGO, anchored)),
    )
    expect(doc.headerImages).toHaveLength(1)
    const row = doc.headerParas!.find((p) => p.cells)!
    const runs = row.cells![0].paras.flat()
    expect(runs.map((r) => r.text).join('')).toBe('Ref')
    expect(runs.some((r) => r.image)).toBe(false)
  })

  it('untouched round-trip of a table header stays byte-identical', async () => {
    const bytes = await buildHeaderLogoDocx(tableHeaderXml())
    const doc = await parseDocx(bytes)
    const blocks = doc.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, blocks)).toEqual(bytes)
  })
})
