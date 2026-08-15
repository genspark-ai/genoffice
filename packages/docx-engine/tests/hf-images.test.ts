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
