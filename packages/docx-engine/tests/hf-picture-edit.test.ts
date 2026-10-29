/**
 * A floating header picture moved or resized in the strip: the save rewrites
 * only that drawing's position (pinned to the page) and extents, and a
 * picture-only edit keeps the part's text paragraphs byte for byte.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx, patchHfPicturesXml, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const anchor = (id: number, x: number, y: number) =>
  '<w:drawing><wp:anchor allowOverlap="1" behindDoc="0" distB="0" distT="0" distL="0" distR="0"' +
  ' layoutInCell="1" locked="0" relativeHeight="0" simplePos="0"><wp:simplePos x="0" y="0"/>' +
  `<wp:positionH relativeFrom="column"><wp:posOffset>${x}</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>${y}</wp:posOffset></wp:positionV>` +
  '<wp:extent cx="952500" cy="476250"/><wp:wrapSquare wrapText="bothSides"/>' +
  `<wp:docPr id="${id}" name="Banner"/>` +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="476250"/></a:xfrm></pic:spPr>' +
  '</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>'

const TEXT_PARA =
  '<w:p w:rsidR="00AB12CD"><w:pPr><w:jc w:val="right"/></w:pPr>' +
  '<w:r><w:rPr><w:b/></w:rPr><w:t>Revista</w:t></w:r></w:p>'

const HEADER_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  `<w:p><w:r>${anchor(5, 19051, 19051)}</w:r></w:p>` +
  TEXT_PARA +
  `<w:p><w:r>${anchor(6, 0, 952500)}</w:r></w:p>` +
  '</w:hdr>'

const HEADER_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
  '</Relationships>'

const buildBannerDocx = () =>
  buildDocx({
    bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
    withImage: true,
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

const savedHeader = async (bytes: Uint8Array) =>
  (await JSZip.loadAsync(bytes)).file('word/header1.xml')!.async('string')

describe('moving and resizing a floating header picture', () => {
  it('tags each single-picture anchor with its drawing ordinal', async () => {
    const doc = await parseDocx(await buildBannerDocx())
    expect(doc.headerImages?.map((im) => im.drawing)).toEqual([0, 1])
  })

  it('pins the edited drawing to the page and resizes it, leaving the other one alone', () => {
    const out = patchHfPicturesXml(HEADER_XML, {
      1: { posXPx: 96, posYPx: 200, widthPx: 150, heightPx: 75 },
    })
    expect(out).toContain(anchor(5, 19051, 19051))
    expect(out).toContain(
      '<wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>',
    )
    expect(out).toContain(
      '<wp:positionV relativeFrom="page"><wp:posOffset>1905000</wp:posOffset></wp:positionV>',
    )
    expect(out).toContain('<wp:extent cx="1428750" cy="714375"/>')
    expect(out).toContain('<a:ext cx="1428750" cy="714375"/>')
    expect(out.replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, '')).toBe(
      HEADER_XML.replace(/<w:drawing>[\s\S]*?<\/w:drawing>/g, ''),
    )
  })

  it('a picture-only save keeps the text paragraphs byte for byte', async () => {
    const parsed = await parseDocx(await buildBannerDocx())
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      header: {
        text: parsed.headerText ?? '',
        paras: parsed.headerParas ?? undefined,
        pictureOnly: true,
        pictureEdits: { 0: { posXPx: 40, posYPx: 30, widthPx: 100, heightPx: 50 } },
      },
    })
    const hdr = await savedHeader(saved)
    expect(hdr).toContain(TEXT_PARA)
    expect(hdr).toContain(anchor(6, 0, 952500))
    expect(hdr).toContain('<wp:posOffset>381000</wp:posOffset></wp:positionH>')
    const reparsed = await parseDocx(saved)
    expect(reparsed.headerImages?.[0]).toMatchObject({
      posXPx: 40,
      posYPx: 30,
      posHRel: 'page',
      posVRel: 'page',
      widthPx: 100,
      heightPx: 50,
    })
  })

  it('a text edit made alongside still moves the picture', async () => {
    const parsed = await parseDocx(await buildBannerDocx())
    const paras = parsed.headerParas!.map((p) =>
      p.runs.some((r) => r.text === 'Revista')
        ? { ...p, runs: p.runs.map((r) => ({ ...r, text: 'Journal' })) }
        : p,
    )
    const saved = await saveDocx(parsed, [{ kind: 'original', docxIndex: 0 }], {
      header: {
        text: 'Journal',
        paras,
        pictureEdits: { 1: { posXPx: 10, posYPx: 20, widthPx: 100, heightPx: 50 } },
      },
    })
    const hdr = await savedHeader(saved)
    expect(hdr).toContain('Journal')
    expect(hdr).not.toContain('Revista')
    expect(hdr).toContain(anchor(5, 19051, 19051))
    expect(hdr).toContain('<wp:posOffset>95250</wp:posOffset></wp:positionH>')
  })
})
