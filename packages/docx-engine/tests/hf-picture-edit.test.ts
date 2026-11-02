/**
 * A floating header picture moved or resized in the strip: the save rewrites
 * only that drawing's extents and the axes it moved along (pinned to the
 * page), and a picture-only edit keeps the part's text paragraphs byte for byte.
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

const WPS_NS = 'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"'

const buildBannerDocx = (headerXml = HEADER_XML) =>
  buildDocx({
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

  it('a resize from the far edge keeps the position base and offset', () => {
    const out = patchHfPicturesXml(HEADER_XML, { 0: { widthPx: 102, heightPx: 51 } })
    const [first] = out.match(/<w:drawing>[\s\S]*?<\/w:drawing>/g)!
    expect(first).toContain(
      '<wp:positionH relativeFrom="column"><wp:posOffset>19051</wp:posOffset></wp:positionH>',
    )
    expect(first).toContain(
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>19051</wp:posOffset></wp:positionV>',
    )
    expect(first).toContain('<wp:extent cx="971550" cy="485775"/>')
  })

  it('a move along one axis pins only that axis to the page', () => {
    const out = patchHfPicturesXml(HEADER_XML, { 0: { posYPx: 60, widthPx: 100, heightPx: 50 } })
    const [first] = out.match(/<w:drawing>[\s\S]*?<\/w:drawing>/g)!
    expect(first).toContain(
      '<wp:positionH relativeFrom="column"><wp:posOffset>19051</wp:posOffset></wp:positionH>',
    )
    expect(first).toContain(
      '<wp:positionV relativeFrom="page"><wp:posOffset>571500</wp:posOffset></wp:positionV>',
    )
  })

  it('an anchor without position elements gets them where the schema orders them', () => {
    const bare = anchor(7, 0, 0).replace(/<wp:position([HV])\b[\s\S]*?<\/wp:position\1>/g, '')
    const out = patchHfPicturesXml(bare, {
      0: { posXPx: 10, posYPx: 20, widthPx: 100, heightPx: 50 },
    })
    expect(out).toContain(
      '<wp:simplePos x="0" y="0"/>' +
        '<wp:positionH relativeFrom="page"><wp:posOffset>95250</wp:posOffset></wp:positionH>' +
        '<wp:positionV relativeFrom="page"><wp:posOffset>190500</wp:posOffset></wp:positionV>' +
        '<wp:extent ',
    )
  })

  it('a text box holding a picture is neither tagged nor patched', async () => {
    // the lazy <w:drawing> match ends at the nested picture's </w:drawing>,
    // where the fragment holds one pic:pic with the text box's extents
    const textBox =
      '<w:drawing><wp:anchor simplePos="0" relativeHeight="0" behindDoc="0" locked="0"' +
      ' layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="2857500" cy="952500"/><wp:wrapNone/><wp:docPr id="9" name="Box"/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<wps:wsp><wps:txbx><w:txbxContent><w:p><w:r>' +
      anchor(10, 0, 0) +
      '</w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic>' +
      '</wp:anchor></w:drawing>'
    const xml = HEADER_XML.replace(TEXT_PARA, `<w:p><w:r>${textBox}</w:r></w:p>`)
    expect(
      patchHfPicturesXml(xml, { 1: { posXPx: 5, posYPx: 5, widthPx: 20, heightPx: 20 } }),
    ).toBe(xml)
    const doc = await parseDocx(
      await buildBannerDocx(xml.replace('xmlns:pic=', `${WPS_NS} xmlns:pic=`)),
    )
    expect(doc.headerImages?.map((im) => im.drawing).filter((d) => d != null)).toEqual([0, 2])
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
