import JSZip from 'jszip'
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

async function buildHeaderLogoDocx(
  headerXml: string = HEADER_XML,
  headerRels: string = HEADER_RELS,
): Promise<Uint8Array> {
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
        xml: headerRels,
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

  it('resolves absolute percent-encoded image targets', async () => {
    const zip = await JSZip.loadAsync(await buildHeaderLogoDocx())
    const image = await zip.file('word/media/image1.png')!.async('uint8array')
    zip.remove('word/media/image1.png')
    zip.file('word/media/image 1.png', image)
    const rels = HEADER_RELS.replace('media/image1.png', '/word/media/image%201.png')
    zip.file('word/_rels/header1.xml.rels', rels)
    const bytes = await zip.generateAsync({ type: 'uint8array' })
    const doc = await parseDocx(bytes)
    expect(doc.headerImages).toHaveLength(1)
    expect(doc.headerImages![0].dataUrl.startsWith('data:image/png;base64,')).toBe(true)
    expect(rels).toContain('image%201.png')
  })

  it('untouched round-trip stays byte-identical', async () => {
    const bytes = await buildHeaderLogoDocx()
    const doc = await parseDocx(bytes)
    const blocks = doc.blocks
      .filter((b) => !b.hidden && b.docxIndex !== null)
      .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! }))
    expect(await saveDocx(doc, blocks)).toEqual(bytes)
  })

  it('parses a:srcRect crop fractions (two same-image crops read as one picture each)', async () => {
    const headerXml = HEADER_XML.replace(
      '<a:blip r:embed="rId1"/>',
      '<a:blip r:embed="rId1"/><a:srcRect t="69600" r="77114"/>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    expect(doc.headerImages![0].crop).toEqual({ l: 0, t: 0.696, r: 0.77114, b: 0 })
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

  it('anchored image reads the wrap mode and keeps paragraph-relative positionV (prod_004 shape)', async () => {
    const headerXml = HEADER_XML.replace(
      '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
        '<wp:extent cx="381000" cy="190500"/>',
      '<wp:anchor behindDoc="0">' +
        '<wp:positionH relativeFrom="column"><wp:posOffset>190500</wp:posOffset></wp:positionH>' +
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>-137208</wp:posOffset></wp:positionV>' +
        '<wp:extent cx="381000" cy="958215"/>' +
        '<wp:wrapSquare wrapText="bothSides"/>',
    ).replace('</wp:inline>', '</wp:anchor>')
    const img = (await parseDocx(await buildHeaderLogoDocx(headerXml))).headerImages![0]
    expect(img.floating).toBe(true)
    expect(img.wrap).toBe('square')
    expect(img.posVRel).toBe('paragraph')
    expect(img.posYPx).toBe(-14)
    expect(img.heightPx).toBe(100.6)
    expect(img.posHRel).toBe('margin')
  })

  it('wrapTopAndBottom and wrapNone map to topBottom / none', async () => {
    const withWrap = (wrapXml: string) =>
      HEADER_XML.replace(
        '<wp:inline distT="0" distB="0" distL="0" distR="0">' +
          '<wp:extent cx="381000" cy="190500"/>',
        '<wp:anchor behindDoc="0">' +
          '<wp:positionV relativeFrom="page"><wp:posOffset>335280</wp:posOffset></wp:positionV>' +
          '<wp:extent cx="381000" cy="377190"/>' +
          wrapXml,
      ).replace('</wp:inline>', '</wp:anchor>')
    const tb = (await parseDocx(await buildHeaderLogoDocx(withWrap('<wp:wrapTopAndBottom/>'))))
      .headerImages![0]
    expect(tb.wrap).toBe('topBottom')
    expect(tb.posVRel).toBe('page')
    const none = (await parseDocx(await buildHeaderLogoDocx(withWrap('<wp:wrapNone/>'))))
      .headerImages![0]
    expect(none.wrap).toBe('none')
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

  it('AlternateContent Fallback re-emitting the Choice picture parses once (prod_043)', async () => {
    // anchored drawing in the Choice, inline copy of the same logo in the Fallback;
    // the AC open tag carries attributes (the engine's own writer inlines xmlns:mc)
    const drawingWith = (inner: string) =>
      `<w:drawing>${inner}<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
      '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic>' +
      `</a:graphicData></a:graphic>${inner.startsWith('<wp:anchor') ? '</wp:anchor>' : '</wp:inline>'}</w:drawing>`
    const anchor =
      '<wp:anchor behindDoc="0" distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" locked="0" layoutInCell="1" allowOverlap="1">' +
      '<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="margin"><wp:posOffset>0</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="381000" cy="190500"/><wp:wrapNone/><wp:docPr id="2" name="Logo"/>'
    const inline =
      '<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="381000" cy="190500"/><wp:docPr id="3" name="Logo"/>'
    const headerXml = HEADER_XML.replace(
      /<w:p><w:r><w:drawing>[\s\S]*?<\/w:drawing><\/w:r><\/w:p>/,
      '<w:p><w:r>' +
        '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
        `<mc:Choice Requires="wps">${drawingWith(anchor)}</mc:Choice>` +
        `<mc:Fallback>${drawingWith(inline)}</mc:Fallback>` +
        '</mc:AlternateContent>' +
        '</w:r></w:p>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    expect(doc.headerImages).toHaveLength(1)

    // when the Choice yields nothing (unresolvable media), the Fallback still counts
    const brokenChoice = headerXml.replace(
      /<mc:Choice Requires="wps">[\s\S]*?<\/mc:Choice>/,
      `<mc:Choice Requires="wps">${drawingWith(anchor).replace('rId1', 'rId9')}</mc:Choice>`,
    )
    const doc2 = await parseDocx(await buildHeaderLogoDocx(brokenChoice))
    expect(doc2.headerImages).toHaveLength(1)
  })

  it('VML picture watermark parses when w:pict carries attributes (prod_091)', async () => {
    const headerXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
      ' xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"' +
      ' xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">' +
      '<w:p><w:r><w:pict w14:anchorId="7C49EC32">' +
      '<v:shape id="WordPictureWatermark1" type="#_x0000_t75" style="position:absolute;' +
      'margin-left:0;margin-top:0;width:480pt;height:360pt;z-index:-251655168;' +
      'mso-position-horizontal:center;mso-position-vertical:center">' +
      '<v:imagedata r:id="rId1" o:title="logo"/></v:shape>' +
      '</w:pict></w:r></w:p></w:hdr>'
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    expect(doc.headerImages).toHaveLength(1)
    const img = doc.headerImages![0]
    expect(img.dataUrl.startsWith('data:image/png;base64,')).toBe(true)
    expect(img.floating).toBe(true)
    expect(img.behind).toBe(true)
    expect(img.posH).toBe('center')
    expect(img.posV).toBe('center')
    expect(img.widthPx).toBe(640)
    expect(img.heightPx).toBe(480)
  })

  it('textless custGeom shape group renders as one SVG float image (prod_044 ornaments)', async () => {
    const headerXml =
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
      ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
      ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
      ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
      ' xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"' +
      ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
      '<w:p><w:r><w:drawing><wp:anchor behindDoc="1">' +
      '<wp:simplePos x="0" y="0"/>' +
      '<wp:positionH relativeFrom="page"><wp:posOffset>952500</wp:posOffset></wp:positionH>' +
      '<wp:positionV relativeFrom="paragraph"><wp:posOffset>-190500</wp:posOffset></wp:positionV>' +
      '<wp:extent cx="190500" cy="190500"/><wp:wrapNone/>' +
      '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">' +
      '<wpg:wgp><wpg:cNvGrpSpPr/><wpg:grpSpPr bwMode="auto"><a:xfrm>' +
      '<a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="381000" cy="381000"/>' +
      '</a:xfrm></wpg:grpSpPr>' +
      '<wps:wsp><wps:cNvPr id="2" name="Graphic 1"/><wps:cNvSpPr/><wps:spPr>' +
      '<a:xfrm><a:off x="0" y="0"/><a:ext cx="381000" cy="381000"/></a:xfrm>' +
      '<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/>' +
      '<a:pathLst><a:path w="381000" h="381000">' +
      '<a:moveTo><a:pt x="0" y="0"/></a:moveTo>' +
      '<a:lnTo><a:pt x="381000" y="0"/></a:lnTo>' +
      '<a:lnTo><a:pt x="190500" y="381000"/></a:lnTo>' +
      '<a:close/></a:path></a:pathLst></a:custGeom>' +
      '<a:solidFill><a:srgbClr val="20093F"/></a:solidFill>' +
      '</wps:spPr><wps:bodyPr/></wps:wsp>' +
      '</wpg:wgp></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:hdr>'
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    expect(doc.headerImages).toHaveLength(1)
    const img = doc.headerImages![0]
    expect(img.floating).toBe(true)
    expect(img.behind).toBe(true)
    expect(img.posXPx).toBe(100)
    expect(img.posYPx).toBe(-20)
    expect(img.posVRel).toBe('paragraph')
    expect(img.widthPx).toBe(20)
    expect(img.heightPx).toBe(20)
    expect(img.dataUrl.startsWith('data:image/svg+xml,')).toBe(true)
    const svg = decodeURIComponent(img.dataUrl.slice('data:image/svg+xml,'.length))
    expect(svg).toContain('viewBox="0 0 20 20"')
    expect(svg).toContain('fill="#20093F"')
    expect(svg).toContain('d="M 0 0 L 20 0 L 10 20 Z"')
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

describe('header/footer VML shapes and watermarks', () => {
  const VML_NS =
    ' xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"' +
    ' xmlns:w10="urn:schemas-microsoft-com:office:word"'
  const hdr = (body: string, extraNs = ''): string =>
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
    ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
    VML_NS +
    extraNs +
    `>${body}</w:hdr>`

  it('text watermark keeps its box, rotation, color, opacity, font and margin-centered anchor', async () => {
    const headerXml = hdr(
      '<w:p><w:r><w:pict>' +
        '<v:shapetype id="_x0000_t136" coordsize="21600,21600" o:spt="136" path="m@7,l@8,m@5,21600l@6,21600e">' +
        '<v:textpath on="t" fitshape="t"/></v:shapetype>' +
        '<v:shape id="PowerPlusWaterMarkObject1" type="#_x0000_t136" style="position:absolute;margin-left:0;' +
        'margin-top:0;width:4in;height:2in;rotation:315;z-index:-251658752;mso-position-horizontal:center;' +
        'mso-position-horizontal-relative:margin;mso-position-vertical:center;mso-position-vertical-relative:margin"' +
        ' fillcolor="red" stroked="f"><v:fill opacity=".5"/>' +
        '<v:textpath style="font-family:&quot;SimSun-ExtB&quot;;font-size:2in" string="SECRET"/>' +
        '</v:shape></w:pict></w:r></w:p>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    expect(doc.watermarkText).toBe('SECRET')
    expect(doc.headerImages).toHaveLength(1)
    const wm = doc.headerImages![0]
    expect(wm.dataUrl).toBe('')
    expect(wm.floating).toBe(true)
    expect(wm.behind).toBe(true)
    expect(wm.widthPx).toBe(384)
    expect(wm.heightPx).toBe(192)
    expect(wm.rotationDeg).toBe(315)
    expect(wm.posH).toBe('center')
    expect(wm.posV).toBe('center')
    expect(wm.posHRel).toBe('margin')
    expect(wm.posVRel).toBe('margin')
    expect(wm.wordArt).toEqual({
      text: 'SECRET',
      colorHex: 'FF0000',
      opacity: 0.5,
      fontFamily: 'SimSun-ExtB',
    })
  })

  it('horizontal watermark has no rotation; silver fill and 1pt declared font do not shrink the box', async () => {
    const headerXml = hdr(
      '<w:p><w:r><w:pict><v:shape id="PowerPlusWaterMarkObject2" type="#_x0000_t136"' +
        ' style="position:absolute;margin-left:0;margin-top:0;width:468pt;height:280.8pt;z-index:-251658752;' +
        'mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;' +
        'mso-position-vertical-relative:margin" fillcolor="silver" stroked="f"><v:fill opacity=".5"/>' +
        '<v:textpath style="font-family:&quot;Calibri&quot;;font-size:1pt" string="DRAFT"/>' +
        '</v:shape></w:pict></w:r></w:p>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    const wm = doc.headerImages![0]
    expect(wm.rotationDeg).toBeUndefined()
    expect(wm.widthPx).toBe(624)
    expect(wm.heightPx).toBe(374)
    expect(wm.wordArt?.colorHex).toBe('C0C0C0')
    expect(wm.wordArt?.fontFamily).toBe('Calibri')
  })

  it('picture watermark reads inch-sized boxes and the gain/blacklevel washout levels', async () => {
    const headerXml = hdr(
      '<w:p><w:r><w:pict><v:shape id="WordPictureWatermark1" type="#_x0000_t75"' +
        ' style="position:absolute;margin-left:0;margin-top:0;width:16in;height:12in;z-index:-251658240;' +
        'mso-position-horizontal:center;mso-position-horizontal-relative:margin;mso-position-vertical:center;' +
        'mso-position-vertical-relative:margin"><v:imagedata r:id="rId1" o:title="Flowers" gain="19661f" blacklevel="22938f"/>' +
        '</v:shape></w:pict></w:r></w:p>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    const img = doc.headerImages![0]
    expect(img.widthPx).toBe(1536)
    expect(img.heightPx).toBe(1152)
    expect(img.posHRel).toBe('margin')
    expect(img.washout?.gain).toBeCloseTo(0.3, 3)
    expect(img.washout?.blackLevel).toBeCloseTo(0.35, 3)
  })

  it('AlternateContent: a wps preset shape the Choice cannot draw falls back to its VML twin as an SVG float', async () => {
    const headerXml = hdr(
      '<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing>' +
        '<wp:anchor behindDoc="0"><wp:positionH relativeFrom="column"><wp:posOffset>584200</wp:posOffset></wp:positionH>' +
        '<wp:positionV relativeFrom="paragraph"><wp:posOffset>127000</wp:posOffset></wp:positionV>' +
        '<wp:extent cx="584200" cy="374650"/><wp:wrapNone/>' +
        '<a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
        '<wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="584200" cy="374650"/></a:xfrm>' +
        '<a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom></wps:spPr>' +
        '<wps:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef></wps:style></wps:wsp>' +
        '</a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice>' +
        '<mc:Fallback><w:pict><v:oval id="Oval 1" style="position:absolute;margin-left:46pt;margin-top:10pt;' +
        'width:46pt;height:29.5pt;z-index:251659264;mso-position-horizontal:absolute;' +
        'mso-position-horizontal-relative:text;mso-position-vertical:absolute;mso-position-vertical-relative:text"' +
        ' fillcolor="#4f81bd [3204]" strokecolor="#243f60 [1604]" strokeweight="2pt"/></w:pict></mc:Fallback>' +
        '</mc:AlternateContent></w:r></w:p>',
      ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
        ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
        ' xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    expect(doc.headerImages).toHaveLength(1)
    const oval = doc.headerImages![0]
    expect(oval.floating).toBe(true)
    expect(oval.behind).toBeUndefined()
    expect(oval.widthPx).toBe(61)
    expect(oval.heightPx).toBe(39)
    expect(oval.posXPx).toBe(61)
    expect(oval.posYPx).toBe(13)
    expect(oval.posHRel).toBe('margin')
    expect(oval.posVRel).toBe('paragraph')
    const svg = decodeURIComponent(oval.dataUrl.replace('data:image/svg+xml,', ''))
    expect(svg).toContain('<ellipse')
    expect(svg).toContain('fill="#4f81bd"')
    expect(svg).toContain('stroke="#243f60"')
    expect(svg).toContain('stroke-width="2.6666666666666665"')
  })

  it('VML groups, gradients and text-bearing shapes stay undrawn rather than partially drawn', async () => {
    const headerXml = hdr(
      '<w:p><w:r><w:pict><v:group style="position:absolute;width:35pt;height:23pt;rotation:90" coordsize="1566,590">' +
        '<v:oval style="position:absolute;left:0;top:0;width:682;height:590" fillcolor="#4f81bd">' +
        '<v:fill color2="#243f60" type="gradient"/></v:oval></v:group></w:pict></w:r></w:p>' +
        '<w:p><w:r><w:pict><v:rect style="position:absolute;width:40pt;height:485pt" filled="f" stroked="f">' +
        '<v:textbox><w:txbxContent><w:p><w:r><w:t>October</w:t></w:r></w:p></w:txbxContent></v:textbox>' +
        '</v:rect></w:pict></w:r></w:p>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(headerXml))
    expect(doc.headerImages ?? []).toHaveLength(0)
  })
})

const ANCHOR_PIC = (posH: string, posV: string): string =>
  '<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="1" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">' +
  '<wp:simplePos x="0" y="0"/>' +
  `<wp:positionH relativeFrom="${posH.slice(0, posH.indexOf(':'))}">${posH.slice(posH.indexOf(':') + 1)}</wp:positionH>` +
  `<wp:positionV relativeFrom="${posV.slice(0, posV.indexOf(':'))}">${posV.slice(posV.indexOf(':') + 1)}</wp:positionV>` +
  '<wp:extent cx="381000" cy="190500"/><wp:wrapNone/><wp:docPr id="1" name="Logo"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>'

const hdr = (inner: string): string => HEADER_XML.replace(/<w:p>[\s\S]*<\/w:p>/, inner)

describe('anchored header/footer object bands (wp:positionH/V relativeFrom)', () => {
  it('keeps the margin bands as the offset/alignment origin', async () => {
    const xml = hdr(
      '<w:p>' +
        ANCHOR_PIC(
          'rightMargin:<wp:posOffset>95250</wp:posOffset>',
          'bottomMargin:<wp:align>center</wp:align>',
        ) +
        '</w:p>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(xml))
    const img = doc.headerImages![0]
    expect(img.floating).toBe(true)
    expect(img.posHRel).toBe('rightMargin')
    expect(img.posXPx).toBe(10)
    expect(img.posV).toBe('center')
    expect(img.posVRel).toBe('bottomMargin')
  })

  it('maps inside/outside margins onto the left/top and right/bottom bands', async () => {
    const xml = hdr(
      '<w:p>' +
        ANCHOR_PIC(
          'insideMargin:<wp:align>center</wp:align>',
          'outsideMargin:<wp:posOffset>0</wp:posOffset>',
        ) +
        '</w:p>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(xml))
    expect(doc.headerImages![0].posHRel).toBe('leftMargin')
    expect(doc.headerImages![0].posVRel).toBe('bottomMargin')
  })

  it('measures a column offset inside a layout-table cell from that cell', async () => {
    const xml = hdr(
      '<w:tbl><w:tblPr><w:tblInd w:w="150" w:type="dxa"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="7860"/><w:gridCol w:w="2940"/></w:tblGrid>' +
        '<w:tr><w:tc><w:p><w:r><w:t>Title</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:p>' +
        ANCHOR_PIC(
          'column:<wp:posOffset>781050</wp:posOffset>',
          'paragraph:<wp:posOffset>0</wp:posOffset>',
        ) +
        '</w:p></w:tc></w:tr></w:tbl>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(xml))
    const img = doc.headerImages!.find((i) => i.floating)!
    // 781050 EMU = 82 px, plus tblInd 150 + gridCol 7860 twips = 534 px
    expect(img.posHRel).toBe('margin')
    expect(img.posXPx).toBe(616)
  })

  it('ignores a nested table in an earlier cell when resolving the cell column', async () => {
    const nested =
      '<w:tbl><w:tblPr><w:tblInd w:w="9999" w:type="dxa"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="500"/><w:gridCol w:w="500"/></w:tblGrid>' +
      '<w:tr><w:tc><w:p/></w:tc><w:tc><w:p/></w:tc></w:tr></w:tbl>'
    const xml = hdr(
      '<w:tbl><w:tblPr><w:tblInd w:w="150" w:type="dxa"/></w:tblPr>' +
        '<w:tblGrid><w:gridCol w:w="7860"/><w:gridCol w:w="2940"/></w:tblGrid>' +
        '<w:tr><w:tc>' +
        nested +
        '<w:p><w:r><w:t>Title</w:t></w:r></w:p></w:tc>' +
        '<w:tc><w:p>' +
        ANCHOR_PIC(
          'column:<wp:posOffset>781050</wp:posOffset>',
          'paragraph:<wp:posOffset>0</wp:posOffset>',
        ) +
        '</w:p></w:tc></w:tr></w:tbl>',
    )
    const doc = await parseDocx(await buildHeaderLogoDocx(xml))
    expect(doc.headerImages!.find((i) => i.floating)!.posXPx).toBe(616)
  })
})
