import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { NEW_PICTURE_TOKEN_PREFIX, parseDocx, saveDocx, type SaveBlock } from '../src/index'
import { buildDocx } from './helpers/build-docx'

/** header edits inside a layout table: only cells flagged `rewrite` are
 *  regenerated inside the original w:tbl; a picture inserted by the editor
 *  lands in the part's own rels */

const NS =
  ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'

const LOGO =
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="381000" cy="190500"/><wp:docPr id="1" name="Logo"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="logo.png"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill>' +
  '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="381000" cy="190500"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'

const HEADER_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<w:hdr${NS}>` +
  '<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblLayout w:type="fixed"/></w:tblPr>' +
  '<w:tblGrid><w:gridCol w:w="2835"/><w:gridCol w:w="7370"/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:tcW w:w="2835" w:type="dxa"/><w:shd w:val="clear" w:fill="EEEEEE"/></w:tcPr>' +
  `<w:p>${LOGO}</w:p></w:tc>` +
  '<w:tc><w:tcPr><w:tcW w:w="7370" w:type="dxa"/></w:tcPr>' +
  '<w:p><w:pPr><w:spacing w:after="40"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="23"/></w:rPr><w:t>Old Company</w:t></w:r></w:p>' +
  '<w:p><w:r><w:t>OLD SCIENCE</w:t></w:r></w:p></w:tc></w:tr></w:tbl>' +
  '<w:p><w:pPr><w:pBdr><w:bottom w:val="single" w:sz="8" w:space="1" w:color="000000"/></w:pBdr></w:pPr></w:p>' +
  '</w:hdr>'

const HEADER_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
  '</Relationships>'

async function build(): Promise<Uint8Array> {
  return buildDocx({
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
}

async function headerPart(saved: Uint8Array, name = 'word/header1.xml'): Promise<string> {
  const zip = await JSZip.loadAsync(saved)
  return (await zip.file(name)?.async('string')) ?? ''
}

describe('header layout-table cell edits', () => {
  it('parses rows with their table / row source and the logo on the cell run', async () => {
    const doc = await parseDocx(await build())
    const row = doc.headerParas![0]
    expect(row.row?.source).toEqual({ tbl: 0, tr: 0 })
    expect(row.cells![0].paras[0][0].image?.xml).toContain('<w:drawing>')
    expect(row.cells![1].paras.map((p) => p.map((r) => r.text).join(''))).toEqual([
      'Old Company',
      'OLD SCIENCE',
    ])
    expect(row.cells![1].locked).toBeUndefined()
  })

  it('regenerates only the rewritten cell inside the original table bytes', async () => {
    const doc = await parseDocx(await build())
    const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: doc.blocks[0].docxIndex! }]
    const row = doc.headerParas![0]
    const edited = {
      text: '',
      paras: [
        {
          ...row,
          cells: [
            row.cells![0],
            {
              ...row.cells![1],
              rewrite: true as const,
              paras: [[{ text: 'New Company', bold: true }], [{ text: 'NEW SCIENCE' }]],
            },
          ],
        },
        ...doc.headerParas!.slice(1),
      ],
    }
    const saved = await saveDocx(doc, saveBlocks, { header: edited })
    const xml = await headerPart(saved)
    // untouched logo cell keeps its bytes, the text cell is regenerated with its pPr
    expect(xml).toContain('<w:shd w:val="clear" w:fill="EEEEEE"/>')
    expect(xml).toContain('<a:blip r:embed="rId1"/>')
    expect(xml).toContain('<w:spacing w:after="40"/>')
    expect(xml).toContain('New Company')
    expect(xml).toContain('NEW SCIENCE')
    expect(xml).not.toContain('Old Company')
    expect(xml).toContain('<w:pBdr>')
    expect((xml.match(/<w:tbl>/g) ?? []).length).toBe(1)
    const reparsed = await parseDocx(saved)
    expect(reparsed.headerParas![0].cells![1].paras[0][0]).toMatchObject({
      text: 'New Company',
      bold: true,
    })
    expect(reparsed.headerParas![0].cells![0].paras[0][0].image).toBeTruthy()
  })

  it('rebuilds only the flagged paragraph pPr groups from paraProps', async () => {
    const doc = await parseDocx(await build())
    const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: doc.blocks[0].docxIndex! }]
    const row = doc.headerParas![0]
    const cell = row.cells![1]
    const edited = {
      text: '',
      paras: [
        {
          ...row,
          cells: [
            row.cells![0],
            {
              ...cell,
              rewrite: true as const,
              rewriteFormat: [0],
              paraProps: [{ ...cell.paraProps?.[0], align: 'right' as const }, cell.paraProps?.[1]],
            },
          ],
        },
        ...doc.headerParas!.slice(1),
      ],
    }
    const saved = await saveDocx(doc, saveBlocks, { header: edited })
    const xml = await headerPart(saved)
    // the first paragraph gains w:jc and keeps its w:spacing; the second keeps its (empty) pPr
    expect(xml).toMatch(/<w:p><w:pPr><w:spacing w:after="40"\/><w:jc w:val="right"\/><\/w:pPr>/)
    const reparsed = await parseDocx(saved)
    expect(reparsed.headerParas![0].cells![1].paraProps![0]).toMatchObject({ align: 'right' })
  })

  it('a format edit keeps unmodeled pPr children and indent pieces, and does not write style-resolved groups', async () => {
    const richPPr =
      '<w:pPr><w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="000000"/></w:pBdr>' +
      '<w:spacing w:after="40" w:afterAutospacing="0"/><w:ind w:left="100" w:right="200"/></w:pPr>'
    const headerXml = HEADER_XML.replace('<w:pPr><w:spacing w:after="40"/></w:pPr>', richPPr)
    const bytes = await buildDocx({
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
    const doc = await parseDocx(bytes)
    const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: doc.blocks[0].docxIndex! }]
    const row = doc.headerParas![0]
    const cell = row.cells![1]
    const prior = cell.paraProps?.[0]
    expect(prior).toMatchObject({ spaceAfter: 40, indentLeft: 100 })
    const save = async (props: NonNullable<typeof prior>) => {
      const edited = {
        text: '',
        paras: [
          {
            ...row,
            cells: [
              row.cells![0],
              {
                ...cell,
                rewrite: true as const,
                rewriteFormat: [0],
                priorParaProps: cell.paraProps,
                paraProps: [props, cell.paraProps?.[1]],
              },
            ],
          },
          ...doc.headerParas!.slice(1),
        ],
      }
      return headerPart(await saveDocx(doc, saveBlocks, { header: edited }))
    }
    // alignment only: borders, spacing and the indent element are byte-identical
    const aligned = await save({ ...prior, align: 'right' })
    expect(aligned).toContain(
      '<w:pBdr><w:bottom w:val="single" w:sz="4" w:space="1" w:color="000000"/></w:pBdr>',
    )
    expect(aligned).toContain('<w:spacing w:after="40" w:afterAutospacing="0"/>')
    expect(aligned).toContain('<w:ind w:left="100" w:right="200"/>')
    expect(aligned).toContain('<w:jc w:val="right"/>')
    // left indent only: the unmodeled right indent rides along, no w:jc appears
    const indented = await save({ ...prior, indentLeft: 300 })
    expect(indented).toMatch(/<w:ind w:left="300"[^>]*w:right="200"\/>/)
    expect(indented).not.toContain('<w:jc')
    expect(indented).toContain('<w:pBdr>')
    // space after only: autospacing attribute is replaced by the modeled value, nothing else moves
    const spaced = await save({ ...prior, spaceAfter: 80 })
    expect(spaced).toContain('<w:spacing w:after="80"/>')
    expect(spaced).toContain('<w:ind w:left="100" w:right="200"/>')
    expect(spaced).not.toContain('<w:jc')
    // cleared fields go away instead of being copied back; the unmodeled right indent survives alone
    const { spaceAfter: _sa, indentLeft: _il, ...cleared } = prior!
    const clearedXml = await save(cleared)
    const textCell = clearedXml.slice(
      clearedXml.indexOf('<w:tc>', clearedXml.indexOf('</w:tc>')),
      clearedXml.indexOf('</w:tbl>'),
    )
    expect(textCell).not.toMatch(/<w:spacing\b/)
    expect(textCell).toContain('<w:ind w:right="200"/>')
    expect(textCell).toContain('<w:pBdr>')
  })

  it('a sibling edit keeps explicit zeros the cell model does not carry', async () => {
    const zeroPPr =
      '<w:pPr><w:spacing w:before="120" w:after="0"/><w:ind w:left="0" w:firstLine="200"/></w:pPr>'
    const headerXml = HEADER_XML.replace('<w:pPr><w:spacing w:after="40"/></w:pPr>', zeroPPr)
    const bytes = await buildDocx({
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
    const doc = await parseDocx(bytes)
    const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: doc.blocks[0].docxIndex! }]
    const row = doc.headerParas![0]
    const cell = row.cells![1]
    const prior = cell.paraProps?.[0]
    // the parser drops the zeros: the model alone cannot regenerate them
    expect(prior).toMatchObject({ spaceBefore: 120, indentFirstLine: 200 })
    expect(prior?.spaceAfter).toBeUndefined()
    expect(prior?.indentLeft).toBeUndefined()
    const save = async (props: NonNullable<typeof prior>) => {
      const edited = {
        text: '',
        paras: [
          {
            ...row,
            cells: [
              row.cells![0],
              {
                ...cell,
                rewrite: true as const,
                rewriteFormat: [0],
                priorParaProps: cell.paraProps,
                paraProps: [props, cell.paraProps?.[1]],
              },
            ],
          },
          ...doc.headerParas!.slice(1),
        ],
      }
      return headerPart(await saveDocx(doc, saveBlocks, { header: edited }))
    }
    const spaced = await save({ ...prior, spaceBefore: 240 })
    expect(spaced).toMatch(/<w:spacing w:before="240" w:after="0"\/>/)
    expect(spaced).toContain('<w:ind w:left="0" w:firstLine="200"/>')
    const indented = await save({ ...prior, indentFirstLine: 400 })
    expect(indented).toContain('<w:spacing w:before="120" w:after="0"/>')
    expect(indented).toMatch(/<w:ind w:firstLine="400" w:left="0"\/>/)
  })

  it('clearing every tab stop removes w:tabs; a self-closing pPr only gains the changed group', async () => {
    const tabbedPPr =
      '<w:pPr><w:tabs><w:tab w:val="right" w:pos="9000"/></w:tabs><w:spacing w:after="40"/></w:pPr>'
    const headerXml = HEADER_XML.replace(
      '<w:pPr><w:spacing w:after="40"/></w:pPr>',
      tabbedPPr,
    ).replace(
      '<w:p><w:r><w:t>OLD SCIENCE</w:t></w:r></w:p>',
      '<w:p><w:pPr/><w:r><w:t>OLD SCIENCE</w:t></w:r></w:p>',
    )
    const bytes = await buildDocx({
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
    const doc = await parseDocx(bytes)
    const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: doc.blocks[0].docxIndex! }]
    const row = doc.headerParas![0]
    const cell = row.cells![1]
    expect(cell.paraProps?.[0]?.tabStops).toHaveLength(1)
    const edited = {
      text: '',
      paras: [
        {
          ...row,
          cells: [
            row.cells![0],
            {
              ...cell,
              rewrite: true as const,
              rewriteFormat: [0, 1],
              priorParaProps: cell.paraProps,
              paraProps: [
                { ...cell.paraProps![0], tabStops: undefined },
                // the editor reports style-resolved spacing on the second paragraph and changes its alignment
                { ...cell.paraProps?.[1], spaceAfter: 120, align: 'center' as const },
              ],
            },
          ],
        },
        ...doc.headerParas!.slice(1),
      ],
    }
    // the second paragraph's spacing was "already there" as far as the editor is concerned
    edited.paras[0].cells![1].priorParaProps = [cell.paraProps![0], { spaceAfter: 120 }]
    const xml = await headerPart(await saveDocx(doc, saveBlocks, { header: edited }))
    expect(xml).not.toContain('<w:tabs>')
    expect(xml).toContain('<w:spacing w:after="40"/>')
    expect(xml).toContain('<w:p><w:pPr><w:jc w:val="center"/></w:pPr>')
    expect(xml).not.toContain('w:after="120"')
  })

  it('patches a table wrapped in a w:sdt content control', async () => {
    const sdtHeader = HEADER_XML.replace(
      '<w:tbl>',
      '<w:sdt><w:sdtPr/><w:sdtContent><w:tbl>',
    ).replace('</w:tbl>', '</w:tbl></w:sdtContent></w:sdt>')
    const bytes = await buildDocx({
      bodyXml: '<w:p><w:r><w:t>Body</w:t></w:r></w:p>',
      withImage: true,
      extraRels:
        '<Relationship Id="rId20" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>',
      sectPrExtra: '<w:headerReference w:type="default" r:id="rId20"/>',
      extraParts: [
        {
          path: 'word/header1.xml',
          xml: sdtHeader,
          contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
        },
        {
          path: 'word/_rels/header1.xml.rels',
          xml: HEADER_RELS,
          contentType: 'application/vnd.openxmlformats-package.relationships+xml',
        },
      ],
    })
    const doc = await parseDocx(bytes)
    const row = doc.headerParas![0]
    expect(row.row?.source).toEqual({ tbl: 0, tr: 0 })
    const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: doc.blocks[0].docxIndex! }]
    const edited = {
      text: '',
      paras: [
        {
          ...row,
          cells: [
            row.cells![0],
            { ...row.cells![1], rewrite: true as const, paras: [[{ text: 'Wrapped Co' }]] },
          ],
        },
        ...doc.headerParas!.slice(1),
      ],
    }
    const saved = await saveDocx(doc, saveBlocks, { header: edited })
    const xml = await headerPart(saved)
    expect(xml).toContain('<w:sdt><w:sdtPr/><w:sdtContent><w:tbl>')
    expect(xml).toContain('Wrapped Co')
    expect(xml).not.toContain('Old Company')
    expect(xml).toContain('<a:blip r:embed="rId1"/>')
  })

  it('keeps the table bytes untouched when no cell is flagged', async () => {
    const doc = await parseDocx(await build())
    const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: doc.blocks[0].docxIndex! }]
    const paras = [doc.headerParas![0], { ...doc.headerParas![1], runs: [{ text: 'Ruled line' }] }]
    const saved = await saveDocx(doc, saveBlocks, { header: { text: 'Ruled line', paras } })
    const xml = await headerPart(saved)
    const original = HEADER_XML.slice(HEADER_XML.indexOf('<w:tbl>'), HEADER_XML.indexOf('</w:tbl>'))
    expect(xml).toContain(original)
    expect(xml).toContain('Ruled line')
  })

  it('lands a picture inserted into a cell as media plus a relationship of the header part', async () => {
    const doc = await parseDocx(await build())
    const saveBlocks: SaveBlock[] = [{ kind: 'original', docxIndex: doc.blocks[0].docxIndex! }]
    const token = `${NEW_PICTURE_TOKEN_PREFIX}7`
    const drawing =
      '<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="190500" cy="190500"/>' +
      `<wp:docPr id="${token}" name="Picture ${token}"/>` +
      '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
      `<pic:pic><pic:nvPicPr><pic:cNvPr id="${token}" name="Picture ${token}"/><pic:cNvPicPr/></pic:nvPicPr>` +
      `<pic:blipFill><a:blip r:embed="${token}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
      '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="190500" cy="190500"/></a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic>' +
      '</a:graphicData></a:graphic></wp:inline></w:drawing>'
    const row = doc.headerParas![0]
    const png =
      'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
    const edited = {
      text: '',
      paras: [
        {
          ...row,
          cells: [
            row.cells![0],
            {
              ...row.cells![1],
              rewrite: true as const,
              paras: [
                [
                  { text: 'Seal ' },
                  {
                    text: '',
                    image: { dataUrl: `data:image/png;base64,${png}`, xml: drawing },
                  },
                ],
              ],
            },
          ],
        },
        ...doc.headerParas!.slice(1),
      ],
      newPictures: { [token]: { base64: png, mime: 'image/png' as const } },
    }
    const saved = await saveDocx(doc, saveBlocks, { header: edited })
    const xml = await headerPart(saved)
    const rels = await headerPart(saved, 'word/_rels/header1.xml.rels')
    expect(xml).not.toContain(NEW_PICTURE_TOKEN_PREFIX)
    const rId = /<a:blip r:embed="(rId\d+)"\/><a:stretch>/.exec(xml)?.[1]
    expect(rId).toBeTruthy()
    expect(rId).not.toBe('rId1')
    expect(rels).toContain(`Id="${rId}"`)
    expect(rels).toContain('Target="media/aidocs')
    expect(xml).toMatch(/<wp:docPr id="\d+" name="Picture \d+"\/>/)
    const reparsed = await parseDocx(saved)
    const cell = reparsed.headerParas![0].cells![1]
    expect(cell.paras[0].map((r) => [r.text, !!r.image])).toEqual([
      ['Seal ', false],
      ['', true],
    ])
  })
})
