import { describe, expect, it } from 'vitest'
import { parseDocx, saveDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const DIAGRAM_DATA_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst>' +
  ['总经理办', '研发部', '产品部', '销售部', '运营部']
    .map(
      (t, i) =>
        `<dgm:pt modelId="{${i}}"><dgm:t><a:bodyPr/><a:p><a:r><a:t>${t}</a:t></a:r></a:p></dgm:t></dgm:pt>`,
    )
    .join('') +
  '</dgm:ptLst></dgm:dataModel>'

const SMARTART_P =
  '<w:p><w:r><w:drawing><wp:inline><wp:extent cx="5486400" cy="419100"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram">' +
  '<dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" ' +
  'r:dm="rId40" r:lo="rId41" r:qs="rId42" r:cs="rId43"/>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>'

const OLE_P =
  '<w:p><w:r><w:object><v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="_x0000_i1025" ' +
  'style="width:32pt;height:32pt"><v:imagedata r:id="rId10" o:title=""/></v:shape>' +
  '<o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" Type="Embed" ' +
  'ProgID="Excel.Sheet.12" ShapeID="_x0000_i1025" DrawAspect="Icon" ObjectID="_1"/>' +
  '</w:object></w:r></w:p>'

async function buildFixture(): Promise<Uint8Array> {
  return buildDocx({
    bodyXml: SMARTART_P + OLE_P,
    withImage: true,
    extraRels:
      '<Relationship Id="rId40" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/>',
    extraParts: [
      {
        path: 'word/diagrams/data1.xml',
        xml: DIAGRAM_DATA_XML,
        contentType: 'application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml',
      },
    ],
  })
}

describe('SmartArt / OLE visible degrade', () => {
  it('SmartArt preview lists the diagram node texts', async () => {
    const parsed = await parseDocx(await buildFixture())
    const block = parsed.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('SmartArt')
    expect(block.previewText).toBe('总经理办\n研发部\n产品部\n销售部\n运营部')
  })

  it('OLE embed surfaces the packaged preview picture and ProgID', async () => {
    const parsed = await parseDocx(await buildFixture())
    const block = parsed.blocks[1]
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Embedded object')
    expect(block.oleProgId).toBe('Excel.Sheet.12')
    expect(block.imageDataUrl).toMatch(/^data:image\/png;base64,/)
    // declared v:shape size (32pt = 43px), not the preview's intrinsic pixels
    expect(block.imageWidthPx).toBe(43)
    expect(block.imageHeightPx).toBe(43)
  })

  it('OLE follows the paragraph w:jc like other inline objects', async () => {
    const oleP = OLE_P.replace('<w:r>', '<w:pPr><w:jc w:val="center"/></w:pPr><w:r>')
    const parsed = await parseDocx(await buildDocx({ bodyXml: oleP, withImage: true }))
    expect(parsed.blocks[0].imageAlign).toBe('center')
  })

  it('OLE size falls back to w:object dxaOrig/dyaOrig (twips)', async () => {
    const oleP = OLE_P.replace('<w:object>', '<w:object w:dxaOrig="1365" w:dyaOrig="765">').replace(
      ' style="width:32pt;height:32pt"',
      '',
    )
    const parsed = await parseDocx(await buildDocx({ bodyXml: oleP, withImage: true }))
    const block = parsed.blocks[0]
    expect(block.imageWidthPx).toBe(91)
    expect(block.imageHeightPx).toBe(51)
  })

  it('VML horizontal rule (v:rect o:hr) is a decorative line, not a chip', async () => {
    const hrP =
      '<w:p><w:r><w:pict><v:rect xmlns:v="urn:schemas-microsoft-com:vml" ' +
      'xmlns:o="urn:schemas-microsoft-com:office:office" id="_x0000_i1026" ' +
      'style="width:0;height:1.5pt" o:hralign="center" o:hrstd="t" o:hr="t" ' +
      'fillcolor="#aca899" stroked="f"/></w:pict></w:r></w:p>'
    const parsed = await parseDocx(await buildDocx({ bodyXml: hrP }))
    const block = parsed.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.decorative).toBe(true)
    expect(block.ruleColorHex).toBe('ACA899')
    expect(block.ruleThicknessPx).toBe(2)
    expect(block.ruleWidthPx).toBeUndefined()
  })

  it('a missing diagram part degrades to the bare label, not an error', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: SMARTART_P }))
    const block = parsed.blocks[0]
    expect(block.label).toBe('SmartArt')
    expect(block.previewText ?? '').toBe('')
  })
})

const OLE_OBJECT_RUN =
  '<w:r><w:object w:dxaOrig="4150" w:dyaOrig="310">' +
  '<v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="_x0000_i1025" ' +
  'style="width:207.75pt;height:15.75pt"><v:imagedata r:id="rId10" o:title=""/></v:shape>' +
  '<o:OLEObject xmlns:o="urn:schemas-microsoft-com:office:office" Type="Embed" ' +
  'ProgID="Excel.Sheet.8" ShapeID="_x0000_i1025" DrawAspect="Content" ObjectID="_1"/>' +
  '</w:object></w:r>'

describe('OLE embed sharing a paragraph with real text', () => {
  const bodyXml =
    '<w:p><w:r><w:rPr><w:color w:val="0000FF"/></w:rPr>' +
    '<w:t xml:space="preserve">Состоится пресс-конференция </w:t></w:r>' +
    '<w:r><w:rPr><w:b/><w:color w:val="FF0000"/></w:rPr><w:t>партнерств</w:t></w:r>' +
    OLE_OBJECT_RUN +
    '</w:p>'

  it('stays an editable paragraph; the object becomes a run-level image', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    const block = parsed.blocks[0]
    expect(block.type).toBe('paragraph')
    expect((block.runs ?? []).map((r) => r.text).join('')).toBe(
      'Состоится пресс-конференция партнерств',
    )
    expect(block.runs?.[0].color).toBe('0000FF')
    expect(block.runs?.[1].bold).toBe(true)
    const imgRun = (block.runs ?? []).find((r) => r.image)
    expect(imgRun?.image?.dataUrl).toMatch(/^data:image\/png;base64,/)
    expect(imgRun?.image?.xml).toContain('<o:OLEObject')
    // declared v:shape size (207.75pt × 15.75pt)
    expect(imgRun?.image?.widthPx).toBe(277)
    expect(imgRun?.image?.heightPx).toBe(21)
  })

  it('run image size falls back to w:object dxaOrig/dyaOrig (twips)', async () => {
    const noStyle = bodyXml.replace(' style="width:207.75pt;height:15.75pt"', '')
    const parsed = await parseDocx(await buildDocx({ bodyXml: noStyle, withImage: true }))
    const imgRun = (parsed.blocks[0].runs ?? []).find((r) => r.image)
    expect(imgRun?.image?.widthPx).toBe(277)
    expect(imgRun?.image?.heightPx).toBe(21)
  })

  it('the w:object fragment round-trips through an edited save', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    const runs = parsed.blocks[0].runs!
    const saved = await saveDocx(parsed, [
      { kind: 'generated', block: { type: 'paragraph', runs } },
    ])
    const reparsed = await parseDocx(saved)
    const block = reparsed.blocks[0]
    expect(block.type).toBe('paragraph')
    expect((block.runs ?? []).map((r) => r.text).join('')).toBe(
      'Состоится пресс-конференция партнерств',
    )
    const imgRun = (block.runs ?? []).find((r) => r.image)
    expect(imgRun?.image?.xml).toContain('<o:OLEObject')
  })

  it('an unresolvable preview keeps the protected chip with the text visible', async () => {
    // no withImage: rId10 has no relationship, the preview cannot resolve
    const parsed = await parseDocx(await buildDocx({ bodyXml }))
    const block = parsed.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Embedded object')
    expect(block.previewText).toContain('Состоится пресс-конференция')
  })
})

function fieldOleParagraph(instr: string): string {
  return (
    '<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
    `<w:r><w:instrText xml:space="preserve">${instr}</w:instrText></w:r>` +
    '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' +
    OLE_OBJECT_RUN +
    '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>'
  )
}

describe('field-form OLE (EMBED/LINK field around w:object)', () => {
  it('EMBED field takes the OLE display path: preview picture, ProgID, declared size', async () => {
    const bodyXml = fieldOleParagraph(' EMBED Excel.Sheet.8 ')
    const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    const block = parsed.blocks[0]
    expect(block.type).toBe('passthrough')
    expect(block.label).toBe('Embedded object')
    expect(block.oleProgId).toBe('Excel.Sheet.8')
    expect(block.imageDataUrl).toMatch(/^data:image\/png;base64,/)
    expect(block.imageWidthPx).toBe(277)
    expect(block.imageHeightPx).toBe(21)
  })

  it('LINK field (linked OLE) takes the same path', async () => {
    const bodyXml = fieldOleParagraph(
      ' LINK Excel.Sheet.8 "Book1.xlsx" "Sheet1!R1C1:R2C2" \\a \\p ',
    ).replace('Type="Embed"', 'Type="Link"')
    const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    const block = parsed.blocks[0]
    expect(block.label).toBe('Embedded object')
    expect(block.oleProgId).toBe('Excel.Sheet.8')
    expect(block.imageDataUrl).toMatch(/^data:image\/png;base64,/)
  })

  it('a paragraph mixing EMBED with another field keeps the protected field chip', async () => {
    const bodyXml = fieldOleParagraph(' EMBED Excel.Sheet.8 ').replace(
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
      '<w:r><w:fldChar w:fldCharType="end"/></w:r>' +
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>' +
        '<w:r><w:instrText xml:space="preserve"> TOC \\o </w:instrText></w:r>' +
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>',
    )
    const parsed = await parseDocx(await buildDocx({ bodyXml, withImage: true }))
    const block = parsed.blocks[0]
    expect(block.label).toBe('Field (EMBED)')
    expect(block.imageDataUrl).toBeUndefined()
  })

  it('no edits -> byte-identical save (passthrough keeps the field XML)', async () => {
    const bytes = await buildDocx({
      bodyXml: fieldOleParagraph(' EMBED Excel.Sheet.8 '),
      withImage: true,
    })
    const doc = await parseDocx(bytes)
    const saved = await saveDocx(
      doc,
      doc.blocks
        .filter((b) => !b.hidden)
        .map((b) => ({ kind: 'original' as const, docxIndex: b.docxIndex! })),
    )
    expect(saved).toBe(bytes)
  })
})
