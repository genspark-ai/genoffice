import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
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
  })

  it('a missing diagram part degrades to the bare label, not an error', async () => {
    const parsed = await parseDocx(await buildDocx({ bodyXml: SMARTART_P }))
    const block = parsed.blocks[0]
    expect(block.label).toBe('SmartArt')
    expect(block.previewText ?? '').toBe('')
  })
})
