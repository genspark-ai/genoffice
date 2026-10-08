import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const NS =
  ' xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"' +
  ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
  ' xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"' +
  ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
  ' xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"'

const LOGO_RUN =
  '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
  '<wp:extent cx="1104900" cy="595603"/><wp:docPr id="1" name="Logo"/>' +
  '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">' +
  '<pic:pic><pic:blipFill><a:blip r:embed="rId1"/></pic:blipFill></pic:pic>' +
  '</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'

const REV = ' w:id="7" w:author="A" w:date="2026-09-29T12:23:00Z"'

function header(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:hdr${NS}>${body}</w:hdr>`
}

async function parseHeader(headerXml: string) {
  return parseDocx(
    await buildDocx({
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
          xml:
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>' +
            '</Relationships>',
          contentType: 'application/vnd.openxmlformats-package.relationships+xml',
        },
      ],
    }),
  )
}

describe('tracked deletions in header/footer parts', () => {
  it('a deleted logo run leaves no strip image; deleted text leaves the strip text', async () => {
    const doc = await parseHeader(
      header(
        `<w:p><w:pPr><w:jc w:val="right"/></w:pPr><w:del${REV}>${LOGO_RUN}</w:del></w:p>` +
          `<w:p><w:r><w:t xml:space="preserve">Kept </w:t></w:r><w:del${REV}><w:r><w:delText>gone</w:delText></w:r></w:del>` +
          `<w:ins${REV}><w:r><w:t>added</w:t></w:r></w:ins></w:p>`,
      ),
    )
    expect(doc.headerImages ?? []).toHaveLength(0)
    expect(doc.headerText).toBe('Kept added')
    const part = Object.values(doc.hfParts ?? {})[0]
    expect(part?.images).toBeUndefined()
    expect(part?.text).toBe('Kept added')
  })

  it('a deleted paragraph mark (self-closing w:del) keeps the live logo', async () => {
    const doc = await parseHeader(
      header(`<w:p><w:pPr><w:rPr><w:del${REV}/></w:rPr></w:pPr>${LOGO_RUN}</w:p>`),
    )
    expect(doc.headerImages).toHaveLength(1)
  })
})
