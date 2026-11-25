/** A block of one document as a fragment another document can hold. */
import { describe, expect, it } from 'vitest'
import { carriableBlockXml, carriableSectionBreakXml, docxRootNamespaces } from '../src/index'
import { buildDocx } from './helpers/build-docx'

const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  wps: 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
  m: 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}

describe('carrying a block into another document', () => {
  it('reads the namespaces a document declares on its root', async () => {
    const namespaces = await docxRootNamespaces(await buildDocx({ bodyXml: '<w:p/>' }))
    expect(namespaces.w).toBe(NS.w)
    expect(namespaces.r).toBe(NS.r)
  })

  it('declares on the fragment what it took from the root, once', () => {
    const xml =
      '<w:p w:rsidR="00AA"><w:r><w:drawing><wp:anchor><wps:wsp xmlns:wps="urn:own"/></wp:anchor></w:drawing></w:r></w:p>'
    const out = carriableBlockXml(xml, NS)!
    expect(out.startsWith('<w:p w:rsidR="00AA" xmlns:w="')).toBe(true)
    expect(out).toContain(`xmlns:wp="${NS.wp}"`)
    expect(out).toContain(`xmlns:wps="${NS.wps}"`)
    expect(out.match(/xmlns:w=/g)).toHaveLength(1)
    expect(out).not.toContain('xmlns:m=')
    expect(out).toContain('<wps:wsp xmlns:wps="urn:own"/>')
  })

  it('takes off list numbering and bookmarks', () => {
    const out = carriableBlockXml(
      '<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr><w:jc w:val="center"/></w:pPr>' +
        '<w:bookmarkStart w:id="4" w:name="Here"/><w:r><w:t>x</w:t></w:r><w:bookmarkEnd w:id="4"/></w:p>',
      NS,
    )!
    expect(out).not.toMatch(/numPr|bookmark/)
    expect(out).toContain('<w:jc w:val="center"/>')
  })

  it('refuses a block that points outside itself', () => {
    const p = (inner: string) => `<w:p><w:r>${inner}</w:r></w:p>`
    expect(carriableBlockXml(p('<a:blip r:embed="rId5"/>'), { ...NS, a: 'urn:a' })).toBeNull()
    expect(carriableBlockXml(p('<w:footnoteReference w:id="2"/>'), NS)).toBeNull()
    expect(carriableBlockXml('<w:p><w:pPr><w:sectPr/></w:pPr></w:p>', NS)).toBeNull()
    expect(carriableBlockXml('<w:tbl/>', NS)).toBeNull()
    // a prefix nobody declared
    expect(carriableBlockXml(p('<zz:thing/>'), NS)).toBeNull()
    // the relationships namespace under a name of the document's own choosing
    expect(
      carriableBlockXml(p('<a:blip rel:embed="rId5"/>'), { ...NS, a: 'urn:a', rel: NS.r }),
    ).toBeNull()
  })

  it('takes a section break with its page setup and without the parts of its own file', () => {
    const sectPr =
      '<w:sectPr w:rsidR="00AA"><w:headerReference w:type="default" r:id="rId7"/>' +
      '<w:footerReference w:type="default" r:id="rId8"/><w:type w:val="continuous"/>' +
      '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' +
      '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/>' +
      '<w:cols w:num="2" w:space="720"/><w:printerSettings r:id="rId9"/></w:sectPr>'
    const out = carriableSectionBreakXml(
      `<w:p w:rsidR="00BB"><w:pPr><w:rPr><w:b/></w:rPr>${sectPr}</w:pPr></w:p>`,
      NS,
    )!
    expect(out).toBe(
      `<w:p xmlns:w="${NS.w}"><w:pPr><w:sectPr w:rsidR="00AA"><w:type w:val="continuous"/>` +
        '<w:pgSz w:w="16838" w:h="11906" w:orient="landscape"/>' +
        '<w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="708" w:footer="708" w:gutter="0"/>' +
        '<w:cols w:num="2" w:space="720"/></w:sectPr></w:pPr></w:p>',
    )
    // the tracked previous setup is not a second section
    expect(
      carriableSectionBreakXml(
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="1" w:h="2"/><w:sectPrChange w:id="1" w:author="a">' +
          '<w:sectPr><w:pgSz w:w="3" w:h="4"/></w:sectPr></w:sectPrChange></w:sectPr></w:pPr></w:p>',
        NS,
      ),
    ).toContain('<w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr>')
    expect(carriableSectionBreakXml('<w:p><w:r><w:t>x</w:t></w:r></w:p>', NS)).toBeNull()
    // a field paragraph that also closes a section is not an empty break
    expect(
      carriableSectionBreakXml(
        '<w:p><w:pPr><w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr></w:pPr>' +
          '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DATE</w:instrText></w:r>' +
          '<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>',
        NS,
      ),
    ).toBeNull()
    expect(
      carriableSectionBreakXml(
        '<w:p><w:bookmarkStart w:id="1" w:name="a"/><w:pPr><w:sectPr><w:pgSz w:w="1" w:h="2"/></w:sectPr></w:pPr><w:bookmarkEnd w:id="1"/></w:p>',
        NS,
      ),
    ).toContain('<w:pgSz w:w="1" w:h="2"/>')
    // a page border drawn from a picture part stays behind
    expect(
      carriableSectionBreakXml(
        '<w:p><w:pPr><w:sectPr><w:pgBorders><w:top w:val="custom" r:id="rId4"/></w:pgBorders></w:sectPr></w:pPr></w:p>',
        NS,
      ),
    ).toBeNull()
  })
})
