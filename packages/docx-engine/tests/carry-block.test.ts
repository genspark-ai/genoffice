/** A block of one document as a fragment another document can hold. */
import { describe, expect, it } from 'vitest'
import { carriableBlockXml, docxRootNamespaces } from '../src/index'
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
})
