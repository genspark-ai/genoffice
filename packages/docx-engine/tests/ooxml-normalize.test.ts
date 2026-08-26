import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { needsOoxmlNormalization, normalizeOoxmlXml } from '../src/ooxml-normalize'
import { loadDocxZip } from '../src/zip-load'
import { parseDocx } from '../src/index'

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
const W_TRANS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
const W_STRICT = 'http://purl.oclc.org/ooxml/wordprocessingml/main'
const R_STRICT = 'http://purl.oclc.org/ooxml/officeDocument/relationships'

async function buildDocxBytes(documentXml: string, relsXml?: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  if (relsXml) zip.file('word/_rels/document.xml.rels', relsXml)
  zip.file('word/document.xml', documentXml)
  return zip.generateAsync({ type: 'uint8array' })
}

describe('needsOoxmlNormalization', () => {
  it('is false for canonical transitional parts', () => {
    expect(needsOoxmlNormalization(`<w:document xmlns:w="${W_TRANS}"><w:body/></w:document>`)).toBe(
      false,
    )
  })

  it('detects strict URIs and non-canonical prefixes', () => {
    expect(needsOoxmlNormalization(`<w:document xmlns:w="${W_STRICT}"/>`)).toBe(true)
    expect(needsOoxmlNormalization(`<fake:document xmlns:fake="${W_TRANS}"/>`)).toBe(true)
    expect(needsOoxmlNormalization(`<x:a xmlns:x="http://example.com/unrelated"/>`)).toBe(false)
  })
})

describe('non-canonical prefixes bound to known URIs', () => {
  it('renames the prefix and its declaration to the canonical one', () => {
    const out = normalizeOoxmlXml(
      `<fake:document xmlns:fake="${W_TRANS}"><fake:body>` +
        '<fake:p fake:rsidR="00A"><fake:r><fake:t>Hi</fake:t></fake:r></fake:p>' +
        '</fake:body></fake:document>',
    )
    expect(out).toContain(`<w:document xmlns:w="${W_TRANS}">`)
    expect(out).toContain('<w:p w:rsidR="00A"><w:r><w:t>Hi</w:t></w:r></w:p>')
    expect(out).not.toContain('fake:')
  })

  it('leaves unrelated namespaces alone even when local names collide', () => {
    const src = `<f:document xmlns:f="http://example.com/foo"><f:body><f:p/></f:body></f:document>`
    expect(normalizeOoxmlXml(src)).toBe(src)
  })

  it('skips the rename when the canonical prefix is taken by another namespace', () => {
    const src =
      `<x:document xmlns:x="${W_TRANS}" xmlns:w="http://example.com/other">` +
      '<x:body><x:p/><w:p/></x:body></x:document>'
    const out = normalizeOoxmlXml(src)
    expect(out).toContain('<x:body><x:p/><w:p/></x:body>')
  })

  it('frees a canonical prefix that is itself renamed away', () => {
    const a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    const out = normalizeOoxmlXml(
      `<x:document xmlns:w="${a}" xmlns:x="${W_TRANS}">` +
        '<x:body><x:p><x:r><w:blip/></x:r></x:p></x:body></x:document>',
    )
    expect(out).toContain(`<w:document xmlns:a="${a}" xmlns:w="${W_TRANS}">`)
    expect(out).toContain('<w:body><w:p><w:r><a:blip/></w:r></w:p></w:body>')
  })

  it('frees a canonical prefix regardless of declaration order', () => {
    const a = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    const out = normalizeOoxmlXml(
      `<x:document xmlns:x="${W_TRANS}" xmlns:w="${a}">` +
        '<x:body><x:p><x:r><w:blip/></x:r></x:p></x:body></x:document>',
    )
    expect(out).toContain(`<w:document xmlns:w="${W_TRANS}" xmlns:a="${a}">`)
    expect(out).toContain('<w:body><w:p><w:r><a:blip/></w:r></w:p></w:body>')
  })

  it('drops a duplicate declaration when the canonical binding already exists', () => {
    const out = normalizeOoxmlXml(
      `<w:document xmlns:w="${W_TRANS}" xmlns:ns3="${W_TRANS}">` +
        '<w:body><ns3:p><ns3:r><ns3:t>x</ns3:t></ns3:r></ns3:p></w:body></w:document>',
    )
    expect(out).toContain('<w:p><w:r><w:t>x</w:t></w:r></w:p>')
    expect(out).not.toContain('ns3')
  })

  it('prefixes elements in a default wordprocessingml namespace', () => {
    const out = normalizeOoxmlXml(
      `<document xmlns="${W_TRANS}"><body><p><r><t>x</t></r></p></body></document>`,
    )
    expect(out).toContain(`xmlns:w="${W_TRANS}"`)
    expect(out).toContain('<w:body><w:p><w:r><w:t>x</w:t></w:r></w:p></w:body>')
  })
})

describe('strict OOXML normalization', () => {
  it('maps strict namespace URIs to transitional and drops w:conformance', () => {
    const out = normalizeOoxmlXml(
      `<w:document xmlns:w="${W_STRICT}" xmlns:r="${R_STRICT}" w:conformance="strict">` +
        '<w:body><w:p><w:hyperlink r:id="rId1"/></w:p></w:body></w:document>',
    )
    expect(out).toContain(`xmlns:w="${W_TRANS}"`)
    expect(out).toContain(
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
    )
    expect(out).not.toContain('conformance')
    expect(out).toContain('<w:hyperlink r:id="rId1"/>')
  })

  it('maps strict relationship types, including divergent tails', () => {
    const out = normalizeOoxmlXml(
      `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
        `<Relationship Id="rId1" Type="${R_STRICT}/styles" Target="styles.xml"/>` +
        `<Relationship Id="rId2" Type="${R_STRICT}/extendedProperties" Target="../docProps/app.xml"/>` +
        '</Relationships>',
    )
    expect(out).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"',
    )
    expect(out).toContain(
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"',
    )
  })

  it('converts universal measures to the native unit of each attribute', () => {
    const out = normalizeOoxmlXml(
      `<w:document xmlns:w="${W_STRICT}"><w:body>` +
        '<w:p><w:pPr><w:ind w:start="72pt" w:hanging="18pt"/>' +
        '<w:spacing w:before="12pt" w:line="12pt"/>' +
        '<w:pBdr><w:top w:val="single" w:sz="0.5pt" w:space="1pt"/></w:pBdr>' +
        '<w:rPr><w:sz w:val="12pt"/><w:position w:val="-3pt"/></w:rPr></w:pPr></w:p>' +
        '<w:sectPr><w:pgSz w:w="792pt" w:h="612pt" w:orient="landscape"/>' +
        '<w:pgMar w:top="72pt" w:right="72pt" w:bottom="72pt" w:left="72pt" w:header="36pt" w:footer="36pt" w:gutter="0pt"/>' +
        '</w:sectPr></w:body></w:document>',
    )
    expect(out).toContain('<w:ind w:start="1440" w:hanging="360"/>')
    expect(out).toContain('<w:spacing w:before="240" w:line="240"/>')
    expect(out).toContain('<w:top w:val="single" w:sz="4" w:space="1"/>')
    expect(out).toContain('<w:sz w:val="24"/>')
    expect(out).toContain('<w:position w:val="-6"/>')
    expect(out).toContain('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>')
    expect(out).toContain('w:header="720" w:footer="720" w:gutter="0"')
  })

  it('leaves plain numeric values and non-w namespaces untouched', () => {
    const src =
      `<w:document xmlns:w="${W_STRICT}" xmlns:x="http://example.com/x"><w:body>` +
      '<w:p><w:pPr><w:ind w:start="720"/></w:pPr></w:p>' +
      '<x:ind x:start="72pt"/></w:body></w:document>'
    const out = normalizeOoxmlXml(src)
    expect(out).toContain('<w:ind w:start="720"/>')
    expect(out).toContain('<x:ind x:start="72pt"/>')
  })

  it('shifts wordprocessingShape children written in the wp namespace into wps', () => {
    const out = normalizeOoxmlXml(
      `<w:document xmlns:w="${W_STRICT}" xmlns:wp="http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing">` +
        '<w:body><w:p><w:r><w:drawing><wp:anchor><wp:extent cx="100" cy="100"/>' +
        '<a:graphic xmlns:a="http://purl.oclc.org/ooxml/drawingml/main">' +
        '<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">' +
        '<wp:wsp><wp:cNvSpPr/><wp:spPr><a:prstGeom prst="rect"/></wp:spPr><wp:bodyPr/></wp:wsp>' +
        '</a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:body></w:document>',
    )
    expect(out).toContain(
      '<wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">',
    )
    expect(out).toContain('<wps:cNvSpPr/><wps:spPr><a:prstGeom prst="rect"/></wps:spPr>')
    expect(out).toContain('<wps:bodyPr/></wps:wsp>')
    // the anchor level stays in the wp namespace
    expect(out).toContain('<wp:anchor><wp:extent cx="100" cy="100"/>')
    expect(out).toContain('xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"')
  })

  it('keeps mc:Ignorable prefix lists aligned with renames', () => {
    const out = normalizeOoxmlXml(
      `<w:document xmlns:w="${W_STRICT}"` +
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' xmlns:cx="http://schemas.microsoft.com/office/word/2010/wordml"' +
        ' mc:Ignorable="cx"><w:body><w:p cx:paraId="1"/></w:body></w:document>',
    )
    expect(out).toContain('mc:Ignorable="w14"')
    expect(out).toContain('<w:p w14:paraId="1"/>')
  })

  it('remaps Ignorable tokens when the mc namespace itself uses another prefix', () => {
    const out = normalizeOoxmlXml(
      `<w:document xmlns:w="${W_STRICT}"` +
        ' xmlns:m2="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' xmlns:cx="http://schemas.microsoft.com/office/word/2010/wordml"' +
        ' m2:Ignorable="cx"><w:body/></w:document>',
    )
    expect(out).toContain('mc:Ignorable="w14"')
    expect(out).toContain('xmlns:mc=')
  })
})

describe('parseDocx over normalized packages', () => {
  it('parses a document with a non-standard prefix for the main namespace', async () => {
    const bytes = await buildDocxBytes(
      `${XML_DECL}<fake:document xmlns:fake="${W_TRANS}"><fake:body>` +
        '<fake:p><fake:r><fake:t>Body text survives</fake:t></fake:r></fake:p>' +
        '</fake:body></fake:document>',
    )
    const doc = await parseDocx(bytes)
    expect(
      doc.blocks
        .flatMap((b) => b.runs ?? [])
        .map((r) => r.text)
        .join(''),
    ).toContain('Body text survives')
  })

  it('parses a strict document with landscape geometry in pt units', async () => {
    const bytes = await buildDocxBytes(
      `${XML_DECL}<w:document xmlns:w="${W_STRICT}" w:conformance="strict"><w:body>` +
        '<w:p><w:r><w:t>Strict text</w:t></w:r></w:p>' +
        '<w:sectPr><w:pgSz w:w="792pt" w:h="612pt" w:orient="landscape"/></w:sectPr>' +
        '</w:body></w:document>',
    )
    const doc = await parseDocx(bytes)
    expect(
      doc.blocks
        .flatMap((b) => b.runs ?? [])
        .map((r) => r.text)
        .join(''),
    ).toContain('Strict text')
    const zip = await loadDocxZip(bytes)
    const normalized = await zip.file('word/document.xml')!.async('string')
    expect(normalized).toContain('<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/>')
    expect(normalized).not.toContain('conformance')
  })
})
