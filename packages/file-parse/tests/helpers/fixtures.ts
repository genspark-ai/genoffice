import { mkdtempSync, writeFileSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import JSZip from 'jszip'

/** 1x1 red PNG */
export const TINY_PNG_BASE64 =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'

let fixtureDir: string | undefined

/** write a fixture into a per-run temp dir and return its absolute path */
export function writeFixture(name: string, contents: string | Uint8Array): string {
  if (!fixtureDir) fixtureDir = mkdtempSync(join(tmpdir(), 'file-parse-'))
  const path = join(fixtureDir, name)
  writeFileSync(path, contents)
  return path
}

/** minimal but valid docx: heading + paragraph + list item + 2x2 table */
export async function buildDocxFixture(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `${XML_DECL}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'word/_rels/document.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      '</Relationships>',
  )
  zip.file(
    'word/styles.xml',
    `${XML_DECL}<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">` +
      '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>' +
      '<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/>' +
      '<w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>' +
      '</w:styles>',
  )
  const body =
    '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Annual Report</w:t></w:r></w:p>' +
    '<w:p><w:r><w:t>First paragraph hello docx</w:t></w:r></w:p>' +
    '<w:tbl><w:tr>' +
    '<w:tc><w:p><w:r><w:t>Metric</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc>' +
    '</w:tr><w:tr>' +
    '<w:tc><w:p><w:r><w:t>Revenue</w:t></w:r></w:p></w:tc>' +
    '<w:tc><w:p><w:r><w:t>100</w:t></w:r></w:p></w:tc>' +
    '</w:tr></w:tbl>'
  const sectPr =
    '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>' +
    '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>' +
    '</w:sectPr>'
  zip.file(
    'word/document.xml',
    `${XML_DECL}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      `<w:body>${body}${sectPr}</w:body></w:document>`,
  )
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

function slideXml(paragraphs: string[][]): string {
  const paras = paragraphs
    .map(
      (runs) =>
        `<a:p>${runs.map((t) => `<a:r><a:t xml:space="preserve">${t}</a:t></a:r>`).join('')}</a:p>`,
    )
    .join('')
  return (
    `${XML_DECL}<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
    `<p:cSld><p:spTree><p:sp><p:txBody>${paras}</p:txBody></p:sp></p:spTree></p:cSld></p:sld>`
  )
}

/** minimal pptx: only the slide parts our parser reads (4 slides incl. slide10 for numeric ordering) */
export async function buildPptxFixture(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file('ppt/slides/slide1.xml', slideXml([['Product', 'Intro'], ['First slide subtitle']]))
  zip.file('ppt/slides/slide2.xml', slideXml([['Market Analysis'], ['Order ', '0042']]))
  // <a:br> as an open/close pair rather than self-closing, the way XML reformatters emit it
  zip.file(
    'ppt/slides/slide3.xml',
    `${XML_DECL}<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ` +
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">' +
      '<p:cSld><p:spTree><p:sp><p:txBody><a:p>' +
      '<a:r><a:t>Before</a:t></a:r><a:br>\n  </a:br><a:br>\n  </a:br><a:r><a:t>After</a:t></a:r>' +
      '</a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>',
  )
  zip.file('ppt/slides/slide10.xml', slideXml([['Summary Slide']]))
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** minimal xlsx: shared strings (incl. rich-text runs), inline string, number, boolean, skipped column */
export async function buildXlsxFixture(): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    `${XML_DECL}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ` +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
      '<sheets><sheet name="Grades" sheetId="1" r:id="rId1"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `${XML_DECL}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      // Target is padded on purpose: xsd:anyURI collapses whitespace, so the reader must too
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target=" worksheets/sheet1.xml "/>' +
      '</Relationships>',
  )
  zip.file(
    'xl/sharedStrings.xml',
    `${XML_DECL}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="4">` +
      '<si><t>Name</t></si>' +
      '<si><r><t>Sco</t></r><r><t>res</t></r></si>' +
      '<si><t>02139</t></si>' +
      '<si><r><t xml:space="preserve">Total </t></r><r><t>due</t></r></si>' +
      '</sst>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `${XML_DECL}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>` +
      '<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row>' +
      '<row r="2"><c r="A2" t="inlineStr"><is><t>Alice</t></is></c><c r="B2"><v>95</v></c><c r="D2" t="b"><v>1</v></c></row>' +
      '<row r="3"><c r="A3" t="s"><v>2</v></c><c r="B3" t="s"><v>3</v></c></row>' +
      '<row r="4"><c r="A4" t="str"><f>A2&amp;" pts"</f><v xml:space="preserve"> Alice pts </v></c></row>' +
      '<row r="5"><c r="A5" t="e"><v xml:space="preserve"> #N/A </v></c></row>' +
      '</sheetData></worksheet>',
  )
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' })
}

/** minimal single-page pdf with an uncompressed content stream and a correct xref table */
export function buildPdfFixture(text: string): Uint8Array {
  const stream = `BT /F1 24 Tf 72 720 Td (${text}) Tj ET`
  const bodies = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>',
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ]
  let out = '%PDF-1.4\n'
  const offsets: number[] = [0]
  for (let i = 0; i < bodies.length; i++) {
    offsets.push(out.length)
    out += `${i + 1} 0 obj\n${bodies[i]}\nendobj\n`
  }
  const xrefStart = out.length
  out += `xref\n0 ${bodies.length + 1}\n0000000000 65535 f \n`
  for (let i = 1; i <= bodies.length; i++) {
    out += `${String(offsets[i]).padStart(10, '0')} 00000 n \n`
  }
  out += `trailer\n<< /Size ${bodies.length + 1} /Root 1 0 R >>\nstartxref\n${xrefStart}\n%%EOF`
  return new TextEncoder().encode(out)
}
