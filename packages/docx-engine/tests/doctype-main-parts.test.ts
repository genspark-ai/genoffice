import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { parseDocx } from '../src/parse'

const BODY = '<w:body><w:p><w:r><w:t>ok</w:t></w:r></w:p></w:body>'

async function docxWith(doctype: string): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>',
  )
  zip.file(
    'word/document.xml',
    `<?xml version="1.0"?>${doctype}<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${BODY}</w:document>`,
  )
  return new Uint8Array(await zip.generateAsync({ type: 'nodebuffer' }))
}

describe('DOCTYPE prologues in main parts', () => {
  it('opens a document.xml declaring external entities (used to throw and fail the whole file)', async () => {
    const bytes = await docxWith(
      '<!DOCTYPE w:document [<!ENTITY xxe SYSTEM "file:///etc/passwd">]>',
    )
    const doc = await parseDocx(bytes) // main's entry threw "External entities are not supported"
    expect(doc.blocks.length).toBeGreaterThan(0)
  })
  it('opens the same file with a plain DOCTYPE', async () => {
    const bytes = await docxWith('<!DOCTYPE w:document>')
    const doc = await parseDocx(bytes)
    expect(doc.blocks.length).toBeGreaterThan(0)
  })
})
