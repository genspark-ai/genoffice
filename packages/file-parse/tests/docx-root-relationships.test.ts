import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { docxToText } from '../src/docx'
it('follows namespace-qualified root relationships with XML attribute quoting', async () => {
  const zip = new JSZip()
  zip.file(
    '_rels/.rels',
    "<r:Relationships xmlns:r='http://schemas.openxmlformats.org/package/2006/relationships'><r:Relationship Id='main' Type='http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument' Target='parts/main&amp;body.xml'/></r:Relationships>",
  )
  zip.file(
    'parts/main&body.xml',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Relocated</w:t></w:r></w:p></w:body></w:document>',
  )
  expect(await docxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe('Relocated')
})
