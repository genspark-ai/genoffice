import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { docxToText } from '../src/docx'
it('reads the declared main document even when the conventional path exists', async () => {
  const zip = new JSZip()
  const document = (text: string) =>
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>' +
    text +
    '</w:t></w:r></w:p></w:body></w:document>'
  zip.file('word/document.xml', document('Unreferenced copy'))
  zip.file('active.xml', document('Active document'))
  zip.file(
    '_rels/.rels',
    '<Relationships><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="active.xml"/></Relationships>',
  )
  expect(await docxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe('Active document')
})
it('keeps conventional-path compatibility for packages without root relationships', async () => {
  const zip = new JSZip()
  zip.file(
    'word/document.xml',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>Legacy</w:t></w:r></w:p></w:body></w:document>',
  )
  expect(await docxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe('Legacy')
})
