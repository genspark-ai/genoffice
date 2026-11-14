import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { docxToText } from '../src/docx'
it.each([
  ['9', 'Body'],
  ['2', '### Body'],
])('honors explicit outline level %s over a heading name', async (level, expected) => {
  const zip = new JSZip()
  zip.file(
    'word/document.xml',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Custom"/></w:pPr><w:r><w:t>Body</w:t></w:r></w:p></w:body></w:document>',
  )
  zip.file(
    'word/styles.xml',
    `<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Custom"><w:name w:val="Heading 1"/><w:pPr><w:outlineLvl w:val="${level}"/></w:pPr></w:style></w:styles>`,
  )
  expect(await docxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(expected)
})
