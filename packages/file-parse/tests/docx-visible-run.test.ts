import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { docxToText } from '../src/docx'
it.each(['off', 'false', '0'])('retains a run whose vanish flag is %s', async (flag) => {
  const z = new JSZip()
  z.file(
    'word/document.xml',
    `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:vanish w:val="${flag}"/></w:rPr><w:t>Visible</w:t></w:r><w:r><w:rPr><w:vanish/></w:rPr><w:t>Hidden</w:t></w:r></w:p></w:body></w:document>`,
  )
  expect(await docxToText(await z.generateAsync({ type: 'uint8array' }))).toBe('Visible')
})
