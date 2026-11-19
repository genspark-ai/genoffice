import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { docxToText } from '../src/docx'
import { xlsxToText } from '../src/xlsx'
it('returns no DOCX text for a zero extraction budget', async () => {
  const d = new JSZip()
  d.file(
    'word/document.xml',
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>body</w:t></w:r></w:p></w:body></w:document>',
  )
  expect(await docxToText(await d.generateAsync({ type: 'uint8array' }), { maxChars: 0 })).toBe('')
})
it('returns no XLSX text for a zero extraction budget', async () => {
  const x = new JSZip()
  x.file(
    'xl/workbook.xml',
    '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Data" sheetId="1" r:id="s"/></sheets></workbook>',
  )
  x.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships><Relationship Id="s" Target="worksheets/sheet1.xml"/></Relationships>',
  )
  x.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet><sheetData><row><c><v>1</v></c></row></sheetData></worksheet>',
  )
  expect(await xlsxToText(await x.generateAsync({ type: 'uint8array' }), { maxChars: 0 })).toBe('')
})
