import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { xlsxToText } from '../src/xlsx'
it('locates workbook, worksheets and metadata through relocated package parts', async () => {
  const zip = new JSZip()
  zip.file(
    '_rels/.rels',
    '<Relationships><Relationship Id="main" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="custom/book.xml"/></Relationships>',
  )
  zip.file(
    'custom/book.xml',
    '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Relocated" r:id="s"/></sheets></workbook>',
  )
  zip.file(
    'custom/_rels/book.xml.rels',
    '<Relationships><Relationship Id="s" Target="../sheets/one.xml"/><Relationship Id="t" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="text.xml"/></Relationships>',
  )
  zip.file('custom/text.xml', '<sst><si><t>Found</t></si></sst>')
  zip.file(
    'sheets/one.xml',
    '<worksheet><sheetData><row><c t="s"><v>0</v></c></row></sheetData></worksheet>',
  )
  expect(await xlsxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
    '# Relocated\nFound',
  )
})
