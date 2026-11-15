import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { xlsxToText } from '../src/xlsx'
it.each(['', ' s="0"'])(
  'uses default cell style zero when style attribute is %s',
  async (attribute) => {
    const zip = new JSZip()
    zip.file(
      'xl/workbook.xml',
      '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Dates" r:id="s"/></sheets></workbook>',
    )
    zip.file(
      'xl/_rels/workbook.xml.rels',
      '<Relationships><Relationship Id="s" Target="worksheets/s.xml"/></Relationships>',
    )
    zip.file('xl/styles.xml', '<styleSheet><cellXfs><xf numFmtId="14"/></cellXfs></styleSheet>')
    zip.file(
      'xl/worksheets/s.xml',
      `<worksheet><sheetData><row><c r="A1"${attribute}><v>45292</v></c></row></sheetData></worksheet>`,
    )
    expect(await xlsxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
      '# Dates\n2024-01-01',
    )
  },
)
