import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { xlsxToText } from '../src/xlsx'
import { classifyFormatCode, formatSerial } from '../src/xlsx-dates'
it.each(['[h]:mm:ss', '[mm]:ss', '[ss]'])('refuses unrepresentable durations in %s', (code) => {
  const parts = classifyFormatCode(code)!
  expect(formatSerial(Number.MAX_VALUE, parts, false)).toBeNull()
  expect(formatSerial(1.5, parts, false)).not.toMatch(/NaN|Infinity/)
})
it('retains the original numeric cell when elapsed formatting overflows', async () => {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    '<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Durations" r:id="s"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships><Relationship Id="s" Target="sheet.xml"/></Relationships>',
  )
  zip.file('xl/styles.xml', '<styleSheet><cellXfs><xf numFmtId="46"/></cellXfs></styleSheet>')
  zip.file(
    'xl/sheet.xml',
    '<worksheet><sheetData><row><c r="A1"><v>1e308</v></c><c r="B1"><v>1.5</v></c></row></sheetData></worksheet>',
  )
  expect(await xlsxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
    '# Durations\n1e308 | 36:00:00',
  )
})
