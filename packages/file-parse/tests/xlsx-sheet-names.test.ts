import JSZip from 'jszip'
import { expect, it } from 'vitest'
import { xlsxToText } from '../src/xlsx'
it.each([
  [' Budget ', ' Budget '],
  ['A_x0020_B', 'A B'],
  ['_x005F_x0041_', '_x0041_'],
])('preserves worksheet name %s', async (name, expected) => {
  const z = new JSZip()
  z.file(
    'xl/workbook.xml',
    `<workbook xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="${name}" sheetId="1" r:id="s"/></sheets></workbook>`,
  )
  z.file(
    'xl/_rels/workbook.xml.rels',
    '<Relationships><Relationship Id="s" Target="worksheets/sheet1.xml"/></Relationships>',
  )
  z.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet><sheetData><row><c><v>2</v></c></row></sheetData></worksheet>',
  )
  expect(await xlsxToText(await z.generateAsync({ type: 'uint8array' }))).toBe(`# ${expected}\n2`)
})
