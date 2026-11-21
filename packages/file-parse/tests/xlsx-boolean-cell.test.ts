import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { xlsxToText } from '../src/xlsx'
import { buildXlsxFixture } from './helpers/fixtures'

async function extract(row: string): Promise<string> {
  const zip = await JSZip.loadAsync(await buildXlsxFixture())
  zip.file('xl/worksheets/sheet1.xml', `<worksheet><sheetData>${row}</sheetData></worksheet>`)
  return xlsxToText(await zip.generateAsync({ type: 'uint8array' }))
}

it('reads xsd:boolean literals in boolean cells as TRUE/FALSE', async () => {
  const row =
    '<row r="1">' +
    '<c r="A1" t="b"><v>1</v></c>' +
    '<c r="B1" t="b"><v>0</v></c>' +
    '<c r="C1" t="b"><v>true</v></c>' +
    '<c r="D1" t="b"><v>false</v></c>' +
    '</row>'
  expect(await extract(row)).toBe('# Grades\nTRUE | FALSE | TRUE | FALSE')
})

it('accepts uppercase and padded boolean literals', async () => {
  const row =
    '<row r="1">' +
    '<c r="A1" t="b"><v xml:space="preserve"> TRUE </v></c>' +
    '<c r="B1" t="b"><v>FALSE</v></c>' +
    '</row>'
  expect(await extract(row)).toBe('# Grades\nTRUE | FALSE')
})
