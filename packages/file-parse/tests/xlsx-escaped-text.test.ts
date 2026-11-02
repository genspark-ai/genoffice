import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { xlsxToText } from '../src/xlsx'
import { buildXlsxFixture } from './helpers/fixtures'

it('decodes shared, rich, inline and cached formula strings exactly once', async () => {
  const zip = await JSZip.loadAsync(await buildXlsxFixture())
  zip.file(
    'xl/sharedStrings.xml',
    '<sst>' +
      '<si><t>First_x000D_Second</t></si>' +
      '<si><r><t>Code_x0009_</t></r><r><t>0042</t></r></si>' +
      '<si><t>_x005F_x000D_</t></si></sst>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet><sheetData><row>' +
      '<c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c>' +
      '<c r="C1" t="s"><v>2</v></c>' +
      '<c r="D1" t="inlineStr"><is><t>A_x000A_B</t></is></c>' +
      '<c r="E1" t="str"><f>CHAR(9)</f><v>_x0009_</v></c>' +
      '</row></sheetData></worksheet>',
  )
  const bytes = await zip.generateAsync({ type: 'uint8array' })
  const original = bytes.slice()
  expect(await xlsxToText(bytes)).toBe('# Grades\nFirst\rSecond | Code\t0042 | _x000D_ | A\nB | \t')
  expect(bytes).toEqual(original)
})

it('keeps formula expressions and incomplete escape-like text unchanged', async () => {
  const zip = await JSZip.loadAsync(await buildXlsxFixture())
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<worksheet><sheetData><row>' +
      '<c r="A1"><f>&quot;_x000D_&quot;</f></c>' +
      '<c r="B1" t="inlineStr"><is><t>_xZZZZ_ _x00D_ 0042 &amp; more</t></is></c>' +
      '</row></sheetData></worksheet>',
  )
  expect(await xlsxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
    '# Grades\n="_x000D_" | _xZZZZ_ _x00D_ 0042 & more',
  )
})
