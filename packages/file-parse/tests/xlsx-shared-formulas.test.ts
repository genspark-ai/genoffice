import { expect, it } from 'vitest'
import JSZip from 'jszip'
import { xlsxToText } from '../src/xlsx'
import { buildXlsxFixture } from './helpers/fixtures'

async function extract(rows: string): Promise<string> {
  const zip = await JSZip.loadAsync(await buildXlsxFixture())
  zip.file('xl/worksheets/sheet1.xml', `<worksheet><sheetData>${rows}</sheetData></worksheet>`)
  return xlsxToText(await zip.generateAsync({ type: 'uint8array' }))
}

it('translates uncached shared formulas by both row and column offsets', async () => {
  expect(
    await extract(
      '<row r="2"><c r="B2"><f t="shared" si="0" ref="B2:C3">A2+$A2+A$1+$A$1+SUM(A2:A3)+&quot;A2&quot;</f></c></row>' +
        '<row r="3"><c r="C3"><f t="shared" si="0"/></c></row>',
    ),
  ).toBe('# Grades\n | =A2+$A2+A$1+$A$1+SUM(A2:A3)+"A2"\n |  | =B3+$A3+B$1+$A$1+SUM(B3:B4)+"A2"')
})

it('uses cached values but still registers the master for uncached followers', async () => {
  expect(
    await extract(
      '<row r="1"><c r="A1"><f t="shared" si="1" ref="A1:A3">B1*2</f><v>10</v></c></row>' +
        '<row r="2"><c r="A2"><f t="shared" si="1"/><v>20</v></c></row>' +
        '<row r="3"><c r="A3"><f t="shared" si="1"/></c></row>',
    ),
  ).toBe('# Grades\n10\n20\n=B3*2')
})

it('resolves a follower even when the master is stored later in the worksheet', async () => {
  expect(
    await extract(
      '<row r="1"><c r="A1"><f t="shared" si="2"/></c></row>' +
        '<row r="2"><c r="A2"><f t="shared" si="2" ref="A1:A2">B2+1</f></c></row>',
    ),
  ).toBe('# Grades\n=B1+1\n=B2+1')
})

it('preserves sheet qualifiers and handles lowercase cell addresses', async () => {
  expect(
    await extract(
      '<row r="1"><c r="a1"><f t="shared" si="3" ref="A1:B1">\'Data Set\'!B1+SUM(B:B)+SUM(1:1)</f></c>' +
        '<c r="b1"><f t="shared" si="3"/></c></row>',
    ),
  ).toBe("# Grades\n='Data Set'!B1+SUM(B:B)+SUM(1:1) | ='Data Set'!C1+SUM(C:C)+SUM(1:1)")
})

it('leaves an orphan follower empty and retains an independent formula', async () => {
  expect(
    await extract(
      '<row><c r="A1"><f t="shared" si="404"/></c><c r="B1"><f>SUM(C1:C3)</f></c></row>',
    ),
  ).toBe('# Grades\n | =SUM(C1:C3)')
})

it('retains the existing formula fallback when cell positions are omitted', async () => {
  expect(
    await extract('<row><c><f t="shared" si="4">C1*2</f></c><c><f t="shared" si="4"/></c></row>'),
  ).toBe('# Grades\n=C1*2 | =C1*2')
})

it('does not discard a worksheet because a shared formula has a malformed cell reference', async () => {
  expect(
    await extract(
      '<row><c r="A1"><f t="shared" si="5">C1*2</f></c><c r="1"><f t="shared" si="5"/></c></row>',
    ),
  ).toBe('# Grades\n=C1*2 | =C1*2')
})
