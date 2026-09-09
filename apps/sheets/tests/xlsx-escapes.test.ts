import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, readBasicWorkbook } from '../src/gateway/xlsx-gateway'
import { decodeXlsxEscapes, encodeXlsxEscapes } from '../src/gateway/xlsx-escapes'

describe('_xHHHH_ cell text escapes (ECMA-376 22.4.2.4)', () => {
  it('decodes CR, TAB and the _x005F_ literal marker; leaves non-matches alone', () => {
    expect(decodeXlsxEscapes('a_x000D_b')).toBe('a\rb')
    expect(decodeXlsxEscapes('a_x0009_b')).toBe('a\tb')
    expect(decodeXlsxEscapes('_x005F_x000D_')).toBe('_x000D_')
    expect(decodeXlsxEscapes('_xZZZZ_ _x00D_')).toBe('_xZZZZ_ _x00D_')
  })

  it('encodes control characters and protects literal _x sequences', () => {
    expect(encodeXlsxEscapes('a\rb\u0001c\td')).toBe('a_x000D_b_x0001_c\td')
    expect(encodeXlsxEscapes('_x000D_')).toBe('_x005F_x000D_')
    expect(decodeXlsxEscapes(encodeXlsxEscapes('_x000D_\r_x005F_'))).toBe('_x000D_\r_x005F_')
  })
})

const PARTS: Record<string, string> = {
  '[Content_Types].xml':
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
    '</Types>',
  '_rels/.rels':
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>',
  'xl/workbook.xml':
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="S" sheetId="1" r:id="rId1"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels':
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
    '</Relationships>',
  'xl/sharedStrings.xml':
    '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<si><t>E151 _x000D_\r\nLUSAKA</t></si><si><t>KITWE_x000D_\r\nZM</t></si></sst>',
  'xl/worksheets/sheet1.xml':
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c><c r="C1"><f>A1</f><v>E151 </v></c></row></sheetData></worksheet>',
}

async function buildXlsx(): Promise<Buffer> {
  const zip = new JSZip()
  for (const [name, xml] of Object.entries(PARTS)) zip.file(name, xml)
  return zip.generateAsync({ type: 'nodebuffer' })
}

describe('escaped shared strings through the gateway', () => {
  it('reads Excel CR LF as one line break', async () => {
    const imported = await readBasicWorkbook(await buildXlsx())
    expect(imported.snapshot.sheets[0]?.cells.A1?.value).toBe('E151 \nLUSAKA')
  })

  it('re-escapes an edited CR and keeps untouched strings byte-identical', async () => {
    const mutation = await applyCellEditsToXlsx(await buildXlsx(), [
      { sheetName: 'S', row: 0, column: 0, writeValue: true, cell: { value: 'A\rB\u0007_x0041_' } },
    ])
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = await zip.file('xl/worksheets/sheet1.xml')?.async('text')
    expect(sheet).toContain('<t xml:space="preserve">A_x000D_B_x0007__x005F_x0041_</t>')
    expect(await zip.file('xl/sharedStrings.xml')?.async('text')).toBe(
      PARTS['xl/sharedStrings.xml'],
    )
    const reread = await readBasicWorkbook(mutation.buffer)
    expect(reread.snapshot.sheets[0]?.cells.A1?.value).toBe('A\nB\u0007_x0041_')
    expect(reread.snapshot.sheets[0]?.cells.B1?.value).toBe('KITWE\nZM')
  })
})

describe('t="str" cached formula results', () => {
  it('encodes a recalculated text result like cell text so the next read decodes it', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildXlsx(),
      [{ sheetName: 'S', row: 0, column: 0, writeValue: true, cell: { value: 'x' } }],
      [],
      [],
      undefined,
      [],
      [],
      [],
      [],
      [],
      null,
      [],
      [],
      [{ sheetName: 'S', cells: [{ row: 0, column: 2, value: 'A\rB_x000D_' }] }],
    )
    const zip = await JSZip.loadAsync(mutation.buffer)
    const sheet = (await zip.file('xl/worksheets/sheet1.xml')?.async('text')) ?? ''
    const cached = /<c r="C1"\s+t="str"><f>A1<\/f><v>([^<]*)<\/v><\/c>/.exec(sheet)?.[1]
    expect(cached).toBe('A_x000D_B_x005F_x000D_')
    expect(decodeXlsxEscapes(cached ?? '')).toBe('A\rB_x000D_')
  })
})
