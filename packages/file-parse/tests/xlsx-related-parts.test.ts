import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { xlsxToText } from '../src/xlsx'

const S = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

async function workbook(target: string, part: string, legacyDecoys = false): Promise<Uint8Array> {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    `<workbook xmlns="${S}" xmlns:r="${R}"><sheets><sheet name="Records" sheetId="1" r:id="sheet"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
      `<Relationship Id="sheet" Type="${R}/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="strings" Type="${R}/sharedStrings" Target="${target}strings.xml"/>` +
      `<Relationship Id="styles" Type="${R}/styles" Target="${target}styles.xml"/></Relationships>`,
  )
  zip.file(
    `${part}strings.xml`,
    `<sst xmlns="${S}" count="1" uniqueCount="1"><si><t>Invoice 0042</t></si></sst>`,
  )
  zip.file(
    `${part}styles.xml`,
    `<styleSheet xmlns="${S}"><cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="14"/></cellXfs></styleSheet>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<worksheet xmlns="${S}"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><v>45292</v></c><c r="C1"><v>42</v></c></row></sheetData></worksheet>`,
  )
  if (legacyDecoys) {
    zip.file('xl/sharedStrings.xml', `<sst xmlns="${S}"><si><t>Unreferenced text</t></si></sst>`)
    zip.file(
      'xl/styles.xml',
      `<styleSheet xmlns="${S}"><cellXfs><xf numFmtId="0"/><xf numFmtId="0"/></cellXfs></styleSheet>`,
    )
  }
  return zip.generateAsync({ type: 'uint8array' })
}

describe('xlsx workbook metadata relationships', () => {
  it('does not read external metadata targets as local ZIP parts', async () => {
    const zip = await JSZip.loadAsync(await workbook('metadata/', 'xl/metadata/'))
    const rels = await zip.file('xl/_rels/workbook.xml.rels')!.async('text')
    zip.file(
      'xl/_rels/workbook.xml.rels',
      rels.replaceAll('Target="metadata/', 'TargetMode="External" Target="metadata/'),
    )
    expect(await xlsxToText(await zip.generateAsync({ type: 'uint8array' }))).toBe(
      '# Records\n | 45292 | 42',
    )
  })

  it.each([
    ['metadata/', 'xl/metadata/'],
    ['/custom/', 'custom/'],
    ['metadata%20parts/', 'xl/metadata parts/'],
  ])('reads shared text and date formats from %s', async (target, part) => {
    expect(await xlsxToText(await workbook(target, part))).toBe(
      '# Records\nInvoice 0042 | 2024-01-01 | 42',
    )
  })

  it('uses the related parts instead of unreferenced conventional filenames', async () => {
    expect(await xlsxToText(await workbook('metadata/', 'xl/metadata/', true))).toBe(
      '# Records\nInvoice 0042 | 2024-01-01 | 42',
    )
  })
})
