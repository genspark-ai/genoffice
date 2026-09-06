/// genoffice#196: a producer wrote every package entry with a leading '/'.
/// Excel and Numbers open such files; the gateway must fold the quirk
/// instead of rejecting the workbook as unsafe.
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, inventoryXlsx, readBasicWorkbook } from '../src/gateway/xlsx-gateway'

const PARTS: Record<string, string> = {
  '[Content_Types].xml':
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
    '<Default Extension="xml" ContentType="application/xml"/>' +
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
    '</Types>',
  '_rels/.rels':
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
    '</Relationships>',
  'xl/workbook.xml':
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">' +
    '<sheets><sheet name="Codes" sheetId="1" r:id="rId1"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels':
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
    '</Relationships>',
  'xl/worksheets/sheet1.xml':
    '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/></sheetView></sheetViews>' +
    '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>',
}

async function buildXlsx(rename: (name: string) => string): Promise<Buffer> {
  const zip = new JSZip()
  for (const [name, xml] of Object.entries(PARTS)) zip.file(rename(name), xml)
  return zip.generateAsync({ type: 'nodebuffer' })
}

describe('non-conformant ZIP entry names (genoffice#196)', () => {
  it('reads a package whose entries carry a leading slash', async () => {
    const imported = await readBasicWorkbook(await buildXlsx((name) => `/${name}`))
    expect(Object.values(imported.sheetNamesById)).toEqual(['Codes'])
    expect(Object.keys(imported.snapshot.sheets[0]?.cells ?? {})).toEqual(['A1'])
  })

  it('reads a package with backslash separators', async () => {
    const imported = await readBasicWorkbook(await buildXlsx((name) => name.replaceAll('/', '\\')))
    expect(Object.values(imported.sheetNamesById)).toEqual(['Codes'])
  })

  it('saves the package back with conformant names', async () => {
    const source = await buildXlsx((name) => `/${name}`)
    const mutation = await applyCellEditsToXlsx(source, [
      { sheetName: 'Codes', row: 0, column: 0, writeValue: true, cell: { value: 'World' } },
    ])
    const names = (await inventoryXlsx(mutation.buffer)).map((entry) => entry.path).sort()
    expect(names).toEqual(Object.keys(PARTS).sort())
    const zip = await JSZip.loadAsync(mutation.buffer)
    expect(await zip.file('xl/worksheets/sheet1.xml')?.async('text')).toContain('World')
  })
})
