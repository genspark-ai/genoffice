import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'
import { readBasicWorkbook } from '../src/gateway/xlsx-gateway'

const WORKBOOK = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Good" sheetId="1" r:id="rId1"/><sheet name="Broken" sheetId="2" r:id="rId9"/></sheets>
</workbook>`

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`

const WORKSHEET = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="1"><c r="A1"><v>5</v></c></row></sheetData>
</worksheet>`

async function fixture(workbook: string): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('xl/workbook.xml', workbook)
  zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
  zip.file('xl/worksheets/sheet1.xml', WORKSHEET)
  return zip.generateAsync({ type: 'nodebuffer' })
}

describe('readBasicWorkbook with a dangling sheet relationship', () => {
  it('skips the unresolvable sheet and keeps the rest', async () => {
    const imported = await readBasicWorkbook(await fixture(WORKBOOK))
    expect(imported.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Good'])
    expect(imported.sheetNamesById).toEqual({ 'sheet-1': 'Good' })
  })

  it('skips a sheet whose relationship points at a missing part', async () => {
    const zip = await JSZip.loadAsync(await fixture(WORKBOOK))
    zip.file(
      'xl/_rels/workbook.xml.rels',
      WORKBOOK_RELS.replace(
        '</Relationships>',
        '<Relationship Id="rId9" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet9.xml"/></Relationships>',
      ),
    )
    const imported = await readBasicWorkbook(await zip.generateAsync({ type: 'nodebuffer' }))
    expect(imported.snapshot.sheets.map((sheet) => sheet.name)).toEqual(['Good'])
  })

  it('still rejects an escaping relationship target', async () => {
    const zip = await JSZip.loadAsync(await fixture(WORKBOOK))
    zip.file(
      'xl/_rels/workbook.xml.rels',
      WORKBOOK_RELS.replace('Target="worksheets/sheet1.xml"', 'Target="../../etc/passwd"'),
    )
    await expect(
      readBasicWorkbook(await zip.generateAsync({ type: 'nodebuffer' })),
    ).rejects.toThrow(/Invalid OPC relationship target/)
  })

  it('still fails when no sheet can be read', async () => {
    const onlyBroken = WORKBOOK.replace('<sheet name="Good" sheetId="1" r:id="rId1"/>', '')
    await expect(readBasicWorkbook(await fixture(onlyBroken))).rejects.toThrow(
      'Workbook contains no readable worksheets.',
    )
  })
})
