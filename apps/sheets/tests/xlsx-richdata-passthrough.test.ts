import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  applyCellEditsToXlsx,
  assertOnlyTouchedEntriesChanged,
  type CellEdit,
} from '../src/gateway/xlsx-gateway'

const contentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="png" ContentType="image/png"/>
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/metadata.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml"/>
  <Override PartName="/xl/richData/richValueRel.xml" ContentType="application/vnd.ms-excel.richvaluerel+xml"/>
  <Override PartName="/xl/richData/rdrichvalue.xml" ContentType="application/vnd.ms-excel.rdrichvalue+xml"/>
  <Override PartName="/xl/richData/rdrichvaluestructure.xml" ContentType="application/vnd.ms-excel.rdrichvaluestructure+xml"/>
</Types>`

const worksheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData><row r="2"><c r="B2" t="e" vm="1"><v>#VALUE!</v></c></row></sheetData>
</worksheet>`

/// Minimal modern-Excel "picture in cell" package: the cell carries vm= and a
/// cached #VALUE!, the picture lives behind xl/metadata.xml → xl/richData.
async function buildRichDataFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', contentTypes)
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" Target="metadata.xml"/><Relationship Id="rId5" Type="http://schemas.microsoft.com/office/2022/10/relationships/richValueRel" Target="richData/richValueRel.xml"/><Relationship Id="rId6" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValue" Target="richData/rdrichvalue.xml"/><Relationship Id="rId7" Type="http://schemas.microsoft.com/office/2017/06/relationships/rdRichValueStructure" Target="richData/rdrichvaluestructure.xml"/></Relationships>`,
  )
  zip.file('xl/worksheets/sheet1.xml', worksheet)
  zip.file(
    'xl/metadata.xml',
    `<metadata xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xlrd="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata"><metadataTypes count="1"><metadataType name="XLRICHVALUE" minSupportedVersion="120000"/></metadataTypes><futureMetadata name="XLRICHVALUE" count="1"><bk><extLst><ext uri="{3e2802c4-a4d2-4d8b-9148-e3be6c30e623}"><xlrd:rvb i="0"/></ext></extLst></bk></futureMetadata><valueMetadata count="1"><bk><rc t="1" v="0"/></bk></valueMetadata></metadata>`,
  )
  zip.file(
    'xl/richData/richValueRel.xml',
    `<richValueRels xmlns="http://schemas.microsoft.com/office/spreadsheetml/2022/richvaluerel" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><rel r:id="rId1"/></richValueRels>`,
  )
  zip.file(
    'xl/richData/_rels/richValueRel.xml.rels',
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/></Relationships>`,
  )
  zip.file(
    'xl/richData/rdrichvalue.xml',
    `<rvData xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><rv s="0"><v>0</v><v>5</v></rv></rvData>`,
  )
  zip.file(
    'xl/richData/rdrichvaluestructure.xml',
    `<rvStructures xmlns="http://schemas.microsoft.com/office/spreadsheetml/2017/richdata" count="1"><s t="_localImage"><k n="_rvRel:LocalImageIdentifier" t="i"/><k n="CalcOrigin" t="i"/></s></rvStructures>`,
  )
  zip.file('xl/media/image1.png', Buffer.from('fake-png-bytes'))
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

describe('richData passthrough on save', () => {
  it('keeps the metadata, richData, and media parts byte-identical', async () => {
    const edit: CellEdit = {
      sheetName: 'Sheet1',
      row: 0,
      column: 0,
      writeValue: true,
      cell: { value: 'hello' },
    }
    const mutation = await applyCellEditsToXlsx(await buildRichDataFixture(), [edit])
    expect(() => assertOnlyTouchedEntriesChanged(mutation)).not.toThrow()
    const before = new Map(mutation.beforeEntries.map((entry) => [entry.path, entry.sha256]))
    const after = new Map(mutation.afterEntries.map((entry) => [entry.path, entry.sha256]))
    for (const part of [
      'xl/metadata.xml',
      'xl/richData/richValueRel.xml',
      'xl/richData/_rels/richValueRel.xml.rels',
      'xl/richData/rdrichvalue.xml',
      'xl/richData/rdrichvaluestructure.xml',
      'xl/media/image1.png',
    ]) {
      expect(after.get(part), part).toBe(before.get(part))
    }
    // The untouched picture cell keeps its vm= marker and cached error.
    const saved = await JSZip.loadAsync(mutation.buffer)
    const sheet = await saved.file('xl/worksheets/sheet1.xml')?.async('string')
    expect(sheet).toContain('<c r="B2" t="e" vm="1"><v>#VALUE!</v></c>')
  })
})
