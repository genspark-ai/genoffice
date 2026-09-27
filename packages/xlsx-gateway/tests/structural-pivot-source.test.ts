import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  createBufferEntrySource,
  planCellEditsToXlsx,
  type SheetStructuralOps,
} from '../src/gateway/xlsx-gateway'
import { StructuralShiftError } from '../src/gateway/xlsx-structure'

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
</Types>`

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const WORKBOOK = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/><sheet name="Report" sheetId="2" r:id="rId2"/><sheet name="Other" sheetId="3" r:id="rId3"/></sheets>
</workbook>`

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet3.xml"/>
</Relationships>`

function worksheetXml(): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetData>
    <row r="1"><c r="A1"><v>1</v></c><c r="B1"><v>2</v></c></row>
    <row r="2"><c r="A2"><v>3</v></c><c r="B2"><v>4</v></c></row>
  </sheetData>
</worksheet>`
}

const CACHE_DEFINITION = `<?xml version="1.0" encoding="UTF-8"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cacheSource type="worksheet"><worksheetSource ref="A1:C100" sheet="Data"/></cacheSource>
</pivotCacheDefinition>`

async function fixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', PACKAGE_RELS)
  zip.file('xl/workbook.xml', WORKBOOK)
  zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
  zip.file('xl/worksheets/sheet1.xml', worksheetXml())
  zip.file('xl/worksheets/sheet2.xml', worksheetXml())
  zip.file('xl/worksheets/sheet3.xml', worksheetXml())
  zip.file('xl/pivotCache/pivotCacheDefinition1.xml', CACHE_DEFINITION)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function plan(ops: SheetStructuralOps[]): Promise<unknown> {
  const source = await createBufferEntrySource(await fixture())
  return planCellEditsToXlsx(source, [], ops)
}

describe('structural edits against a pivot cache source sheet', () => {
  it('refuses a row insert on the sheet the cache reads from', async () => {
    await expect(
      plan([{ sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 0, count: 1 }] }]),
    ).rejects.toThrow(StructuralShiftError)
  })

  it('refuses a column delete on the sheet the cache reads from', async () => {
    await expect(
      plan([{ sheetName: 'Data', ops: [{ kind: 'remove-cols', index: 0, count: 1 }] }]),
    ).rejects.toThrow(StructuralShiftError)
  })

  it('still saves a structural edit on an unrelated sheet', async () => {
    const mutation = await plan([
      { sheetName: 'Other', ops: [{ kind: 'insert-rows', index: 0, count: 1 }] },
    ])
    expect(mutation).toBeDefined()
  })
})

describe('non-shifting edits against a pivot cache source sheet', () => {
  it('saves a row-size change on the sheet the cache reads from', async () => {
    const mutation = await plan([
      { sheetName: 'Data', ops: [{ kind: 'set-row-size', start: 0, end: 0, size: 30 }] },
    ])
    expect(mutation).toBeDefined()
  })

  it('saves a hidden-rows change on the sheet the cache reads from', async () => {
    const mutation = await plan([
      { sheetName: 'Data', ops: [{ kind: 'set-rows-hidden', start: 1, end: 1, hidden: true }] },
    ])
    expect(mutation).toBeDefined()
  })

  it('still refuses when a shifting op rides along with a non-shifting one', async () => {
    await expect(
      plan([
        {
          sheetName: 'Data',
          ops: [
            { kind: 'set-row-size', start: 0, end: 0, size: 30 },
            { kind: 'insert-rows', index: 0, count: 1 },
          ],
        },
      ]),
    ).rejects.toThrow(StructuralShiftError)
  })
})
