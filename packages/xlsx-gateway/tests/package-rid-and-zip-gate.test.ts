import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  createBufferEntrySource,
  inventoryXlsx,
  planCellEditsToXlsx,
  readBasicWorkbook,
} from '../src/gateway/xlsx-gateway'

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`

const PACKAGE_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`

const WORKBOOK = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`

const WORKBOOK_RELS = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`

const WORKSHEET = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:A3"/>
  <sheetData>
    <row r="1"><c r="A1"><v>3</v></c></row>
    <row r="2"><c r="A2"><v>1</v></c></row>
    <row r="3"><c r="A3"><v>2</v></c></row>
  </sheetData>
</worksheet>`

async function fixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', PACKAGE_RELS)
  zip.file('xl/workbook.xml', WORKBOOK)
  zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
  zip.file('xl/worksheets/sheet1.xml', WORKSHEET)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

function relationshipIds(relsXml: string): string[] {
  return [...relsXml.matchAll(/\bId="(rId\d+)"/g)].map((match) => match[1]!)
}

describe('workbook relationship ids stay unique across one save', () => {
  it('adds a sheet together with the first dynamic-array formula and stylesheet', async () => {
    const source = await createBufferEntrySource(await fixture())
    const plan = await planCellEditsToXlsx(
      source,
      [
        {
          sheetName: 'Data',
          row: 0,
          column: 1,
          writeValue: true,
          cell: { value: 1, formula: '=SORT(A1:A3)' },
          style: { bold: true },
        },
      ],
      [],
      [],
      { renames: [], additions: [{ name: 'Extra' }], removals: [], order: ['Data', 'Extra'] },
    )
    const rels = plan.replaced.get('xl/_rels/workbook.xml.rels')
    expect(rels).toBeDefined()
    const ids = relationshipIds(rels!)
    expect(new Set(ids).size).toBe(ids.length)
    expect(rels).toContain('Target="worksheets/sheet2.xml"')
    expect(rels).toContain('Target="metadata.xml"')
    expect(rels).toContain('Target="styles.xml"')
    expect(plan.added.has('xl/worksheets/sheet2.xml')).toBe(true)
  })
})

/** A valid archive whose central directory lies about one part's inflated size. */
async function forgedDeclaredSize(declares: number): Promise<Buffer> {
  const bytes = await fixture()
  const name = 'xl/worksheets/sheet1.xml'
  const centralName = bytes.lastIndexOf(Buffer.from(name))
  bytes.writeUInt32LE(declares, centralName - 46 + 24)
  return bytes
}

describe('zip inflation gate', () => {
  it('rejects a central directory that declares a huge part before inflating it', async () => {
    const bomb = await forgedDeclaredSize(0xffffff00)
    await expect(inventoryXlsx(bomb)).rejects.toThrow(
      /zip rejected: part xl\/worksheets\/sheet1\.xml declares/,
    )
    await expect(readBasicWorkbook(bomb)).rejects.toThrow(/zip rejected/)
    await expect(createBufferEntrySource(bomb)).rejects.toThrow(/zip rejected/)
  })

  it('rejects a part that inflates past what it declares', async () => {
    await expect(inventoryXlsx(await forgedDeclaredSize(8))).rejects.toThrow(/zip rejected/)
  })

  it('still accepts an honest workbook', async () => {
    const entries = await inventoryXlsx(await fixture())
    expect(entries.map((entry) => entry.path)).toContain('xl/worksheets/sheet1.xml')
  })
})
