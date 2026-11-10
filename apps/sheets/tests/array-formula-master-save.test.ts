/**
 * Legacy CSE array masters on the streamed save path: editing a master whose
 * <f t="array" ref="…"> carries no spilling function must keep the array type
 * and extent, or the master degrades to an ordinary formula while its followers
 * keep stale cached values.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, type CellEdit } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'

/// Master at A1 spans A1:C3; Excel writes the followers as cached values only.
const CSE_MASTER = '<c r="A1" s="5"><f t="array" ref="A1:C3">SUM(A2:A3*B2:B3)</f><v>1</v></c>'

async function buildWorksheet(sheetData: string, dimension = 'A1:C3'): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="${dimension}"/>
  <sheetData>${sheetData}</sheetData>
</worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function saveWorksheet(source: Buffer, edits: readonly CellEdit[]): Promise<string> {
  const mutation = await applyCellEditsToXlsx(source, edits)
  const zip = await JSZip.loadAsync(mutation.buffer)
  const worksheet = await zip.file('xl/worksheets/sheet1.xml')?.async('string')
  expect(worksheet).toBeDefined()
  return worksheet!
}

function editMaster(formula: string): CellEdit {
  return {
    sheetName: 'Data',
    row: 0,
    column: 0,
    writeValue: true,
    cell: { value: null, formula },
  }
}

const CSE_SHEET_DATA =
  `<row r="1">${CSE_MASTER}<c r="B1"><v>2</v></c><c r="C1"><v>3</v></c></row>` +
  '<row r="2"><c r="A2"><v>1</v></c><c r="B2"><v>2</v></c><c r="C2"><v>3</v></c></row>' +
  '<row r="3"><c r="A3"><v>4</v></c><c r="B3"><v>5</v></c><c r="C3"><v>6</v></c></row>'

describe('legacy CSE array master on the streamed save path', () => {
  it('keeps t="array" and the CSE extent when the master is edited', async () => {
    const source = await buildWorksheet(CSE_SHEET_DATA)
    const worksheet = await saveWorksheet(source, [editMaster('=SUM(A2:A3*B2:B3)*2')])
    expect(worksheet).toContain(
      '<c r="A1" s="5"><f t="array" ref="A1:C3">SUM(A2:A3*B2:B3)*2</f></c>',
    )
  })

  it('leaves the followers of an edited master untouched', async () => {
    const source = await buildWorksheet(CSE_SHEET_DATA)
    const worksheet = await saveWorksheet(source, [editMaster('=SUM(A2:A3*B2:B3)*2')])
    expect(worksheet).toContain('<c r="B1"><v>2</v></c>')
    expect(worksheet).toContain('<c r="C1"><v>3</v></c>')
    expect(worksheet).toContain('<c r="C3"><v>6</v></c>')
  })

  it('keeps a single-cell CSE extent', async () => {
    const source = await buildWorksheet(
      `<row r="1"><c r="A1"><f t="array" ref="A1">SUM(B1:B3)</f><v>6</v></c></row>` +
        '<row r="2"><c r="B2"><v>2</v></c></row>',
      'A1:B3',
    )
    const worksheet = await saveWorksheet(source, [editMaster('=SUM(B1:B3)+1')])
    expect(worksheet).toContain('<f t="array" ref="A1">SUM(B1:B3)+1</f>')
  })

  it('keeps the full spill extent of an existing dynamic array master', async () => {
    // The spill heuristic would collapse the extent to the master's own address.
    const source = await buildWorksheet(
      `<row r="1"><c r="A1"><f t="array" ref="A1:A3">_xlfn._xlws.FILTER(B1:B3,C1:C3)</f><v>1</v></c></row>` +
        '<row r="2"><c r="A2"><v>2</v></c></row>' +
        '<row r="3"><c r="A3"><v>3</v></c></row>',
    )
    const worksheet = await saveWorksheet(source, [editMaster('=_xlfn._xlws.FILTER(B1:B3,C1:C3)')])
    expect(worksheet).toContain('ref="A1:A3"')
  })

  it('still writes a plain formula for a cell that never held an array', async () => {
    const source = await buildWorksheet(
      '<row r="1"><c r="A1"><f>SUM(B1:B3)</f><v>6</v></c></row><row r="2"><c r="B2"><v>2</v></c></row>',
      'A1:B3',
    )
    const worksheet = await saveWorksheet(source, [editMaster('=SUM(B1:B3)+1')])
    expect(worksheet).toContain('<c r="A1"><f>SUM(B1:B3)+1</f></c>')
    expect(worksheet).not.toContain('t="array"')
  })

  it('falls back to a plain formula when the stored extent is not a grid range', async () => {
    const source = await buildWorksheet(
      '<row r="1"><c r="A1"><f t="array" ref="not-a-range">SUM(B1:B3)</f><v>6</v></c></row>' +
        '<row r="2"><c r="B2"><v>2</v></c></row>',
      'A1:B3',
    )
    const worksheet = await saveWorksheet(source, [editMaster('=SUM(B1:B3)+1')])
    expect(worksheet).toContain('<c r="A1"><f>SUM(B1:B3)+1</f></c>')
    expect(worksheet).not.toContain('ref="not-a-range"')
  })

  it('does not invent an array master for a cell that had none', async () => {
    const source = await buildWorksheet(
      '<row r="1"><c r="A1"><v>4</v></c></row><row r="2"><c r="B2"><v>2</v></c></row>',
    )
    const worksheet = await saveWorksheet(source, [editMaster('=SUM(B2:B3)')])
    expect(worksheet).toContain('<c r="A1"><f>SUM(B2:B3)</f></c>')
    expect(worksheet).not.toContain('t="array"')
  })
})
