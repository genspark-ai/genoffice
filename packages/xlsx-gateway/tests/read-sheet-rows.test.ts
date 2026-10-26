import { describe, expect, it } from 'vitest'
import JSZip from 'jszip'
import { csvToXlsxBuffer } from '../src/gateway/csv-import'
import { readSheetRows } from '../src/gateway/xlsx-gateway'

/**
 * Reading a worksheet as plain text rows: the reader a data source (a mail
 * merge recipient list) needs — the values as the file stores them, not the
 * workbook's editable model. genoffice#2097.
 */

/** a workbook whose cells live in the shared string table, as Excel writes them */
async function sharedStringWorkbook(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Recipients" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  )
  zip.file(
    'xl/sharedStrings.xml',
    '<?xml version="1.0" encoding="UTF-8"?><sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="2"><si><t>Name</t></si><si><t>Ada</t></si></sst>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c></row><row r="2"><c r="A2" t="s"><v>1</v></c></row></sheetData></worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

/** a workbook with a date-styled cell, a time-styled cell and a plain number */
async function datedWorkbook(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    'xl/workbook.xml',
    '<?xml version="1.0" encoding="UTF-8"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets><sheet name="Dates" sheetId="1" r:id="rId1" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"/></sheets></workbook>',
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
  )
  zip.file(
    'xl/styles.xml',
    '<?xml version="1.0" encoding="UTF-8"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
      '<numFmts count="2"><numFmt numFmtId="164" formatCode="yyyy&quot;\u5e74&quot;m&quot;\u6708&quot;d&quot;\u65e5&quot;"/><numFmt numFmtId="165" formatCode="#,##0.00"/></numFmts>' +
      '<cellXfs count="4"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="164"/><xf numFmtId="165"/><xf numFmtId="20"/></cellXfs></styleSheet>',
  )
  zip.file(
    'xl/worksheets/sheet1.xml',
    '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData>' +
      '<row r="1"><c r="A1" t="inlineStr"><is><t>Joined</t></is></c><c r="B1" t="inlineStr"><is><t>Custom</t></is></c><c r="C1" t="inlineStr"><is><t>Amount</t></is></c><c r="D1" t="inlineStr"><is><t>At</t></is></c></row>' +
      '<row r="2"><c r="A2" s="1" t="n"><v>45930</v></c><c r="B2" s="2"><v>45930.5</v></c><c r="C2" s="3"><v>45930</v></c><c r="D2" s="4"><v>0.5</v></c></row>' +
      '</sheetData></worksheet>',
  )
  return zip.generateAsync({ type: 'nodebuffer' })
}

describe('readSheetRows', () => {
  it('reads a date-styled number as a date and leaves other numbers alone', async () => {
    const rows = await readSheetRows(await datedWorkbook())
    expect(rows[1]).toEqual(['2025-09-30', '2025-09-30 12:00:00', '45930', '12:00:00'])
  })

  it('reads the first sheet as rows of text', async () => {
    const buffer = await csvToXlsxBuffer('Name,Age\nAda,36\nGrace,45\n')
    expect(await readSheetRows(buffer)).toEqual([
      ['Name', 'Age'],
      ['Ada', '36'],
      ['Grace', '45'],
    ])
  })

  it('reads a cell that lives in the shared string table', async () => {
    expect(await readSheetRows(await sharedStringWorkbook())).toEqual([['Name'], ['Ada']])
  })

  it('keeps a cell that holds a comma inside a quoted field', async () => {
    const buffer = await csvToXlsxBuffer('Name,Note\n"Ada","a, b"\n')
    expect(await readSheetRows(buffer)).toEqual([
      ['Name', 'Note'],
      ['Ada', 'a, b'],
    ])
  })

  it('reads a sheet that is not the first one when it is named', async () => {
    const buffer = await csvToXlsxBuffer('Name\nAda\n', 'Recipients')
    expect(await readSheetRows(buffer, 'Recipients')).toEqual([['Name'], ['Ada']])
  })

  it('reports a workbook without a readable worksheet', async () => {
    const empty = await csvToXlsxBuffer('Name\nAda\n')
    await expect(readSheetRows(empty, 'Missing')).rejects.toThrow(/Missing/)
  })
})
