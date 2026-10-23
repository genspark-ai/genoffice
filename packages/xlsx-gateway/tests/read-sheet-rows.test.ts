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

describe('readSheetRows', () => {
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
