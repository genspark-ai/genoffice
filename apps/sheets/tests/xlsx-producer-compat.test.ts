/// Regression tests for genoffice#10: workbook.xml serialization styles that
/// differ from Excel's (attribute order, numeric character references,
/// relationship namespace prefix) are valid OOXML and must not break saves.
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx, readBasicWorkbook } from '../src/gateway/xlsx-gateway'
import { parseSheetElements } from '../src/gateway/xlsx-sheets'

const CONTENT_TYPES =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
  '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
  '</Types>'

const ROOT_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
  '</Relationships>'

const WORKBOOK_RELS =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
  '</Relationships>'

const WORKSHEET =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
  '<sheetData><row r="1"><c r="A1"><v>1</v></c></row></sheetData></worksheet>'

const RELATIONSHIPS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

async function buildXlsx(sheetElement: string, relationshipPrefix = 'r'): Promise<Buffer> {
  const workbookXml =
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
    ` xmlns:${relationshipPrefix}="${RELATIONSHIPS_NS}">` +
    `<sheets>${sheetElement}</sheets></workbook>`
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', ROOT_RELS)
  zip.file('xl/workbook.xml', workbookXml)
  zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
  zip.file('xl/worksheets/sheet1.xml', WORKSHEET)
  return zip.generateAsync({ type: 'nodebuffer' })
}

async function editedWorksheet(buffer: Buffer, sheetName: string): Promise<string> {
  const mutation = await applyCellEditsToXlsx(buffer, [
    { sheetName, row: 0, column: 0, writeValue: true, cell: { value: 'World' } },
  ])
  const zip = await JSZip.loadAsync(mutation.buffer)
  const entry = zip.file('xl/worksheets/sheet1.xml')
  if (!entry) throw new Error('worksheet part missing after save')
  return entry.async('text')
}

describe('workbook.xml producer compatibility (issue #10)', () => {
  it('saves when r:id precedes name in the sheet element', async () => {
    const buffer = await buildXlsx('<sheet sheetId="1" r:id="rId1" name="Inventory"/>')
    const worksheet = await editedWorksheet(buffer, 'Inventory')
    expect(worksheet).toContain('World')
  })

  it('saves when the sheet name uses numeric character references', async () => {
    // "&#x5E93;&#x5B58;" decodes to the two CJK characters below.
    const buffer = await buildXlsx('<sheet name="&#x5E93;&#x5B58;" sheetId="1" r:id="rId1"/>')
    const worksheet = await editedWorksheet(buffer, '\u5E93\u5B58')
    expect(worksheet).toContain('World')
  })

  it('saves when the relationships namespace uses a non-standard prefix', async () => {
    const buffer = await buildXlsx('<sheet name="Inventory" sheetId="1" ns1:id="rId1"/>', 'ns1')
    const worksheet = await editedWorksheet(buffer, 'Inventory')
    expect(worksheet).toContain('World')
  })

  it('still fails closed for a genuinely missing sheet', async () => {
    const buffer = await buildXlsx('<sheet sheetId="1" r:id="rId1" name="Inventory"/>')
    await expect(editedWorksheet(buffer, 'Missing')).rejects.toThrow(
      'Sheet "Missing" was not found in workbook.xml.',
    )
  })

  it('decodes numeric character references in sheet names on read', async () => {
    const buffer = await buildXlsx('<sheet name="&#x5E93;&#x5B58;" sheetId="1" r:id="rId1"/>')
    const imported = await readBasicWorkbook(buffer)
    expect(Object.values(imported.sheetNamesById)).toEqual(['\u5E93\u5B58'])
  })
})

describe('parseSheetElements decoding', () => {
  it('decodes named entities, decimal and hex character references', () => {
    const workbookXml =
      '<workbook><sheets>' +
      '<sheet name="A &amp; B &#233; &#x4E2D;" sheetId="1" r:id="rId1"/>' +
      '</sheets></workbook>'
    expect(parseSheetElements(workbookXml)[0]?.name).toBe('A & B \u00E9 \u4E2D')
  })

  it('reads the relationship id regardless of attribute position', () => {
    const workbookXml =
      '<workbook><sheets><sheet r:id="rId7" sheetId="1" name="S"/></sheets></workbook>'
    expect(parseSheetElements(workbookXml)[0]?.relationshipId).toBe('rId7')
  })
})
