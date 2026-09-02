/**
 * '>' is legal inside XML attribute values, and real sheet names carry it
 * (only \ / ? * [ ] : are forbidden). The [^>]* element scans truncated at
 * the '>' inside the name, breaking the parsed sheet set — saves failed with
 * "The sheet order does not match the final sheet set." while the edits were
 * already applied on screen.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { applyCellEditsToXlsx } from '../src/gateway/xlsx-gateway'
import {
  maxSheetIdInWorkbook,
  parseSheetElements,
  pivotCacheReadsFromSheet,
  renameSheetInPivotCacheSource,
} from '../src/gateway/xlsx-sheets'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const workbook =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">` +
  '<sheets>' +
  '<sheet name="Summary" sheetId="1" r:id="rId1"/>' +
  '<sheet name="2. Output>" sheetId="5" r:id="rId2"/>' +
  '</sheets><calcPr calcId="1"/></workbook>'

const worksheet =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
  `<worksheet xmlns="${MAIN_NS}"><sheetData>` +
  '<row r="1"><c r="A1"><v>1</v></c></row>' +
  '</sheetData></worksheet>'

async function buildFixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '<Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
      '</Types>',
  )
  zip.file(
    '_rels/.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>` +
      '</Relationships>',
  )
  zip.file('xl/workbook.xml', workbook)
  zip.file(
    'xl/_rels/workbook.xml.rels',
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      `<Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>` +
      `<Relationship Id="rId2" Type="${REL_NS}/worksheet" Target="worksheets/sheet2.xml"/>` +
      '</Relationships>',
  )
  zip.file('xl/worksheets/sheet1.xml', worksheet)
  zip.file('xl/worksheets/sheet2.xml', worksheet)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

describe('element scans with ">" inside attribute values', () => {
  it('parseSheetElements sees every sheet', () => {
    const sheets = parseSheetElements(workbook)
    expect(sheets.map((sheet) => sheet.name)).toEqual(['Summary', '2. Output>'])
    expect(sheets.map((sheet) => sheet.relationshipId)).toEqual(['rId1', 'rId2'])
  })

  it('maxSheetIdInWorkbook counts a sheet whose name carries ">"', () => {
    expect(maxSheetIdInWorkbook(workbook)).toBe(5)
  })

  it('worksheetSource scans survive a ">" sheet name', () => {
    const cache =
      '<cacheSource type="worksheet"><worksheetSource sheet="2. Output>" ref="A1:B2"/></cacheSource>'
    expect(pivotCacheReadsFromSheet(cache, '2. output>')).toBe(true)
    expect(renameSheetInPivotCacheSource(cache, '2. Output>', 'Renamed')).toContain(
      'sheet="Renamed"',
    )
  })
})

describe('saving a workbook with a ">" sheet name', () => {
  it('cell edits and a sheet addition save through', async () => {
    const mutation = await applyCellEditsToXlsx(
      await buildFixture(),
      [{ sheetName: '2. Output>', row: 0, column: 0, writeValue: true, cell: { value: 7 } }],
      [],
      [],
      {
        renames: [],
        additions: [{ name: 'Added' }],
        removals: [],
        order: ['Summary', '2. Output>', 'Added'],
      },
    )
    const zip = await JSZip.loadAsync(mutation.buffer)
    const savedWorkbook = (await zip.file('xl/workbook.xml')?.async('text')) ?? ''
    expect(savedWorkbook).toContain('<sheet name="Summary" sheetId="1" r:id="rId1"/>')
    // Existing elements are reused verbatim, raw '>' included.
    expect(savedWorkbook).toContain('<sheet name="2. Output>" sheetId="5" r:id="rId2"/>')
    expect(savedWorkbook).toContain('<sheet name="Added" sheetId="6" r:id="rId3"/>')
    const sheet = (await zip.file('xl/worksheets/sheet2.xml')?.async('text')) ?? ''
    expect(sheet).toContain('<v>7</v>')
  })
})
