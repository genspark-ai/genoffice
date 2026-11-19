import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  createBufferEntrySource,
  planCellEditsToXlsx,
  PivotExpandError,
  type MutationPlan,
  type PivotRefreshUpdate,
  type SheetPivotAddition,
} from '../src/gateway/xlsx-gateway'

const CONTENT_TYPES = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`

const STYLES = `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>
</styleSheet>`

const SOURCE_ROWS = `    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Region</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Amount</t></is></c>
    </row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>East</t></is></c><c r="B2"><v>100</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>West</t></is></c><c r="B3"><v>50</v></c></row>`

function worksheet(extraRows: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
${SOURCE_ROWS}
${extraRows}
  </sheetData>
</worksheet>`
}

async function fixture(extraRows: string): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', PACKAGE_RELS)
  zip.file('xl/workbook.xml', WORKBOOK)
  zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
  zip.file('xl/styles.xml', STYLES)
  zip.file('xl/worksheets/sheet1.xml', worksheet(extraRows))
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function materialize(buf: Buffer, plan: MutationPlan): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buf)
  for (const path of plan.removedEntries) zip.remove(path)
  for (const [path, xml] of plan.replaced) zip.file(path, xml)
  for (const [path, xml] of plan.added) zip.file(path, xml)
  for (const [path, bytes] of plan.addedBinary) zip.file(path, bytes)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

const SOURCE_AREA = { startRow: 0, startColumn: 0, endRow: 2, endColumn: 1 }
const LOCATION = { startRow: 0, startColumn: 5, endRow: 2, endColumn: 7 }

const ADDITION: SheetPivotAddition = {
  sheetName: 'Data',
  sourceSheetName: 'Data',
  sourceArea: SOURCE_AREA,
  location: LOCATION,
  name: 'PivotData',
  fieldNames: ['Region', 'Amount'],
  rowFieldIndices: [0],
  rowItems: ['East', 'West'],
  values: [{ fieldIndex: 1, agg: 'sum' }],
}

/// Build a package holding a gateway-written pivot at F1:H3, then ask for an
/// expansion to F1:H4 — the only added area is row 4, columns F to H.
async function expandIntoRow4(extraRows: string): Promise<MutationPlan> {
  const initial = await fixture(extraRows)
  const addPlan = await planCellEditsToXlsx(
    await createBufferEntrySource(initial),
    [],
    [],
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    [],
    [ADDITION],
  )
  const cachePath = [...addPlan.added.keys()].find((path) =>
    /^xl\/pivotCache\/pivotCacheDefinition[^/]*\.xml$/.test(path),
  )
  expect(cachePath).toBeDefined()
  const withPivot = await materialize(initial, addPlan)
  const update: PivotRefreshUpdate = {
    cachePath: cachePath!,
    sheetName: 'Data',
    newOutputRef: 'F1:H4',
  }
  return planCellEditsToXlsx(
    await createBufferEntrySource(withPivot),
    [],
    [],
    [],
    undefined,
    [],
    [],
    [],
    [],
    [],
    null,
    [],
    [],
    [],
    [],
    /* pivotAdditions */ [],
    /* pivotCacheRefreshPaths */ [update.cachePath],
    [update],
  )
}

describe('pivot expansion conflict scan', () => {
  it('treats a self-closing <c/> as empty instead of borrowing the next cell', async () => {
    // F4 is styled but empty; I4 holds a value but sits outside the added area
    // F4:H4. The expansion must go through, and F1:H4 must land in the table.
    const plan = await expandIntoRow4(
      '    <row r="4"><c r="F4" s="1"/><c r="I4"><v>5</v></c></row>',
    )
    const tablePath = [...plan.replaced.keys()].find((path) => /^xl\/pivotTables\//.test(path))!
    expect(plan.replaced.get(tablePath)).toContain('ref="F1:H4"')

    // A value genuinely inside the added area must still be refused.
    await expect(
      expandIntoRow4('    <row r="4"><c r="F4" s="1"/><c r="G4"><v>5</v></c></row>'),
    ).rejects.toThrow(PivotExpandError)
  })
})
