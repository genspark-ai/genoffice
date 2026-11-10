import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  createBufferEntrySource,
  planCellEditsToXlsx,
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

const WORKSHEET = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Region</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Product</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Amount</t></is></c>
    </row>
    <row r="2"><c r="A2" t="inlineStr"><is><t>East</t></is></c><c r="B2" t="inlineStr"><is><t>A</t></is></c><c r="C2"><v>100</v></c></row>
    <row r="3"><c r="A3" t="inlineStr"><is><t>West</t></is></c><c r="B3" t="inlineStr"><is><t>B</t></is></c><c r="C3"><v>50</v></c></row>
    <row r="4"><c r="A4" t="inlineStr"><is><t>East</t></is></c><c r="B4" t="inlineStr"><is><t>B</t></is></c><c r="C4"><v>50</v></c></row>
  </sheetData>
</worksheet>`

async function fixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file('[Content_Types].xml', CONTENT_TYPES)
  zip.file('_rels/.rels', PACKAGE_RELS)
  zip.file('xl/workbook.xml', WORKBOOK)
  zip.file('xl/_rels/workbook.xml.rels', WORKBOOK_RELS)
  zip.file('xl/styles.xml', STYLES)
  zip.file('xl/worksheets/sheet1.xml', WORKSHEET)
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

// Two value fields: the writer emits `<i i="1">` in <colItems>, so the guard
// must also accept the i attribute the builders themselves write.
const VALUES = [
  { fieldIndex: 2, agg: 'sum' },
  { fieldIndex: 2, agg: 'count' },
] as const

const SOURCE_AREA = { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 }
const LOCATION = { startRow: 0, startColumn: 5, endRow: 3, endColumn: 7 }

const ADDITION: SheetPivotAddition = {
  sheetName: 'Data',
  sourceSheetName: 'Data',
  sourceArea: SOURCE_AREA,
  location: LOCATION,
  name: 'PivotData',
  fieldNames: ['Region', 'Product', 'Amount'],
  rowFieldIndices: [0],
  rowItems: ['East', 'West'],
  values: [...VALUES],
}

function relayout(cachePath: string, rowField: number, rowItems: string[]): PivotRefreshUpdate {
  return {
    cachePath,
    sheetName: 'Data',
    newOutputRef: 'F1:H4',
    relayout: {
      sourceSheetName: 'Data',
      sourceArea: SOURCE_AREA,
      location: LOCATION,
      name: 'placeholder',
      fieldNames: ['Region', 'Product', 'Amount'],
      rowFieldIndices: [rowField],
      rowItems,
      values: [...VALUES],
    },
  }
}

async function planRelayout(buf: Buffer, update: PivotRefreshUpdate): Promise<MutationPlan> {
  const source = await createBufferEntrySource(buf)
  return planCellEditsToXlsx(
    source,
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

describe('relayout round trip over a pivot the gateway built', () => {
  it('accepts a second layout edit after the first save set refreshOnLoad', async () => {
    const initial = await fixture()
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

    // First layout edit: rows switch from Region to Product; the same save
    // flags the cache with refreshOnLoad="1".
    const firstPlan = await planRelayout(withPivot, relayout(cachePath!, 1, ['A', 'B']))
    const firstCache = firstPlan.replaced.get(cachePath!)!
    expect(firstCache).toContain('refreshOnLoad="1"')
    const tablePath = [...firstPlan.replaced.keys()].find((path) =>
      /^xl\/pivotTables\//.test(path),
    )!
    expect(firstPlan.replaced.get(tablePath)).toContain('<i i="1">')
    const afterFirst = await materialize(withPivot, firstPlan)

    // Second layout edit after "reopen": the parts carry only content the
    // gateway wrote itself, so the fail-closed guard must let it through.
    const secondPlan = await planRelayout(afterFirst, relayout(cachePath!, 0, ['East', 'West']))
    expect(secondPlan.replaced.get(tablePath)).toContain(
      '<rowFields count="1"><field x="0"/></rowFields>',
    )
    expect(secondPlan.replaced.get(cachePath!)).toContain('refreshOnLoad="1"')
  })
})
