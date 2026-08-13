/**
 * Tests for pivot layout growth: when source data gains new row/column items,
 * applyPivotLayoutExpansions() widens the pivotTableDefinition location ref
 * and fails closed if existing (non-pivot) content is in the way.
 */
import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import {
  PivotExpandError,
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '../src/gateway/xlsx-gateway'
import type { PivotRefreshUpdate } from '../src/gateway/xlsx-gateway'
import { parsePivotDefinition } from '../src/gateway/xlsx-pivot'
import { recomputePivotData } from '../src/domain/pivot-engine'

// ─── Minimal workbook fixture with one pivot table ───────────────────────────

/// Build an xlsx buffer that has:
///   - Sheet "Data" with source data at A1:C4 (header + 3 rows)
///   - A pivot table pointing at that source, output at F1:G5
///   - The pivot parts correctly linked via rels
async function buildPivotFixture(
  opts: {
    extraCellAtRow?: number // if set, put non-empty content at G<extraCellAtRow>
  } = {},
): Promise<Buffer> {
  const zip = new JSZip()

  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/pivotTables/pivotTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheDefinition1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"/>
  <Override PartName="/xl/pivotCache/pivotCacheRecords1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
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
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
</Relationships>`,
  )

  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>
</styleSheet>`,
  )

  const extraCell =
    opts.extraCellAtRow !== undefined
      ? `    <row r="${opts.extraCellAtRow}"><c r="G${opts.extraCellAtRow}"><v>999</v></c></row>\n`
      : ''

  // Source: A1:C4 (header row 1, data rows 2-4)
  // Current pivot output: F1:G5 (header + 3 data rows + grand total)
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <sheetData>
    <row r="1">
      <c r="A1" t="inlineStr"><is><t>Region</t></is></c>
      <c r="B1" t="inlineStr"><is><t>Product</t></is></c>
      <c r="C1" t="inlineStr"><is><t>Amount</t></is></c>
      <c r="F1" t="inlineStr"><is><t>Region</t></is></c>
      <c r="G1" t="inlineStr"><is><t>Sum of Amount</t></is></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>East</t></is></c>
      <c r="B2" t="inlineStr"><is><t>A</t></is></c>
      <c r="C2"><v>100</v></c>
      <c r="F2" t="inlineStr"><is><t>East</t></is></c>
      <c r="G2"><v>150</v></c>
    </row>
    <row r="3">
      <c r="A3" t="inlineStr"><is><t>West</t></is></c>
      <c r="B3" t="inlineStr"><is><t>B</t></is></c>
      <c r="C3"><v>50</v></c>
      <c r="F3" t="inlineStr"><is><t>West</t></is></c>
      <c r="G3"><v>50</v></c>
    </row>
    <row r="4">
      <c r="A4" t="inlineStr"><is><t>East</t></is></c>
      <c r="B4" t="inlineStr"><is><t>B</t></is></c>
      <c r="C4"><v>50</v></c>
      <c r="F4" t="inlineStr"><is><t>Grand Total</t></is></c>
      <c r="G4"><v>200</v></c>
    </row>
${extraCell}  </sheetData>
</worksheet>`,
  )

  // Pivot table: current output F1:G4 (header + 2 items + grand total = 4 rows)
  zip.file(
    'xl/pivotTables/pivotTable1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
  name="PivotData" cacheId="1">
  <location ref="F1:G4" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>
  <pivotFields count="3">
    <pivotField axis="axisRow" showAll="0">
      <items count="3"><item x="0"/><item x="1"/><item t="default"/></items>
    </pivotField>
    <pivotField dataField="1" showAll="0"/>
    <pivotField dataField="1" showAll="0"/>
  </pivotFields>
  <rowFields count="1"><field x="0"/></rowFields>
  <rowItems count="3"><i><x/></i><i><x v="1"/></i><i t="grand"><x/></i></rowItems>
  <colItems count="1"><i/></colItems>
  <dataFields count="1"><dataField name="Sum of Amount" fld="2" baseField="0" baseItem="0"/></dataFields>
</pivotTableDefinition>`,
  )

  // The pivot table rels: links back to the cache definition
  zip.file(
    'xl/pivotTables/_rels/pivotTable1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" Target="../pivotCache/pivotCacheDefinition1.xml"/>
</Relationships>`,
  )

  // Pivot cache definition
  zip.file(
    'xl/pivotCache/pivotCacheDefinition1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rId1" recordCount="3">
  <cacheSource type="worksheet">
    <worksheetSource ref="A1:C4" sheet="Data"/>
  </cacheSource>
  <cacheFields count="3">
    <cacheField name="Region" numFmtId="0">
      <sharedItems count="2"><s v="East"/><s v="West"/></sharedItems>
    </cacheField>
    <cacheField name="Product" numFmtId="0"><sharedItems/></cacheField>
    <cacheField name="Amount" numFmtId="0"><sharedItems/></cacheField>
  </cacheFields>
</pivotCacheDefinition>`,
  )

  // Cache rels: links to records
  zip.file(
    'xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords" Target="pivotCacheRecords1.xml"/>
</Relationships>`,
  )

  // Pivot cache records
  zip.file(
    'xl/pivotCache/pivotCacheRecords1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" count="3">
  <r><x v="0"/><m/><n v="100"/></r>
  <r><x v="1"/><m/><n v="50"/></r>
  <r><x v="0"/><m/><n v="50"/></r>
</pivotCacheRecords>`,
  )

  // Sheet rels: pivot table relationship
  zip.file(
    'xl/worksheets/_rels/sheet1.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" Target="../pivotTables/pivotTable1.xml"/>
</Relationships>`,
  )

  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function planWithRefreshUpdate(buf: Buffer, updates: PivotRefreshUpdate[]) {
  const source = await createBufferEntrySource(buf)
  // Need at least one of: edits, pivotCacheRefreshPaths, pivotRefreshUpdates
  return planCellEditsToXlsx(
    source,
    /* edits */ [],
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
    /* pivotCacheRefreshPaths */ ['xl/pivotCache/pivotCacheDefinition1.xml'],
    updates,
  )
}

describe('applyPivotLayoutExpansions', () => {
  it('no-op when newOutputRef matches current location ref', async () => {
    const buf = await buildPivotFixture()
    const update: PivotRefreshUpdate = {
      cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
      worksheetPath: 'xl/worksheets/sheet1.xml',
      newOutputRef: 'F1:G4', // same as current
    }
    const plan = await planWithRefreshUpdate(buf, [update])
    // PivotTable xml should NOT be in replaced (no change)
    expect(plan.replaced.has('xl/pivotTables/pivotTable1.xml')).toBe(false)
  })

  it('expands the location ref when new area is larger (one new row)', async () => {
    const buf = await buildPivotFixture()
    // Old: F1:G4 (2 items + grand total = 4 rows)
    // New: F1:G5 (3 items + grand total = 5 rows)
    const update: PivotRefreshUpdate = {
      cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
      worksheetPath: 'xl/worksheets/sheet1.xml',
      newOutputRef: 'F1:G5',
    }
    const plan = await planWithRefreshUpdate(buf, [update])
    const updatedTable = plan.replaced.get('xl/pivotTables/pivotTable1.xml')
    expect(updatedTable).toBeDefined()
    expect(updatedTable).toContain('ref="F1:G5"')
    // Original ref should be gone
    expect(updatedTable).not.toContain('ref="F1:G4"')
  })

  it('expands two extra rows correctly', async () => {
    const buf = await buildPivotFixture()
    const update: PivotRefreshUpdate = {
      cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
      worksheetPath: 'xl/worksheets/sheet1.xml',
      newOutputRef: 'F1:G6', // 2 new rows
    }
    const plan = await planWithRefreshUpdate(buf, [update])
    expect(plan.replaced.get('xl/pivotTables/pivotTable1.xml')).toContain('ref="F1:G6"')
  })

  it('fails closed when new rows conflict with existing non-pivot content', async () => {
    // Row 5 (G5) has content in fixture — new area F1:G5 would overwrite it
    const buf = await buildPivotFixture({ extraCellAtRow: 5 })
    const update: PivotRefreshUpdate = {
      cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
      worksheetPath: 'xl/worksheets/sheet1.xml',
      newOutputRef: 'F1:G5', // G5 is occupied
    }
    await expect(planWithRefreshUpdate(buf, [update])).rejects.toThrow(PivotExpandError)
    await expect(planWithRefreshUpdate(buf, [update])).rejects.toThrow(/conflict/)
  })

  it('does not fail when conflict is outside the new area', async () => {
    // Row 6 (G6) has content, but we only expand to row 5
    const buf = await buildPivotFixture({ extraCellAtRow: 6 })
    const update: PivotRefreshUpdate = {
      cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
      worksheetPath: 'xl/worksheets/sheet1.xml',
      newOutputRef: 'F1:G5', // G6 is outside — should pass
    }
    const plan = await planWithRefreshUpdate(buf, [update])
    expect(plan.replaced.get('xl/pivotTables/pivotTable1.xml')).toContain('ref="F1:G5"')
  })

  it('resolves the worksheet from sheetName when worksheetPath is absent', async () => {
    // The renderer only knows the sheet name (refresh write-back after layout growth takes this path).
    const buf = await buildPivotFixture()
    const update: PivotRefreshUpdate = {
      cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
      sheetName: 'Data',
      newOutputRef: 'F1:G5',
    }
    const plan = await planWithRefreshUpdate(buf, [update])
    expect(plan.replaced.get('xl/pivotTables/pivotTable1.xml')).toContain('ref="F1:G5"')
  })

  it('rejects a moved origin (top-left changed)', async () => {
    const buf = await buildPivotFixture()
    const update: PivotRefreshUpdate = {
      cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
      worksheetPath: 'xl/worksheets/sheet1.xml',
      newOutputRef: 'F2:G5', // start row changed from 1 to 2
    }
    await expect(planWithRefreshUpdate(buf, [update])).rejects.toThrow(PivotExpandError)
  })

  it('silently skips when cache path is not found in rels', async () => {
    const buf = await buildPivotFixture()
    const update: PivotRefreshUpdate = {
      cachePath: 'xl/pivotCache/pivotCacheDefinitionNONEXISTENT.xml',
      worksheetPath: 'xl/worksheets/sheet1.xml',
      newOutputRef: 'F1:G5',
    }
    // Should not throw — just skip
    const plan = await planWithRefreshUpdate(buf, [update])
    expect(plan.replaced.has('xl/pivotTables/pivotTable1.xml')).toBe(false)
  })
})

describe('pivot layout edits (relayout)', () => {
  it('shrinks the location ref when the edited layout got smaller', async () => {
    const buf = await buildPivotFixture()
    const plan = await planWithRefreshUpdate(buf, [
      {
        cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
        worksheetPath: 'xl/worksheets/sheet1.xml',
        newOutputRef: 'F1:G3',
      },
    ])
    expect(plan.replaced.get('xl/pivotTables/pivotTable1.xml')).toContain('ref="F1:G3"')
  })

  it('rewrites definition, cache, and records keeping name/cacheId/rel ids', async () => {
    const buf = await buildPivotFixture()
    const plan = await planWithRefreshUpdate(buf, [
      {
        cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
        sheetName: 'Data',
        newOutputRef: 'F1:G4',
        relayout: {
          sourceSheetName: 'Data',
          sourceArea: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 },
          location: { startRow: 0, startColumn: 5, endRow: 3, endColumn: 6 },
          name: 'PivotEdit', // renderer placeholder — the file name must win
          fieldNames: ['Region', 'Product', 'Amount'],
          rowFieldIndices: [1], // rows switch from Region to Product
          rowItems: ['A', 'B'],
          values: [{ fieldIndex: 2, agg: 'sum' }],
        },
      },
    ])
    const table = plan.replaced.get('xl/pivotTables/pivotTable1.xml')!
    expect(table).toContain('name="PivotData"')
    expect(table).toContain('cacheId="1"')
    expect(table).toContain('<rowFields count="1"><field x="1"/></rowFields>')
    const cache = plan.replaced.get('xl/pivotCache/pivotCacheDefinition1.xml')!
    expect(cache).toContain('r:id="rId1"')
    expect(cache).toContain('<s v="A"/>')
    expect(cache).toContain('refreshOnLoad="1"')
    const records = plan.replaced.get('xl/pivotCache/pivotCacheRecords1.xml')!
    expect(records).toContain('count="0"')
    expect(records).not.toContain('<r>')

    // Round-trip: the rewritten parts parse back and recompute correctly.
    const definition = parsePivotDefinition(table, cache)
    expect(definition.unsupported).toEqual([])
    const { data } = recomputePivotData(definition, [
      ['Region', 'Product', 'Amount'],
      ['East', 'A', 100],
      ['West', 'B', 50],
      ['East', 'B', 50],
    ])
    // Rows: A, B, Grand Total — single value column.
    expect(data).toEqual([[100], [100], [200]])
  })

  it('fails closed when the relayouted pivot loses its name or cacheId', async () => {
    const buf = await buildPivotFixture()
    const zip = await JSZip.loadAsync(buf)
    const tableXml = await zip.file('xl/pivotTables/pivotTable1.xml')!.async('string')
    zip.file('xl/pivotTables/pivotTable1.xml', tableXml.replace('cacheId="1"', ''))
    const broken = await zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
    await expect(
      planWithRefreshUpdate(broken, [
        {
          cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
          sheetName: 'Data',
          newOutputRef: 'F1:G4',
          relayout: {
            sourceSheetName: 'Data',
            sourceArea: { startRow: 0, startColumn: 0, endRow: 3, endColumn: 2 },
            location: { startRow: 0, startColumn: 5, endRow: 3, endColumn: 6 },
            name: 'PivotEdit',
            fieldNames: ['Region', 'Product', 'Amount'],
            rowFieldIndices: [1],
            rowItems: ['A', 'B'],
            values: [{ fieldIndex: 2, agg: 'sum' }],
          },
        },
      ]),
    ).rejects.toThrow(PivotExpandError)
  })
})
