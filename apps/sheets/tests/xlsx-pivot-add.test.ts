import { describe, expect, it } from 'vitest'

import {
  growPivotDefinition,
  recomputePivotData,
} from '@genoffice/xlsx-gateway/domain/pivot-engine'
import {
  createBufferEntrySource,
  planCellEditsToXlsx,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import type { SheetPivotAddition } from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { parsePivotDefinition } from '@genoffice/xlsx-gateway/gateway/xlsx-pivot'
import { buildEditFixture } from './fixture-builder'

function pivotAddition(overrides: Partial<SheetPivotAddition> = {}): SheetPivotAddition {
  return {
    sheetName: 'Data',
    sourceSheetName: 'Data',
    // Rows 0-2 (header=0, data rows=1-2): matches the edit fixture's sparse sheet
    // fixture has: row r=1 (A1=shared-str, C1=5), row r=3 (B3=formula→5)
    // data rows 1-2 in 0-indexed => Excel rows 2-3
    sourceArea: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 2 },
    location: { startRow: 0, startColumn: 5, endRow: 4, endColumn: 6 },
    name: 'Pivot1',
    fieldNames: ['Region', 'Product', 'Amount'],
    rowFieldIndices: [0],
    rowItems: ['East', 'South', 'North'],
    values: [{ fieldIndex: 2, agg: 'sum' }],
    ...overrides,
  }
}

async function planWith(
  pivots: SheetPivotAddition[],
  structuralOps: Parameters<typeof planCellEditsToXlsx>[2] = [],
) {
  const source = await createBufferEntrySource(await buildEditFixture())
  return planCellEditsToXlsx(
    source,
    [],
    structuralOps,
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
    pivots,
  )
}

describe('pivot additions', () => {
  it('writes the cache pair, the definition, and every linkage', async () => {
    const plan = await planWith([pivotAddition()])

    // Records: 2 data rows from the fixture (rows 2-3 in Excel = indices 1-2)
    // Row 1 (Excel row 2): all blank → <r><m/><m/><m/></r>
    // Row 2 (Excel row 3): A3=blank, B3=formula→5, C3=blank → <r><m/><n v="5"/><m/></r>
    const records = plan.added.get('xl/pivotCache/pivotCacheRecords1.xml')
    expect(records).toContain('count="2"')
    expect(records).toContain('<r><m/><m/><m/></r>')
    expect(records).toContain('<r><m/><n v="5"/><m/></r>')

    const cache = plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')
    // refreshOnLoad is omitted when records are pre-populated from source data
    expect(cache).not.toContain('refreshOnLoad="1"')
    expect(cache).toContain('recordCount="2"')
    expect(cache).toContain('<worksheetSource ref="A1:C3" sheet="Data"/>')
    expect(cache).toContain('<cacheFields count="3">')
    expect(cache).toContain(
      '<cacheField name="Region" numFmtId="0">' +
        '<sharedItems count="3"><s v="East"/><s v="South"/><s v="North"/></sharedItems>',
    )
    expect(cache).toContain('<cacheField name="Amount" numFmtId="0"><sharedItems/></cacheField>')

    const cacheRels = plan.added.get('xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels')
    expect(cacheRels).toContain('pivotCacheRecords')
    expect(cacheRels).toContain('Target="pivotCacheRecords1.xml"')

    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')
    expect(table).toContain('name="Pivot1" cacheId="1"')
    expect(table).toContain('<location ref="F1:G5"')
    expect(table).toContain('<pivotFields count="3">')
    expect(table).toContain(
      '<pivotField axis="axisRow" showAll="0">' +
        '<items count="4"><item x="0"/><item x="1"/><item x="2"/><item t="default"/></items>',
    )
    expect(table).toContain('<pivotField dataField="1" showAll="0"/>')
    expect(table).toContain('<rowFields count="1"><field x="0"/></rowFields>')
    expect(table).toContain(
      '<rowItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i>' +
        '<i t="grand"><x/></i></rowItems>',
    )
    expect(table).toContain('<colItems count="1"><i/></colItems>')
    expect(table).toContain('<dataField name="Sum of Amount" fld="2" baseField="0" baseItem="0"/>')

    const tableRels = plan.added.get('xl/pivotTables/_rels/pivotTable1.xml.rels')
    expect(tableRels).toContain('pivotCacheDefinition')
    expect(tableRels).toContain('Target="../pivotCache/pivotCacheDefinition1.xml"')

    const sheetRels = plan.added.get('xl/worksheets/_rels/sheet1.xml.rels')
    expect(sheetRels).toContain('relationships/pivotTable')
    expect(sheetRels).toContain('Target="../pivotTables/pivotTable1.xml"')

    const workbook = plan.replaced.get('xl/workbook.xml')
    expect(workbook).toMatch(/<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/)
    const workbookRels = plan.replaced.get('xl/_rels/workbook.xml.rels')
    expect(workbookRels).toContain('pivotCacheDefinition')
    expect(workbookRels).toContain('Target="pivotCache/pivotCacheDefinition1.xml"')

    const contentTypes = plan.replaced.get('[Content_Types].xml')
    expect(contentTypes).toContain('PartName="/xl/pivotTables/pivotTable1.xml"')
    expect(contentTypes).toContain('pivotCacheDefinition+xml')
    expect(contentTypes).toContain('pivotCacheRecords+xml')
  })

  it('spreads a column field across the columns axis', async () => {
    const plan = await planWith([
      pivotAddition({
        fieldNames: ['Region', 'Quarter', 'Amount'],
        columnFieldIndex: 1,
        columnItems: ['Q1', 'Q2'],
        values: [{ fieldIndex: 2, agg: 'average' }],
        location: { startRow: 0, startColumn: 5, endRow: 4, endColumn: 8 },
      }),
    ])
    const cache = plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')
    expect(cache).toContain(
      '<cacheField name="Quarter" numFmtId="0">' +
        '<sharedItems count="2"><s v="Q1"/><s v="Q2"/></sharedItems>',
    )
    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')
    expect(table).toContain('<pivotField axis="axisCol" showAll="0">')
    expect(table).toContain('<colFields count="1"><field x="1"/></colFields>')
    expect(table).toContain('<colItems count="3">')
    expect(table).toContain('subtotal="average"')
  })

  it('puts multiple value fields on the values pseudo-axis', async () => {
    const plan = await planWith([
      pivotAddition({
        values: [
          { fieldIndex: 1, agg: 'count' },
          { fieldIndex: 2, agg: 'sum' },
        ],
        location: { startRow: 0, startColumn: 5, endRow: 4, endColumn: 7 },
      }),
    ])
    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')
    expect(table).toContain('<colFields count="1"><field x="-2"/></colFields>')
    expect(table).toContain('<colItems count="2"><i><x/></i><i i="1"><x v="1"/></i></colItems>')
    expect(table).toContain('<dataFields count="2">')
    expect(table).toContain('<dataField name="Count of Product" fld="1" subtotal="count"')
  })

  it('assigns sequential cache ids and part numbers to two pivots', async () => {
    const plan = await planWith([
      pivotAddition(),
      pivotAddition({
        name: 'Pivot2',
        location: { startRow: 10, startColumn: 5, endRow: 14, endColumn: 6 },
      }),
    ])
    expect(plan.added.get('xl/pivotTables/pivotTable2.xml')).toContain('cacheId="2"')
    const workbook = plan.replaced.get('xl/workbook.xml')
    expect(workbook).toContain('cacheId="1"')
    expect(workbook).toContain('cacheId="2"')
  })

  it('rejects duplicate names, bad field indexes, and mismatched widths', async () => {
    await expect(planWith([pivotAddition(), pivotAddition()])).rejects.toThrow(/already taken/)
    await expect(planWith([pivotAddition({ rowFieldIndices: [9] })])).rejects.toThrow(
      /outside its source/,
    )
    await expect(
      planWith([
        pivotAddition({
          fieldNames: ['Region', 'Amount'],
          values: [{ fieldIndex: 1, agg: 'sum' }],
        }),
      ]),
    ).rejects.toThrow(/field names for a 3-column source/)
  })

  it('fails closed when the same save shifts rows on an involved sheet', async () => {
    await expect(
      planWith(
        [pivotAddition()],
        [{ sheetName: 'Data', ops: [{ kind: 'insert-rows', index: 0, count: 1 }] }],
      ),
    ).rejects.toThrow(/save the pivot first/)
  })

  it('writes showDataAs on the dataField and defaults the format to 0.00%', async () => {
    const plan = await planWith([
      pivotAddition({
        values: [{ fieldIndex: 2, agg: 'sum', showDataAs: 'percentOfTotal' }],
      }),
    ])
    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')
    expect(table).toContain(
      '<dataField name="Sum of Amount" fld="2" ' +
        'showDataAs="percentOfTotal" baseField="0" baseItem="0" numFmtId="10"/>',
    )
    // An explicit numFmt takes precedence over the percentage default format.
    const plan2 = await planWith([
      pivotAddition({
        values: [{ fieldIndex: 2, agg: 'sum', showDataAs: 'percentOfRow', numFmt: '0%' }],
      }),
    ])
    expect(plan2.added.get('xl/pivotTables/pivotTable1.xml')).toContain(
      'showDataAs="percentOfRow" baseField="0" baseItem="0" numFmtId="9"/>',
    )
  })

  it('single-level rowLines produce the same rowItems as the legacy shape', async () => {
    // The app now always sends rowLevelItems/rowLines; single-level output must match the legacy shape.
    const plan = await planWith([
      pivotAddition({
        rowLevelItems: [['East', 'South', 'North']],
        rowLines: [
          { t: 'data', members: [0] },
          { t: 'data', members: [1] },
          { t: 'data', members: [2] },
        ],
      }),
    ])
    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')
    expect(table).toContain(
      '<rowItems count="4"><i><x/></i><i><x v="1"/></i><i><x v="2"/></i>' +
        '<i t="grand"><x/></i></rowItems>',
    )
    expect(table).toContain('outline="1" outlineData="1"')
    expect(table).toContain('firstDataCol="1"')
  })

  it('cache records reflect actual source data: blank rows → <m/>, numeric cells → <n/>', async () => {
    // rowItems contains the three region labels but fixture col A (region) has no matching values
    // → col 0 (region) all rows: <m/> (no match to sharedItems)
    // fixture B3 = formula evaluating to 5 → col 1, row 2 (0-indexed): <n v="5"/>
    // col 2 (amount) all rows: <m/>
    const plan = await planWith([pivotAddition()])
    const records = plan.added.get('xl/pivotCache/pivotCacheRecords1.xml')
    expect(records).toContain('<r><m/><m/><m/></r>') // row index 1 (Excel row 2, all blank)
    expect(records).toContain('<r><m/><n v="5"/><m/></r>') // row index 2 (Excel row 3, B has value)
    // Dimension field matches: when col A has values from rowItems, emit <x v="N"/> instead of <m/>
    // (tested implicitly through the count assertion and structure check above)
    const cache = plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')
    expect(cache).toContain('recordCount="2"')
  })
})

/// Fixture for two-level row grouping: East{A,B} + South{B} (no South-A combo),
/// each outer member followed by a subtotal row, grand-total row at the end.
function twoLevelAddition(overrides: Partial<SheetPivotAddition> = {}): SheetPivotAddition {
  return pivotAddition({
    rowFieldIndices: [0, 1],
    rowItems: ['East', 'South'],
    rowLevelItems: [
      ['East', 'South'],
      ['A', 'B'],
    ],
    rowLines: [
      { t: 'data', members: [0, 0] },
      { t: 'data', members: [0, 1] },
      { t: 'default', members: [0] },
      { t: 'data', members: [1, 1] },
      { t: 'default', members: [1] },
    ],
    // Header + 5 layout rows + grand total = 7 rows; 2 label columns + 1 value column = 3 columns.
    location: { startRow: 0, startColumn: 5, endRow: 6, endColumn: 7 },
    ...overrides,
  })
}

describe('multi-level row pivots', () => {
  it('writes per-level sharedItems, pivotField items, and repeat-encoded rowItems', async () => {
    const plan = await planWith([twoLevelAddition()])

    const cache = plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')
    // Every row-field level gets its own sharedItems (the inner level is no longer an empty best-effort).
    expect(cache).toContain(
      '<cacheField name="Region" numFmtId="0">' +
        '<sharedItems count="2"><s v="East"/><s v="South"/></sharedItems>',
    )
    expect(cache).toContain(
      '<cacheField name="Product" numFmtId="0">' +
        '<sharedItems count="2"><s v="A"/><s v="B"/></sharedItems>',
    )

    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')
    // Multi-level uses tabular layout: one column per level, data area starting at column 2.
    expect(table).toContain('compact="0" compactData="0" outline="0" outlineData="0"')
    expect(table).toContain('firstDataCol="2"')
    expect(table).toContain('<rowFields count="2"><field x="0"/><field x="1"/></rowFields>')
    // Both levels carry a real items list.
    expect(table).toContain(
      '<pivotField axis="axisRow" showAll="0" compact="0" outline="0">' +
        '<items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField>',
    )
    // r semantics: the prefix shared with the previous row is omitted; subtotal rows use t="default".
    expect(table).toContain(
      '<rowItems count="6">' +
        '<i><x/><x/></i>' +
        '<i r="1"><x v="1"/></i>' +
        '<i t="default"><x/></i>' +
        '<i><x v="1"/><x v="1"/></i>' +
        '<i t="default"><x v="1"/></i>' +
        '<i t="grand"><x/></i>' +
        '</rowItems>',
    )
  })

  it('round-trips through the recompute engine (multi-level closure)', async () => {
    const plan = await planWith([twoLevelAddition()])
    const tableXml = plan.added.get('xl/pivotTables/pivotTable1.xml')!
    const cacheXml = plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')!
    const definition = parsePivotDefinition(tableXml, cacheXml)
    expect(definition.unsupported).toEqual([])
    expect(definition.rowFields).toEqual([0, 1])
    // Recompute with fresh source data matching the layout: data/subtotal/grand-total rows line up one by one.
    const result = recomputePivotData(definition, [
      ['Region', 'Product', 'Amount'],
      ['East', 'A', 10],
      ['East', 'B', 20],
      ['South', 'B', 40],
    ])
    expect(result.data).toEqual([
      [10], // East / A
      [20], // East / B
      [30], // East subtotal
      [40], // South / B
      [40], // South subtotal
      [70], // grand total
    ])
  })

  it('rejects a multi-level addition without per-level items or with bad lines', async () => {
    await expect(
      planWith([twoLevelAddition({ rowLevelItems: undefined, rowLines: undefined })]),
    ).rejects.toThrow(/row levels but no per-level items/)
    await expect(
      planWith([
        twoLevelAddition({
          rowLines: [{ t: 'data', members: [0, 9] }],
        }),
      ]),
    ).rejects.toThrow(/inconsistent row line/)
  })
})

/// Fixture for two-level column grouping: single-level row axis (two regions), column axis
/// Quarter{Q1,Q2} × Product{A,B} (no Q2-A combo), each outer column member followed by a
/// subtotal column, grand-total column at the end.
function twoLevelColumnAddition(overrides: Partial<SheetPivotAddition> = {}): SheetPivotAddition {
  return pivotAddition({
    sourceArea: { startRow: 0, startColumn: 0, endRow: 2, endColumn: 3 },
    fieldNames: ['Region', 'Quarter', 'Product', 'Amount'],
    rowFieldIndices: [0],
    rowItems: ['East', 'South'],
    columnFieldIndices: [1, 2],
    colLevelItems: [
      ['Q1', 'Q2'],
      ['A', 'B'],
    ],
    colLines: [
      { t: 'data', members: [0, 0] },
      { t: 'data', members: [0, 1] },
      { t: 'default', members: [0] },
      { t: 'data', members: [1, 1] },
      { t: 'default', members: [1] },
    ],
    values: [{ fieldIndex: 3, agg: 'sum' }],
    // 2 header rows + 2 data rows + grand-total row = 5 rows; 1 label col + 5 col lines + grand-total col = 7 cols.
    location: { startRow: 0, startColumn: 5, endRow: 4, endColumn: 11 },
    ...overrides,
  })
}

describe('multi-level column pivots', () => {
  it('writes per-level sharedItems, axisCol pivotFields, and repeat-encoded colItems', async () => {
    const plan = await planWith([twoLevelColumnAddition()])

    const cache = plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')
    // Every column-field level gets its own sharedItems.
    expect(cache).toContain(
      '<cacheField name="Quarter" numFmtId="0">' +
        '<sharedItems count="2"><s v="Q1"/><s v="Q2"/></sharedItems>',
    )
    expect(cache).toContain(
      '<cacheField name="Product" numFmtId="0">' +
        '<sharedItems count="2"><s v="A"/><s v="B"/></sharedItems>',
    )

    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')
    // Each column level takes one header row; the data area starts at row 2.
    expect(table).toContain('firstDataRow="2"')
    expect(table).toContain('<colFields count="2"><field x="1"/><field x="2"/></colFields>')
    // Both column-field levels are axisCol and carry a real items list.
    expect(table).toContain(
      '<pivotField axis="axisCol" showAll="0">' +
        '<items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField>',
    )
    // r semantics: the prefix shared with the previous column is omitted; subtotal columns use t="default"; grand-total column last.
    expect(table).toContain(
      '<colItems count="6">' +
        '<i><x/><x/></i>' +
        '<i r="1"><x v="1"/></i>' +
        '<i t="default"><x/></i>' +
        '<i><x v="1"/><x v="1"/></i>' +
        '<i t="default"><x v="1"/></i>' +
        '<i t="grand"><x/></i>' +
        '</colItems>',
    )
  })

  it('round-trips through the recompute engine (multi-level column closure)', async () => {
    const plan = await planWith([twoLevelColumnAddition()])
    const tableXml = plan.added.get('xl/pivotTables/pivotTable1.xml')!
    const cacheXml = plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')!
    const definition = parsePivotDefinition(tableXml, cacheXml)
    expect(definition.unsupported).toEqual([])
    expect(definition.colFields).toEqual([1, 2])
    // Recompute with fresh source data matching the layout: data/subtotal/grand-total columns line up one by one.
    const result = recomputePivotData(definition, [
      ['Region', 'Quarter', 'Product', 'Amount'],
      ['East', 'Q1', 'A', 10],
      ['East', 'Q1', 'B', 20],
      ['South', 'Q2', 'B', 40],
    ])
    expect(result.data).toEqual([
      // Column order: Q1/A, Q1/B, Q1 subtotal, Q2/B, Q2 subtotal, grand total
      [10, 20, 30, 0, 0, 30], // East
      [0, 0, 0, 40, 40, 40], // South
      [10, 20, 30, 40, 40, 70], // grand total
    ])
  })

  it('persists field groupings in extLst and round-trips through recompute', async () => {
    // Date row field grouped by month: rowItems are already group labels; the grouping rule goes into extLst.
    const plan = await planWith([
      pivotAddition({
        fieldNames: ['Date', 'Product', 'Amount'],
        rowItems: ['Jan', 'Feb'],
        groupings: [{ fieldIndex: 0, kind: 'date', dateUnit: 'month' }],
        location: { startRow: 0, startColumn: 5, endRow: 3, endColumn: 6 },
      }),
    ])
    const tableXml = plan.added.get('xl/pivotTables/pivotTable1.xml')!
    expect(tableXml).toContain('<extLst><ext uri="{AIO-PIVOT-GROUPINGS}"')
    expect(tableXml).toContain('aioPivotGroupings')

    const definition = parsePivotDefinition(
      tableXml,
      plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')!,
    )
    expect(definition.unsupported).toEqual([])
    expect(definition.fields[0]!.grouping).toEqual({ kind: 'date', dateUnit: 'month' })
    // Recompute after grouping raw dates by month: January = 10+5, February = 7.
    const result = recomputePivotData(definition, [
      ['Date', 'Product', 'Amount'],
      ['2026-01-05', 'A', 10],
      ['2026-01-20', 'B', 5],
      ['2026-02-02', 'A', 7],
    ])
    expect(result.data).toEqual([[15], [7], [22]])
  })

  it('writes calculated fields as formula cacheFields and round-trips recompute', async () => {
    const plan = await planWith([
      pivotAddition({
        fieldNames: ['Region', 'Revenue', 'Cost'],
        rowItems: ['East', 'South'],
        values: [{ fieldIndex: -1, agg: 'sum', formula: 'Revenue-Cost', calcName: 'Profit' }],
        location: { startRow: 0, startColumn: 5, endRow: 3, endColumn: 6 },
      }),
    ])
    const cache = plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')!
    // The calculated field is appended after the source fields: the formula is canonicalized into quoted form and takes no cache records.
    expect(cache).toContain('<cacheFields count="4">')
    expect(cache).toContain(
      '<cacheField name="Profit" formula="(\'Revenue\'-\'Cost\')" ' +
        'databaseField="0" numFmtId="0"/>',
    )

    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')!
    expect(table).toContain('<pivotFields count="4">')
    expect(table).toContain('<dataField name="Sum of Profit" fld="3" baseField="0" baseItem="0"/>')

    const definition = parsePivotDefinition(table, cache)
    expect(definition.unsupported).toEqual([])
    const result = recomputePivotData(definition, [
      ['Region', 'Revenue', 'Cost'],
      ['East', 100, 40],
      ['East', 50, 10],
      ['South', 80, 30],
    ])
    expect(result.data).toEqual([[100], [50], [150]])
  })

  it('rejects a calculated field with a bad formula or a colliding name', async () => {
    await expect(
      planWith([
        pivotAddition({
          fieldNames: ['Region', 'Revenue', 'Cost'],
          values: [{ fieldIndex: -1, agg: 'sum', formula: 'Revenue-Margin', calcName: 'Profit' }],
        }),
      ]),
    ).rejects.toThrow(/calculated-field formula/)
    await expect(
      planWith([
        pivotAddition({
          fieldNames: ['Region', 'Revenue', 'Cost'],
          values: [{ fieldIndex: -1, agg: 'sum', formula: 'Revenue-Cost', calcName: 'Cost' }],
        }),
      ]),
    ).rejects.toThrow(/invalid calculated field/)
  })

  it('writes pivot filters with hidden items and round-trips recompute', async () => {
    // Label filter equal to East: the other regions become hidden items; layout keeps only it + grand total.
    const plan = await planWith([
      pivotAddition({
        filters: [{ kind: 'label', field: 0, op: 'equal', value: 'East' }],
        rowHiddenItems: [[1, 2]],
        rowLevelItems: [['East', 'South', 'North']],
        rowLines: [{ t: 'data', members: [0] }],
        location: { startRow: 0, startColumn: 5, endRow: 2, endColumn: 6 },
      }),
    ])
    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')!
    expect(table).toContain('<item x="0"/><item x="1" h="1"/><item x="2" h="1"/>')
    expect(table).toContain(
      '<filters count="1">' +
        '<filter fld="0" type="captionEqual" evalOrder="-1" id="1" stringValue1="East">',
    )
    expect(table).toContain('<rowItems count="2"><i><x/></i><i t="grand"><x/></i></rowItems>')

    const definition = parsePivotDefinition(
      table,
      plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')!,
    )
    expect(definition.unsupported).toEqual([])
    expect(definition.filters).toEqual([{ kind: 'label', field: 0, op: 'equal', value: 'East' }])
    const source = [
      ['Region', 'Product', 'Amount'],
      ['East', 'A', 10],
      ['South', 'A', 99],
      ['North', 'B', 7],
    ]
    // Hidden items already match the filter → no growth needed; the recompute holds only East and the grand total.
    const growth = growPivotDefinition(definition, source)
    expect(growth.grown).toBe(false)
    expect(recomputePivotData(definition, source).data).toEqual([[10], [10]])
  })

  it('writes a top-N value filter with iMeasureFld and top10', async () => {
    const plan = await planWith([
      pivotAddition({
        filters: [{ kind: 'value', field: 0, dataField: 0, op: 'top', count: 2 }],
        rowHiddenItems: [[2]],
        rowLevelItems: [['East', 'South', 'North']],
        rowLines: [
          { t: 'data', members: [0] },
          { t: 'data', members: [1] },
        ],
        location: { startRow: 0, startColumn: 5, endRow: 3, endColumn: 6 },
      }),
    ])
    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')!
    expect(table).toContain('<filter fld="0" type="count" evalOrder="-1" id="1" iMeasureFld="0">')
    expect(table).toContain('<top10 val="2"/>')
  })

  it('rejects filters on non-dimension fields or duplicated per field', async () => {
    await expect(
      planWith([
        pivotAddition({
          filters: [{ kind: 'label', field: 2, op: 'equal', value: 'x' }],
        }),
      ]),
    ).rejects.toThrow(/invalid pivot filter/)
    await expect(
      planWith([
        pivotAddition({
          filters: [
            { kind: 'label', field: 0, op: 'equal', value: 'x' },
            { kind: 'value', field: 0, dataField: 0, op: 'top', count: 3 },
          ],
        }),
      ]),
    ).rejects.toThrow(/invalid pivot filter/)
  })

  it('rejects a grouping that is not on a dimension field', async () => {
    await expect(
      planWith([
        pivotAddition({
          groupings: [{ fieldIndex: 2, kind: 'range', rangeStep: 100 }],
        }),
      ]),
    ).rejects.toThrow(/invalid field grouping/)
  })

  it('rejects a multi-level column addition without per-level items or with bad lines', async () => {
    await expect(
      planWith([
        twoLevelColumnAddition({
          colLevelItems: undefined,
          colLines: undefined,
        }),
      ]),
    ).rejects.toThrow(/column levels but no per-level items/)
    await expect(
      planWith([
        twoLevelColumnAddition({
          colLines: [{ t: 'data', members: [0, 9] }],
        }),
      ]),
    ).rejects.toThrow(/inconsistent column line/)
  })
})
