import JSZip from 'jszip'
import { describe, expect, it } from 'vitest'

import { recomputePivotData } from '@genoffice/xlsx-gateway/domain/pivot-engine'
import {
  buildPivotLayout,
  PivotLayoutError,
  type PivotLayoutSpec,
} from '@genoffice/xlsx-gateway/domain/pivot-layout'
import {
  createBufferEntrySource,
  planCellEditsToXlsx,
  type MutationPlan,
  type PivotRefreshUpdate,
  type SheetPivotAddition,
} from '@genoffice/xlsx-gateway/gateway/xlsx-gateway'
import { parsePivotDefinition } from '@genoffice/xlsx-gateway/gateway/xlsx-pivot'

const MAIN_NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

const GRID: (string | number)[][] = [
  ['Region', 'Product', 'Amount'],
  ['East', 'A', 100],
  ['West', 'B', 50],
  ['East', 'B', 50],
  ['West', 'A', 25],
]

function cell(ref: string, value: string | number): string {
  return typeof value === 'number'
    ? `<c r="${ref}"><v>${value}</v></c>`
    : `<c r="${ref}" t="inlineStr"><is><t>${value}</t></is></c>`
}

async function fixture(): Promise<Buffer> {
  const zip = new JSZip()
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>`,
  )
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_NS}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/workbook.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="${MAIN_NS}" xmlns:r="${REL_NS}">
  <sheets><sheet name="Data" sheetId="1" r:id="rId1"/></sheets>
</workbook>`,
  )
  zip.file(
    'xl/_rels/workbook.xml.rels',
    `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="${REL_NS}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="${REL_NS}/styles" Target="styles.xml"/>
</Relationships>`,
  )
  zip.file(
    'xl/styles.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<styleSheet xmlns="${MAIN_NS}">
  <fonts count="1"><font/></fonts><fills count="1"><fill/></fills><borders count="1"><border/></borders>
  <cellStyleXfs count="1"><xf/></cellStyleXfs><cellXfs count="1"><xf/></cellXfs>
</styleSheet>`,
  )
  const rows = GRID.map(
    (line, r) =>
      `<row r="${r + 1}">${line.map((value, c) => cell(`${'ABC'[c]}${r + 1}`, value)).join('')}</row>`,
  ).join('')
  zip.file(
    'xl/worksheets/sheet1.xml',
    `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="${MAIN_NS}"><sheetData>${rows}</sheetData></worksheet>`,
  )
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function materialize(buf: Buffer, plan: MutationPlan): Promise<Buffer> {
  const zip = await JSZip.loadAsync(buf)
  for (const path of plan.removedEntries) zip.remove(path)
  for (const [path, xml] of plan.replaced) zip.file(path, xml)
  for (const [path, xml] of plan.added) zip.file(path, xml)
  return zip.generateAsync({ type: 'nodebuffer', compression: 'DEFLATE' })
}

async function planWith(
  buf: Buffer,
  pivots: SheetPivotAddition[],
  updates: PivotRefreshUpdate[] = [],
): Promise<MutationPlan> {
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
    pivots,
    [],
    updates,
  )
}

/// The app's addition for a layout spec: body at F3 (two report-filter rows
/// above it when page fields exist, like Excel).
function additionFor(spec: PivotLayoutSpec): SheetPivotAddition {
  const layout = buildPivotLayout(GRID, spec)
  const pageRows = (layout.definition.pageFieldIndices?.length ?? 0) > 0 ? 2 : 0
  return {
    sheetName: 'Data',
    sourceSheetName: 'Data',
    sourceArea: { startRow: 0, startColumn: 0, endRow: 4, endColumn: 2 },
    location: {
      startRow: pageRows,
      startColumn: 5,
      endRow: pageRows + layout.height - 1,
      endColumn: 5 + layout.width - 1,
    },
    name: 'Pivot1',
    ...layout.definition,
  }
}

/// Numeric data area of a baked matrix (after the header rows, right of the
/// row-label columns).
function dataArea(spec: PivotLayoutSpec): (number | null)[][] {
  const layout = buildPivotLayout(GRID, spec)
  const rowLevels = Array.isArray(spec.rowFields) ? spec.rowFields.length : 1
  const colLevels =
    spec.columnField === undefined
      ? 0
      : Array.isArray(spec.columnField)
        ? spec.columnField.length
        : 1
  return layout.matrix
    .slice(Math.max(1, colLevels))
    .map((line) => line.slice(rowLevels).map((value) => (typeof value === 'number' ? value : null)))
}

const SUM: PivotLayoutSpec['values'] = [{ field: 'Amount', agg: 'sum' }]

describe('report filters in the layout builder', () => {
  it('a bare page field shows every row and records the member list', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: 'Region',
      pageFields: ['Product'],
      values: SUM,
    })
    expect(layout.definition.pageFieldIndices).toEqual([1])
    expect(layout.definition.pageLevelItems).toEqual([['A', 'B']])
    expect(layout.definition.pageItems).toEqual([null])
    expect(layout.matrix).toEqual([
      ['Region', 'Sum of Amount'],
      ['East', 150],
      ['West', 75],
      ['Grand Total', 225],
    ])
  })

  it('a selected item narrows every aggregate', () => {
    const layout = buildPivotLayout(GRID, {
      rowFields: 'Region',
      pageFields: [{ field: 'Product', item: 'B' }],
      values: SUM,
    })
    expect(layout.definition.pageItems).toEqual([1])
    expect(layout.matrix).toEqual([
      ['Region', 'Sum of Amount'],
      ['East', 50],
      ['West', 50],
      ['Grand Total', 100],
    ])
  })

  it('rejects an item that is not a member', () => {
    expect(() =>
      buildPivotLayout(GRID, {
        rowFields: 'Region',
        pageFields: [{ field: 'Product', item: 'Z' }],
        values: SUM,
      }),
    ).toThrow(PivotLayoutError)
  })
})

describe('report filters in the pivot parts', () => {
  const spec: PivotLayoutSpec = {
    rowFields: 'Region',
    pageFields: [{ field: 'Product', item: 'B' }],
    values: SUM,
  }

  it('writes rowPageCount, the selected pageField item, axisPage items, and indexed records', async () => {
    const plan = await planWith(await fixture(), [additionFor(spec)])
    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')!
    expect(table).toContain(
      '<location ref="F3:G6" firstHeaderRow="1" firstDataRow="1" firstDataCol="1" rowPageCount="1" colPageCount="1"/>',
    )
    expect(table).toContain(
      '<pageFields count="1"><pageField fld="1" item="1" hier="-1"/></pageFields>',
    )
    expect(table).toContain(
      '<pivotField axis="axisPage" showAll="0"><items count="3"><item x="0"/><item x="1"/><item t="default"/></items></pivotField>',
    )
    const cache = plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')!
    expect(cache).toContain(
      '<cacheField name="Product" numFmtId="0"><sharedItems count="2"><s v="A"/><s v="B"/></sharedItems></cacheField>',
    )
    const records = plan.added.get('xl/pivotCache/pivotCacheRecords1.xml')!
    expect(records).toContain('<r><x v="0"/><x v="0"/><n v="100"/></r>')
    expect(records).toContain('<r><x v="1"/><x v="1"/><n v="50"/></r>')
  })

  it('round-trips: the parsed page selection filters the recompute like the baked grid', async () => {
    const plan = await planWith(await fixture(), [additionFor(spec)])
    const definition = parsePivotDefinition(
      plan.added.get('xl/pivotTables/pivotTable1.xml')!,
      plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')!,
    )
    expect(definition.unsupported).toEqual([])
    expect(definition.pageFields).toEqual([{ field: 1, item: 1 }])
    expect(recomputePivotData(definition, GRID).data).toEqual(dataArea(spec))
  })

  it('a layout edit rewrites pageFields and the location through the modelled-content guard', async () => {
    const created = await materialize(
      await fixture(),
      await planWith(await fixture(), [additionFor(spec)]),
    )
    const { sheetName: _sheet, ...relayout } = additionFor({
      rowFields: 'Region',
      pageFields: ['Product'],
      values: SUM,
    })
    const plan = await planWith(
      created,
      [],
      [
        {
          cachePath: 'xl/pivotCache/pivotCacheDefinition1.xml',
          sheetName: 'Data',
          newOutputRef: 'F3:G6',
          relayout,
        },
      ],
    )
    const table = plan.replaced.get('xl/pivotTables/pivotTable1.xml')!
    expect(table).toContain('<pageFields count="1"><pageField fld="1" hier="-1"/></pageFields>')
    expect(table).toContain('rowPageCount="1" colPageCount="1"')
    const definition = parsePivotDefinition(
      table,
      plan.replaced.get('xl/pivotCache/pivotCacheDefinition1.xml')!,
    )
    expect(definition.pageFields).toEqual([{ field: 1, item: null }])
    expect(recomputePivotData(definition, GRID).data).toEqual([[150], [75], [225]])
  })
})

describe('value field modes the engine supports', () => {
  async function roundTrip(spec: PivotLayoutSpec): Promise<{
    table: string
    engine: readonly (readonly (number | null)[])[]
    baked: (number | null)[][]
  }> {
    const plan = await planWith(await fixture(), [additionFor(spec)])
    const table = plan.added.get('xl/pivotTables/pivotTable1.xml')!
    const definition = parsePivotDefinition(
      table,
      plan.added.get('xl/pivotCache/pivotCacheDefinition1.xml')!,
    )
    expect(definition.unsupported).toEqual([])
    return { table, engine: recomputePivotData(definition, GRID).data, baked: dataArea(spec) }
  }

  it('product and count-numbers aggregate and persist as subtotal values', async () => {
    const { table, engine, baked } = await roundTrip({
      rowFields: 'Region',
      values: [
        { field: 'Amount', agg: 'product' },
        { field: 'Amount', agg: 'countNums' },
      ],
    })
    expect(table).toContain('<dataField name="Product of Amount" fld="2" subtotal="product"')
    expect(table).toContain('<dataField name="Count of Amount" fld="2" subtotal="countNums"')
    expect(baked).toEqual([
      [5000, 2],
      [1250, 2],
      [6_250_000, 4],
    ])
    expect(engine).toEqual(baked)
  })

  it('% of parent row total divides by the next level up', async () => {
    const { engine, baked } = await roundTrip({
      rowFields: ['Region', 'Product'],
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'percentOfParentRow' }],
    })
    expect(baked).toEqual([
      [100 / 150],
      [50 / 150],
      [150 / 225],
      [25 / 75],
      [50 / 75],
      [75 / 225],
      [1],
    ])
    expect(engine).toEqual(baked)
  })

  it('% of parent column total and index agree between baked grid and recompute', async () => {
    const parent = await roundTrip({
      rowFields: 'Region',
      columnField: 'Product',
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'percentOfParentCol' }],
    })
    expect(parent.baked[0]).toEqual([100 / 150, 50 / 150, 1])
    expect(parent.engine).toEqual(parent.baked)
    const index = await roundTrip({
      rowFields: 'Region',
      columnField: 'Product',
      values: [{ field: 'Amount', agg: 'sum', showDataAs: 'index' }],
    })
    expect(index.table).toContain('showDataAs="index" baseField="0" baseItem="0"/>')
    expect(index.baked[0]).toEqual([(100 * 225) / (150 * 125), (50 * 225) / (150 * 100), 1])
    expect(index.engine).toEqual(index.baked)
  })

  it('custom names caption the data field and duplicates get a numeric suffix', async () => {
    const { table } = await roundTrip({
      rowFields: 'Region',
      values: [
        { field: 'Amount', agg: 'sum', name: 'Revenue' },
        { field: 'Amount', agg: 'count' },
        { field: 'Amount', agg: 'countNums' },
      ],
    })
    expect(table).toContain('<dataField name="Revenue" fld="2" baseField="0"')
    expect(table).toContain('<dataField name="Count of Amount" fld="2" subtotal="count"')
    expect(table).toContain('<dataField name="Count of Amount2" fld="2" subtotal="countNums"')
  })
})
